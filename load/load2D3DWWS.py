import os
import time
import pandas as pd
import shutil
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urlparse
import sys

# ---------------- CONFIG ----------------
excel_input = r"C:\Users\tanapat\Downloads\1_WWS_model id to get 2D-3D_20Aug25_updated style no.xlsx"
base_folder = r"D:\WWS\2D&3D"
log_file = r"D:\WWS\download_log.txt"

# --- redirect print ไป log file ---
class Logger(object):
    def __init__(self, logfile_path):
        self.terminal = sys.stdout
        self.log = open(logfile_path, "a", encoding="utf-8")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()  # เขียนทันที

    def flush(self):
        self.terminal.flush()
        self.log.flush()

sys.stdout = Logger(log_file)

# อ่านข้อมูลจากไฟล์ Excel
df = pd.read_excel(excel_input)
search_list = df['Style No.'].dropna().astype(str).tolist()
id_list = df['id'].dropna().astype(str).tolist()

# ---  function สำหรับดาวน์โหลดและย้ายไฟล์ ---
def download_and_move_file(driver_instance, url, target_folder, file_context="file", custom_name=None):
    if custom_name:
        print(f"⌛ กำลังดาวน์โหลด {file_context} จาก {url} (ตั้งชื่อไฟล์เป็น '{custom_name}')")
    else:
        print(f"⌛ กำลังดาวน์โหลด {file_context} จาก {url}")
    
    old_files = set(os.listdir(base_folder))

    try:
        driver_instance.get(url)
        time.sleep(3)

        new_files = set(os.listdir(base_folder))
        downloaded_files = list(new_files - old_files)

        if downloaded_files:
            latest_file = max(downloaded_files, key=lambda f: os.path.getmtime(os.path.join(base_folder, f)))
            source_path = os.path.join(base_folder, latest_file)

            if custom_name:
                _, file_extension = os.path.splitext(latest_file)
                safe_name = "".join(c for c in custom_name if c not in r'\/:*?"<>|')
                new_filename = f"{safe_name}{file_extension}"
            else:
                filename_from_url = os.path.basename(urlparse(url).path)
                if filename_from_url and "." in filename_from_url:
                    base_filename, file_extension = os.path.splitext(filename_from_url)
                else:
                    base_filename, file_extension = os.path.splitext(latest_file)
                new_filename = f"{base_filename}{file_extension}"

            os.makedirs(target_folder, exist_ok=True)
            save_path = os.path.join(target_folder, new_filename)
            shutil.move(source_path, save_path)
            print(f"✅ บันทึก {file_context} '{new_filename}' เสร็จแล้วที่ {target_folder}")
            return True
        else:
            print(f"⚠️ ไม่พบไฟล์ {file_context} ที่ดาวน์โหลดจาก {url}")
            return False

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการดาวน์โหลด {file_context} จาก {url}: {e}")
        return False

# --- ตั้งค่า Chrome Driver ---
options = uc.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-notifications")
options.add_argument("--disable-extensions")

prefs = {
    "plugins.always_open_pdf_externally": True,
    "download.default_directory": base_folder
}
options.add_experimental_option("prefs", prefs)

driver = uc.Chrome(options=options)
wait = WebDriverWait(driver, 20)

try:
    for idx, vid_search in enumerate(search_list):
        max_search_retries = 2
        current_search_retry = 0
        search_successful = False

        while not search_successful and current_search_retry < max_search_retries:
            try:
                vid_folder = id_list[idx]
                id_folder = os.path.join(base_folder, vid_folder)

                datasheet_folder = os.path.join(id_folder, "Datasheet")
                folder_2d = os.path.join(id_folder, "2D")
                folder_3d = os.path.join(id_folder, "3D")

                driver.get("https://www.waterworks.com/us_en/")
                time.sleep(3)

                search_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", search_box)
                search_box.click()
                search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]/input")))
                search_input.clear()
                search_input.send_keys(vid_search)
                search_input.send_keys(Keys.RETURN)
                time.sleep(2)

                try:
                    first_item = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//ol/li[1]/div/div[1]"))
                    )
                    first_item.click()
                except:
                    pass
                time.sleep(2)
                
                li_list = []
                for i in range(7, 10):
                    try:
                        container = driver.find_element(By.XPATH,
                            f"/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[{i}]/div[1]/div[2]/div")
                        ul_element = container.find_element(By.XPATH, ".//ul")
                        li_list = ul_element.find_elements(By.TAG_NAME, "li")
                        if li_list:
                            print(f"ℹ️ {vid_search} (ลองครั้งที่ {current_search_retry+1}): พบรายการลิงก์ใน div[{i}]")
                            break
                    except:
                        continue
                
                if li_list:
                    search_successful = True
                else:
                    current_search_retry += 1
                    print(f"⚠️ {vid_search}: ไม่พบรายการลิงก์ใน div[7]-div[9] (ลองใหม่ครั้งที่ {current_search_retry})")
                    if current_search_retry == max_search_retries:
                        print(f"❌ {vid_search}: ไม่สามารถหารายการลิงก์ได้หลังจาก {max_search_retries} ครั้ง, ข้ามไปสินค้ารายการถัดไป")
                        break

            except Exception as e:
                current_search_retry += 1
                print(f"❌ Error {vid_search} (ลองใหม่ครั้งที่ {current_search_retry}): {e}")
                if current_search_retry == max_search_retries:
                    print(f"❌ {vid_search}: เกิดข้อผิดพลาดซ้ำๆ หลังจาก {max_search_retries} ครั้ง, ข้ามไปสินค้ารายการถัดไป")
                    break
                try:
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                except:
                    pass
                time.sleep(2)

        if not search_successful:
            continue

        # --- ดาวน์โหลดไฟล์ PDF / 2D / 3D ---
        for li in li_list:
            li_text = li.text.strip()
            try:
                a_tags = li.find_elements(By.TAG_NAME, "a")
                if a_tags:
                    for a_tag in a_tags:
                        href = a_tag.get_attribute("href")
                        a_text = driver.execute_script("return arguments[0].innerText;", a_tag).strip()
                        if not a_text:
                            a_text = li_text
                        
                        href_lower = href.lower()
                        a_text_lower = a_text.lower()

                        if "tearsheet" in a_text_lower or href_lower.endswith(".pdf"):
                            download_and_move_file(driver, href, datasheet_folder, "PDF Datasheet", custom_name=a_text)
                        elif "cad block" in a_text_lower:
                            download_and_move_file(driver, href, folder_2d, "CAD BLOCK (2D)")
                        elif any(ext in href_lower for ext in [".dwg", ".obj", ".3ds", ".stl", ".max", ".skp", ".fbx", ".rvt"]):
                            download_and_move_file(driver, href, folder_3d, "3D Model")
            except Exception as e:
                print(f"⚠️ {vid_search}: เกิดข้อผิดพลาดในการอ่านลิงก์ใน li '{li_text}': {e}")
        
        print(f"✅ {vid_search}: ตรวจสอบลิงก์และจัดการการดาวน์โหลดเสร็จสมบูรณ์")

        # --- ปิดแท็บที่เกิน ---
        try:
            while len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[-1])
                driver.close()
            driver.switch_to.window(driver.window_handles[0])
        except:
            pass

finally:
    try:
        driver.quit()
    except:
        pass
