import os
import time
import pandas as pd
import shutil
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urlparse # ใช้สำหรับแยกชื่อไฟล์จาก URL

# ---------------- CONFIG ----------------
excel_input = r"C:\Users\phunk\Downloads\WWS_model id to get 2D-3D_20Aug25.xlsx"
base_folder = r"C:\Users\phunk\OneDrive\Desktop\WWS\2D&3D"

# อ่านข้อมูลจากไฟล์ Excel
df = pd.read_excel(excel_input)
search_list = df['style'].dropna().astype(str).tolist() # รายการ 'style' สำหรับค้นหา
id_list = df['id'].dropna().astype(str).tolist()       # รายการ 'id' สำหรับตั้งชื่อโฟลเดอร์

# --- Helper function สำหรับดาวน์โหลดและย้ายไฟล์ ---
def download_and_move_file(driver_instance, url, target_folder, product_id, file_context="file"):
    """
    ดาวน์โหลดไฟล์จาก URL ที่กำหนดและย้ายไปยังโฟลเดอร์เป้าหมาย
    จะตรวจสอบไฟล์ที่ดาวน์โหลดล่าสุดใน base_folder
    """
    print(f"⌛ กำลังดาวน์โหลด {file_context} จาก {url}")
    old_files = set(os.listdir(base_folder)) # รายชื่อไฟล์ใน base_folder ก่อนเริ่มดาวน์โหลด

    try:
        # เข้าถึง URL เพื่อเริ่มการดาวน์โหลด
        driver_instance.get(url)
        time.sleep(5) # ให้เวลาเพียงพอสำหรับการดาวน์โหลด (อาจต้องปรับค่านี้)

        new_files = set(os.listdir(base_folder))
        downloaded_files = list(new_files - old_files) # หาไฟล์ที่เพิ่มเข้ามาใหม่

        if downloaded_files:
            # เลือกไฟล์ที่ดาวน์โหลดล่าสุด (สมมติว่าเป็นไฟล์ที่เราต้องการ)
            latest_file = max(downloaded_files, key=lambda f: os.path.getmtime(os.path.join(base_folder, f)))
            source_path = os.path.join(base_folder, latest_file)

            # พยายามดึงชื่อไฟล์และนามสกุลจาก URL ก่อน ถ้าไม่ได้ ให้ใช้ชื่อไฟล์ที่ดาวน์โหลดมา
            filename_from_url = os.path.basename(urlparse(url).path)
            if filename_from_url and "." in filename_from_url: # ตรวจสอบให้แน่ใจว่ามีนามสกุลไฟล์
                base_filename, file_extension = os.path.splitext(filename_from_url)
            else:
                base_filename, file_extension = os.path.splitext(latest_file) # ใช้ชื่อไฟล์ที่ดาวน์โหลดจริงเป็นตัวสำรอง

            # ตั้งชื่อไฟล์ใหม่เพื่อหลีกเลี่ยงการทับซ้อนและให้มี product_id นำหน้า
            new_filename = f"{product_id}_{base_filename}{file_extension}"
            save_path = os.path.join(target_folder, new_filename)

            # จัดการกรณีที่ชื่อไฟล์ซ้ำในโฟลเดอร์เป้าหมาย
            counter = 1
            while os.path.exists(save_path):
                new_filename = f"{product_id}_{base_filename}_{counter}{file_extension}"
                save_path = os.path.join(target_folder, new_filename)
                counter += 1

            shutil.move(source_path, save_path) # ย้ายไฟล์ไปยังโฟลเดอร์ที่ถูกต้อง
            print(f"✅ บันทึก {file_context} '{new_filename}' เสร็จแล้วที่ {target_folder}")
            return True
        else:
            print(f"⚠️ ไม่พบไฟล์ {file_context} ที่ดาวน์โหลดจาก {url}")
            return False

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการดาวน์โหลด {file_context} จาก {url}: {e}")
        return False


# --- ตั้งค่า Chrome Driver (ครั้งเดียวสำหรับทั้งสคริปต์) ---
options = uc.ChromeOptions()
options.add_argument("--start-maximized")         # เปิดหน้าต่างแบบเต็มจอ
options.add_argument("--disable-popup-blocking")  # ปิดการบล็อกป๊อปอัพ
options.add_argument("--disable-notifications")   # ปิดการแจ้งเตือน
options.add_argument("--disable-extensions")      # ปิดส่วนขยาย

# ตั้งค่าเบราว์เซอร์ให้ดาวน์โหลดไฟล์ PDF และไฟล์อื่นๆ โดยตรง
# "plugins.always_open_pdf_externally": True ทำให้ PDF ถูกดาวน์โหลดแทนที่จะเปิดในเบราว์เซอร์
# "download.default_directory": base_folder กำหนดโฟลเดอร์เริ่มต้นสำหรับการดาวน์โหลด
prefs = {
    "plugins.always_open_pdf_externally": True,
    "download.default_directory": base_folder
}
options.add_experimental_option("prefs", prefs)

# สร้าง Chrome Driver
driver = uc.Chrome(options=options)
wait = WebDriverWait(driver, 20)

try:
    # ---------------- Main Loop: วนลูปตามรายการสินค้า ----------------
    for idx, vid_search in enumerate(search_list):
        max_search_retries = 3 # จำนวนครั้งสูงสุดที่จะลองใหม่
        current_search_retry = 0
        search_successful = False

        while not search_successful and current_search_retry < max_search_retries:
            try:
                vid_folder = id_list[idx]
                id_folder = os.path.join(base_folder, vid_folder)

                datasheet_folder = os.path.join(id_folder, "Datasheet")
                folder_2d = os.path.join(id_folder, "2D")
                folder_3d = os.path.join(id_folder, "3D")
                os.makedirs(datasheet_folder, exist_ok=True)
                os.makedirs(folder_2d, exist_ok=True)
                os.makedirs(folder_3d, exist_ok=True)

                driver.get("https://www.waterworks.com/us_en/")
                time.sleep(3)

                # --- ค้นหาสินค้า ---
                search_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", search_box) # เลื่อนหน้าจอไปยังช่องค้นหา
                search_box.click() # คลิกช่องค้นหา
                search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]/input")))
                search_input.clear() # ล้างข้อความเก่า (ถ้ามี)
                search_input.send_keys(vid_search) # พิมพ์รหัสสินค้า
                search_input.send_keys(Keys.RETURN) # กด Enter เพื่อค้นหา
                time.sleep(2) # รอผลการค้นหา

                # --- คลิกรายการแนะนำ (autocomplete) ---
                try:
                    # พยายามคลิกรายการแรกที่แสดงใน autocomplete
                    first_item = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//ol/li[1]/div/div[1]"))
                    )
                    first_item.click()
                except:
                    pass

                time.sleep(2)

                # --- หา container ที่มีรายการลิงก์ (ul/li) และวนลูป ---
                li_list = []
                
                # ลองหาใน div ตั้งแต่ div[7] ถึง div[10] เพื่อความทนทานของโค้ด
                for i in range(7, 11):
                    try:
                        container = driver.find_element(By.XPATH,
                            f"/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[{i}]/div[1]/div[2]/div")
                        ul_element = container.find_element(By.XPATH, ".//ul")
                        li_list = ul_element.find_elements(By.TAG_NAME, "li")
                        if li_list: # ถ้าเจอรายการ li ให้หยุดการวนลูป
                            print(f"ℹ️ {vid_search} (ลองใหม่ครั้งที่ {current_search_retry+1}/{max_search_retries}): พบรายการลิงก์ใน div[{i}]")
                            break
                    except:
                        continue # ถ้าไม่เจอใน div ปัจจุบัน ให้ลอง div ถัดไป
                
                if li_list: # ถ้า li_list ถูกพบสำเร็จ
                    search_successful = True
                else:
                    current_search_retry += 1
                    print(f"⚠️ {vid_search}: ไม่พบรายการลิงก์ใน div[7]-div[10] (กำลังลองใหม่ครั้งที่ {current_search_retry}/{max_search_retries})")
                    if current_search_retry == max_search_retries:
                         print(f"❌ {vid_search}: ไม่สามารถหารายการลิงก์ได้หลังจาก {max_search_retries} ครั้ง, ข้ามไปสินค้ารายการถัดไป")
                         break # ออกจากลูป while เพื่อไปสินค้ารายการถัดไป

            except Exception as e:
                current_search_retry += 1
                print(f"❌ Error {vid_search} (ลองใหม่ครั้งที่ {current_search_retry}/{max_search_retries}): {e}")
                if current_search_retry == max_search_retries:
                    print(f"❌ {vid_search}: เกิดข้อผิดพลาดซ้ำๆ หลังจาก {max_search_retries} ครั้ง, ข้ามไปสินค้ารายการถัดไป")
                    break # ออกจากลูป while เพื่อไปสินค้ารายการถัดไป
                # หากเกิดข้อผิดพลาด ให้ปิดแท็บเสริมที่เปิดอยู่ก่อนลองใหม่
                try:
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                except:
                    pass
                time.sleep(2) # หน่วงเวลาเล็กน้อยก่อนลองใหม่

        # หากหลังจากลองใหม่หลายครั้งแล้วยังไม่สำเร็จ ให้ข้ามไปยังสินค้าถัดไป
        if not search_successful:
            continue # ไปยัง vid_search ถัดไปในลูป for หลัก

        # หากการค้นหาสำเร็จ ให้ดำเนินการดาวน์โหลดลิงก์ต่อ
        for li in li_list:
            li_text = li.text.strip()
            try:
                a_tags = li.find_elements(By.TAG_NAME, "a")
                if a_tags:
                    for a_tag in a_tags:
                        href = a_tag.get_attribute("href")
                        # ดึงข้อความของลิงก์โดยใช้ JavaScript เพื่อให้ได้ข้อความที่มองเห็นจริงๆ
                        a_text = driver.execute_script("return arguments[0].innerText;", a_tag).strip()
                        if not a_text:
                            a_text = li_text # ถ้าไม่มีข้อความลิงก์ ให้ใช้ข้อความของ li แทน
                        
                        # --- โหลดไฟล์ PDF, CAD, 3D ตามเงื่อนไข ---
                        href_lower = href.lower()
                        a_text_lower = a_text.lower()

                        # เงื่อนไขสำหรับ Datasheet (PDF)
                        if "tearsheet" in a_text_lower or href_lower.endswith(".pdf"):
                            download_and_move_file(driver, href, datasheet_folder, vid_folder, "PDF Datasheet")
                        # เงื่อนไขสำหรับ CAD BLOCK (2D)
                        elif "cad block" in a_text_lower:
                            download_and_move_file(driver, href, folder_2d, vid_folder, "CAD BLOCK (2D)")
                        # เงื่อนไขสำหรับโมเดล 3D (ไฟล์ประเภทอื่น)
                        # ตรวจสอบจากนามสกุลไฟล์ที่พบบ่อยสำหรับโมเดล 3D
                        elif any(ext in href_lower for ext in [".dwg", ".obj", ".3ds", ".stl", ".max", ".skp", ".fbx", ".rvt"]):
                            download_and_move_file(driver, href, folder_3d, vid_folder, "3D Model")
                        # ถ้าเป็นลิงก์อื่นๆ ที่ไม่ใช่ไฟล์ดาวน์โหลดที่ระบุ ก็จะถูกข้ามไป (ไม่มีการ print ข้อความ)

            except Exception as e:
                print(f"⚠️ {vid_search}: เกิดข้อผิดพลาดในการอ่านลิงก์ใน li '{li_text}': {e}")
        
        print(f"✅ {vid_search}: ตรวจสอบลิงก์และจัดการการดาวน์โหลดเสร็จสมบูรณ์")

        # ทำความสะอาดแท็บเพิ่มเติมสำหรับสินค้านี้ก่อนไปยังรายการถัดไป
        try:
            while len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[-1]) # สลับไปแท็บสุดท้าย
                driver.close() # ปิดแท็บ
            driver.switch_to.window(driver.window_handles[0]) # สลับกลับไปที่แท็บหลัก
        except:
            pass # จัดการข้อผิดพลาดถ้าไม่มีแท็บให้ปิดหรือสลับ

finally:
    # ปิดเบราว์เซอร์หลังจากทำงานทั้งหมดเสร็จสิ้น
    try:
        driver.quit()
    except:
        pass # จัดการข้อผิดพลาดถ้า driver ปิดไปแล้ว
