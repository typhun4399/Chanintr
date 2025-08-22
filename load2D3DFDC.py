import os
import time
import shutil
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# 📂 โหลดข้อมูลจาก Excel
excel_path = r"C:\Users\tanapat\Downloads\1_FDC active model id_27Jun25.xlsx"
df = pd.read_excel(excel_path)

# ✅ ตรวจสอบคอลัมน์
required_columns = {"product_vendor", "Foldername"}
if not required_columns.issubset(df.columns):
    raise ValueError(f"ไม่พบคอลัมน์ {required_columns} ในไฟล์ Excel")

# 📁 Path หลัก
base_download_dir = r"C:\Users\tanapat\Desktop\FDC\2D&3D"
temp_download_dir = r"C:\Users\tanapat\Desktop\FDC\temp"
os.makedirs(temp_download_dir, exist_ok=True)

# 🧼 ฟังก์ชันแก้ชื่อโฟลเดอร์ให้ปลอดภัย
def sanitize_foldername(name):
    return re.sub(r'[<>:"/\\|?*]', "_", name)

# 🌐 ตั้งค่า Chrome
options = Options()
options.add_argument("--start-maximized")
prefs = {
    "download.default_directory": temp_download_dir,
    "profile.default_content_setting_values.automatic_downloads": 1
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(service=Service(), options=options)

# 🔖 ปุ่มดาวน์โหลดที่ต้องการ
download_labels = ["product sheet", "2d dwg", "3d dwg", "obj"]

try:
    driver.get("https://www.fredericia.com/")

    # 🍪 Accept cookies
    try:
        accept_cookie_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.coi-banner__accept[aria-label='Accept all']"))
        )
        accept_cookie_button.click()
        print("✅ Accepted cookies.")
    except Exception:
        print("⚠️ No cookie banner found or could not click 'Accept all'.")

    # 🔁 ลูปแต่ละ vendor
    for _, row in df.iterrows():
        vendor = str(row["product_vendor"])
        foldername_raw = str(row["Foldername"])
        foldername_safe = sanitize_foldername(foldername_raw)
        target_folder = os.path.join(base_download_dir, foldername_safe)
        os.makedirs(target_folder, exist_ok=True)

        try:
            # 🔍 กดปุ่ม Search ด้วย JavaScript click
            try:
                search_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/header/div/div/nav[2]/div[1]/div[3]/div/div/button[1]"))
                )
                driver.execute_script("arguments[0].click();", search_button)
                print("✅ Clicked search button.")
            except Exception as e:
                print(f"⚠️ Failed to click search button: {e}")
                continue

            # 🔤 ใส่ vendor ลงในช่องค้นหา
            try:
                search_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Search here']"))
                )
                driver.execute_script("arguments[0].value = '';", search_input)
                search_input.send_keys(vendor)
                print(f"🔎 Typed '{vendor}' into search input.")
                time.sleep(5)
            except Exception as e:
                print(f"⚠️ Failed to type in search input: {e}")
                continue

            # ตรวจสอบว่ามีลิงก์ผลลัพธ์
            search_results = driver.find_elements(By.XPATH, "/html/body/div[2]/div/header/div/div/div/div[2]/div[2]/div/div/div[1]/div/a")
            if not search_results:
                print(f"⚠️ No search results found for '{vendor}', clearing input.")
                search_input.clear()
                continue

            # คลิกผลลัพธ์ลำดับแรก
            search_results[0].click()
            print(f"✅ Clicked top result for '{vendor}'")
            time.sleep(5)

            # ✅ คลิก dropdown
            try:
                product_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/main/div[2]/div/div[3]/div[2]/div/div[5]/button"))
                )
                driver.execute_script("arguments[0].click();", product_button)
                print(f"🔘 Opened product dropdown for '{vendor}'")
                time.sleep(2)

                # ✅ หา <a> ทั้งหมด
                all_links = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.TAG_NAME, "a"))
                )

                available_downloads = {}
                for link in all_links:
                    try:
                        span = link.find_element(By.TAG_NAME, "span")
                        label = span.text.strip().lower()
                        if label in download_labels:
                            available_downloads[label] = link
                    except Exception:
                        continue

                if not available_downloads:
                    print("⚠️ No downloadable files found for this product.")
                else:
                    for label in download_labels:
                        if label in available_downloads:
                            try:
                                before = set(os.listdir(temp_download_dir))
                                driver.execute_script("arguments[0].click();", available_downloads[label])
                                print(f"⬇️ Downloading: {label}")

                                # ⏳ รอไฟล์ใหม่
                                timeout = time.time() + 20
                                downloaded_file = None
                                while True:
                                    time.sleep(1)
                                    after = set(os.listdir(temp_download_dir))
                                    new_files = after - before
                                    if new_files:
                                        downloaded_file = new_files.pop()
                                        break
                                    if time.time() > timeout:
                                        raise TimeoutError("Download timeout")

                                # ✅ จัดเก็บในโฟลเดอร์ย่อย
                                if label == "product sheet":
                                    subfolder = "Datasheet"
                                elif label == "2d dwg":
                                    subfolder = "2D"
                                elif label in ["3d dwg", "obj"]:
                                    subfolder = "3D"
                                else:
                                    subfolder = ""

                                final_folder = os.path.join(target_folder, subfolder)
                                os.makedirs(final_folder, exist_ok=True)

                                src = os.path.join(temp_download_dir, downloaded_file)
                                dst = os.path.join(final_folder, downloaded_file)
                                shutil.move(src, dst)
                                print(f"✅ Moved {label} → {dst}")

                            except Exception as e:
                                print(f"❌ Failed to download {label}: {e}")

            except Exception as e:
                print(f"❌ Failed to handle dropdown: {e}")

        except Exception as e:
            print(f"❌ Error processing '{vendor}': {e}")

finally:
    driver.quit()
    print("🧹 Browser closed.")
