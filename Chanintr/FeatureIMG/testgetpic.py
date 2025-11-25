# -*- coding: utf-8 -*-
import os
import time
import logging
import pandas as pd
import urllib.request
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from PIL import Image, ImageOps

# ---------------- CONFIG ----------------
output_path = r"C:\Users\tanapat\Downloads\base_feature_images_all_pages_BAK_Leather.xlsx"
base_folder = r"D:\HIC Feture\test"
origin_folder = os.path.join(base_folder, "origin")
crop_folder = os.path.join(base_folder, "crop")

GOOGLE_EMAIL = "tanapat@chanintr.com"
GOOGLE_PASSWORD = "Qwerty12345$$"

crop_top_bottom = 75
crop_left_right = 75

os.makedirs(origin_folder, exist_ok=True)
os.makedirs(crop_folder, exist_ok=True)

# ---------------- Chrome Options ----------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# ---------------- Logging ----------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# ---------------- STEP 1: Google Login ----------------
try:
    driver.get("https://accounts.google.com/signin/v2/identifier")

    email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
    email_input.clear()
    email_input.send_keys(GOOGLE_EMAIL)
    driver.find_element(By.ID, "identifierNext").click()

    password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
    password_input.clear()
    password_input.send_keys(GOOGLE_PASSWORD)
    driver.find_element(By.ID, "passwordNext").click()

    logging.info("✅ ล็อกอิน Google สำเร็จ")
    time.sleep(5)
except TimeoutException:
    logging.error("❌ ล็อกอิน Google ไม่สำเร็จ")
    driver.quit()
    exit()

# ---------------- STEP 2: Base Login ----------------
try:
    driver.get("https://base.chanintr.com/login")
    google_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Sign in with Google')]")))
    google_btn.click()
    logging.info("✅ กดปุ่ม Sign in with Google สำเร็จ")
    time.sleep(10)
except TimeoutException:
    logging.error("❌ ไม่พบปุ่ม Sign in with Google")
    driver.quit()
    exit()

# ---------------- STEP 3: เข้าไปหน้า features ----------------
target_url = "https://base.chanintr.com/brand/10/features?featureTypeId=6&isUnassigned=false&isSearch=false"
driver.get(target_url)
logging.info("🌐 เปิดหน้า features แล้ว รอโหลดเนื้อหา...")

# ---------------- STEP 4: ดึงข้อมูลแต่ละหน้า ----------------
def extract_page_data():
    data_page = []
    seen_names = set()
    try:
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul > li > div.cell-thumbnail img")))
        time.sleep(1.5)
        items = driver.find_elements(By.CSS_SELECTOR, "ul > li")

        for li in items:
            try:
                img_el = li.find_element(By.CSS_SELECTOR, "div.cell-thumbnail img")
                img_url = img_el.get_attribute("src")

                name_el = li.find_element(By.CSS_SELECTOR, "div:nth-child(3) h3")
                name_text = name_el.text.strip()

                row_data = {"name": name_text, "image_url": img_url}

                # ถ้า name ซ้ำ -> ดึง code เพิ่ม
                if name_text in seen_names:
                    try:
                        code_el = li.find_element(By.CSS_SELECTOR, "div:nth-child(3) p")
                        row_data["code"] = code_el.text.strip()
                    except NoSuchElementException:
                        row_data["code"] = ""
                else:
                    seen_names.add(name_text)

                data_page.append(row_data)
            except Exception:
                continue

        logging.info(f"✅ ดึงข้อมูลจากหน้านี้ได้ {len(data_page)} รายการ")
    except TimeoutException:
        logging.warning("⚠️ ไม่มีรายการในหน้านี้")
    return data_page

# ---------------- STEP 5: วนทุกหน้า ----------------
all_data = []
page = 1
while True:
    logging.info(f"📄 กำลังดึงข้อมูลหน้า {page} ...")
    all_data.extend(extract_page_data())

    # หา next page
    try:
        next_btn = None
        buttons = driver.find_elements(By.CSS_SELECTOR, "ul.pagination li button")
        for btn in buttons:
            if btn.find_elements(By.CSS_SELECTOR, "svg[data-icon='angle-right']"):
                next_btn = btn
                break

        if next_btn and next_btn.is_enabled():
            driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
            time.sleep(1)
            next_btn.click()
            logging.info(f"➡️ ไปหน้าถัดไป (หน้า {page + 1})")
            page += 1
            time.sleep(1)
        else:
            logging.info("✅ ไม่มีหน้าถัดไปแล้ว — หยุดลูป")
            break
    except Exception:
        logging.info("✅ ไม่มีปุ่ม Next — หยุดลูป")
        break

driver.quit()

# ---------------- STEP 6: บันทึก Excel ----------------
pd.DataFrame(all_data).to_excel(output_path, index=False)
logging.info(f"💾 บันทึกผลลัพธ์ทั้งหมดลงไฟล์: {output_path}")

# ---------------- STEP 7: โหลดภาพ origin ----------------
df = pd.read_excel(output_path)

if 'image_url' not in df.columns or 'name' not in df.columns:
    print("❌ Columns 'image_url' or 'name' not found.")
else:
    for idx, row in df.dropna(subset=['image_url', 'name']).iterrows():
        try:
            url = row['image_url']
            base_name = str(row['code']).strip() if 'code' in row and pd.notna(row['code']) else str(row['name']).strip()
            base_name = base_name.replace('/', '_').replace('\\', '_')
            filename = f"{base_name}.jpg"
            origin_path = os.path.join(origin_folder, filename)
            crop_path = os.path.join(crop_folder, filename)

            if os.path.exists(origin_path):
                print(f"⏩ Skip existing: {filename}")
                continue

            print(f"Downloading {idx+1}: {url} -> {filename}")
            urllib.request.urlretrieve(url, origin_path)
            time.sleep(0.2)

            # ก็อปปี้ไป crop
            with Image.open(origin_path) as img:
                img.save(crop_path, format='JPEG')

        except Exception as e:
            print(f"❌ Failed to download {url}: {e}")

print("✅ Finished downloading all original images")

# ---------------- STEP 8: Crop 75px บน/ล่าง/ซ้าย/ขวา ----------------
for file in os.listdir(crop_folder):
    if file.lower().endswith(".jpg"):
        filepath = os.path.join(crop_folder, file)
        try:
            with Image.open(filepath) as img:
                width, height = img.size
                if height > crop_top_bottom * 2 and width > crop_left_right * 2:
                    img_cropped = img.crop((
                        crop_left_right,
                        crop_top_bottom,
                        width - crop_left_right,
                        height - crop_top_bottom
                    ))
                    img_cropped.save(filepath, format='JPEG')
        except Exception as e:
            print(f"❌ Failed to crop {file}: {e}")

print("✅ Finished cropping all images in 'crop'")