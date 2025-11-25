# -*- coding: utf-8 -*-
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
import time
import logging

# ---------------- CONFIG ----------------
output_path = r"C:\Users\tanapat\Downloads\base_feature_images_all_pages_Finish.xlsx"

GOOGLE_EMAIL = "tanapat@chanintr.com"
GOOGLE_PASSWORD = "Qwerty12345$$"
wait_time = 10

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
target_url = "https://base.chanintr.com/brand/16/features?featureTypeId=1&isUnassigned=false&isSearch=false"
driver.get(target_url)
logging.info("🌐 เปิดหน้า features แล้ว รอโหลดเนื้อหา...")

# ---------------- STEP 4: ฟังก์ชันดึงข้อมูลในแต่ละหน้า ----------------
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

                # ถ้า name ซ้ำ ให้ดึง code ด้วย
                if name_text in seen_names:
                    try:
                        code_el = li.find_element(By.CSS_SELECTOR, "div:nth-child(3) p")
                        row_data["code"] = code_el.text.strip()
                    except NoSuchElementException:
                        row_data["code"] = ""
                else:
                    seen_names.add(name_text)

                data_page.append(row_data)
                print(f"ชื่อ: {name_text}")
                print(f"รูป: {img_url}")
                if "code" in row_data:
                    print(f"Code: {row_data['code']}")
                print("-" * 80)
            except Exception:
                continue

        logging.info(f"✅ ดึงข้อมูลจากหน้านี้ได้ {len(data_page)} รายการ")
    except TimeoutException:
        logging.warning("⚠️ ไม่มีรายการในหน้านี้")
    return data_page

# ---------------- STEP 5: วนลูปทุกหน้า ----------------
all_data = []
page = 1

while True:
    logging.info(f"📄 กำลังดึงข้อมูลหน้า {page} ...")
    all_data.extend(extract_page_data())

    # หา "ปุ่มถัดไป" จาก pagination
    try:
        next_btn = None
        buttons = driver.find_elements(By.CSS_SELECTOR, "ul.pagination li button")
        for btn in buttons:
            svg = btn.find_elements(By.CSS_SELECTOR, "svg[data-icon='angle-right']")
            if svg:
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

    except (TimeoutException, ElementClickInterceptedException, NoSuchElementException):
        logging.info("✅ ไม่มีปุ่ม Next — หยุดลูป")
        break

# ---------------- STEP 6: บันทึกเป็น Excel ----------------
pd.DataFrame(all_data).to_excel(output_path, index=False)
logging.info(f"💾 บันทึกผลลัพธ์ทั้งหมดลงไฟล์เรียบร้อย: {output_path}")
logging.info("🚪 ปิด Browser แล้วเสร็จสิ้น")

# -*- coding: utf-8 -*-
import os
import pandas as pd
import urllib.request
import time
from PIL import Image, ImageOps

# ---------------- CONFIG ----------------
excel_files = [r"C:\Users\tanapat\Downloads\base_feature_images_all_pages_Finish.xlsx"]
base_folder = r"D:\HIC Feture\Finish"
origin_folder = os.path.join(base_folder, "origin")
crop_folder = os.path.join(base_folder, "crop")

# สร้างโฟลเดอร์
os.makedirs(origin_folder, exist_ok=True)
os.makedirs(crop_folder, exist_ok=True)

initial_size = 500  # ขนาด 500x500 หลังโหลด
crop_top_bottom = 75  # ตัดบนและล่างออก 75px
crop_left_right = 75  # ตัดซ้าย/ขวาออก 75px

# ---------------- READ EXCEL ----------------
all_dfs = []
for file in excel_files:
    try:
        df = pd.read_excel(file)
        all_dfs.append(df)
        print(f"Loaded {file} with {len(df)} rows.")
    except Exception as e:
        print(f"❌ Failed to load {file}: {e}")

combined_df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

# ---------------- STEP 1: DOWNLOAD & RESIZE 500x500 (เก็บใน origin) ----------------
if 'image_url' not in combined_df.columns or 'name' not in combined_df.columns:
    print("❌ Columns 'image_url' or 'name' not found in Excel.")
else:
    for idx, row in combined_df.dropna(subset=['image_url', 'name']).iterrows():
        try:
            url = row['image_url']

            # ตั้งชื่อไฟล์: ถ้า code มีค่าใช้ code, ถ้าไม่มีใช้ name
            if 'code' in row and pd.notna(row['code']):
                base_name = str(row['code']).strip()
            else:
                base_name = str(row['name']).strip()

            base_name = base_name.replace('/', '_').replace('\\', '_')
            filename = f"{base_name}.jpg"
            origin_path = os.path.join(origin_folder, filename)
            crop_path = os.path.join(crop_folder, filename)

            # ดาวน์โหลดภาพ
            print(f"Downloading {idx+1}: {url} -> {filename}")
            urllib.request.urlretrieve(url, origin_path)
            time.sleep(0.2)

            # Resize เป็น 500x500
            with Image.open(origin_path) as img:
                img_resized = ImageOps.fit(img, (initial_size, initial_size), method=Image.LANCZOS)
                img_resized.save(origin_path, format='JPEG')

            # ก็อปปี้ไปโฟลเดอร์ crop สำหรับขั้นตอนต่อไป
            img_resized.save(crop_path, format='JPEG')

        except Exception as e:
            print(f"❌ Failed to process {url}: {e}")

    print("✅ Finished downloading and resizing all images to 500x500 px (saved in origin & crop)")

# ---------------- STEP 2: CROP บน/ล่าง 75px ----------------
for file in os.listdir(crop_folder):
    if file.lower().endswith(".jpg"):
        filepath = os.path.join(crop_folder, file)
        try:
            with Image.open(filepath) as img:
                width, height = img.size
                top = crop_top_bottom
                bottom = height - crop_top_bottom
                img_cropped = img.crop((0, top, width, bottom))
                img_cropped.save(filepath, format='JPEG')
        except Exception as e:
            print(f"❌ Failed to crop top/bottom {file}: {e}")

print(f"✅ Finished cropping top/bottom {crop_top_bottom}px for all images in 'crop'")

# ---------------- STEP 3: CROP ซ้าย/ขวา 75px ----------------
for file in os.listdir(crop_folder):
    if file.lower().endswith(".jpg"):
        filepath = os.path.join(crop_folder, file)
        try:
            with Image.open(filepath) as img:
                width, height = img.size
                left = crop_left_right
                right = width - crop_left_right
                img_cropped = img.crop((left, 0, right, height))
                img_cropped.save(filepath, format='JPEG')
        except Exception as e:
            print(f"❌ Failed to crop left/right {file}: {e}")

print(f"✅ Finished cropping left/right {crop_left_right}px for all images in 'crop'")
