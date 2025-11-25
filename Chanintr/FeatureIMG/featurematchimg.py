# -*- coding: utf-8 -*-
import os
import re
import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
from difflib import SequenceMatcher

# ---------------- CONFIG ----------------
output_path = r"C:\Users\tanapat\Downloads\base_feature_images_all_pages_Match_Finish.xlsx"
GOOGLE_EMAIL = "tanapat@chanintr.com"
GOOGLE_PASSWORD = "Qwerty12345$$"
wait_time = 10
folder_path = r"D:\HIC Feture\Finish"

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

# ---------------- ฟังก์ชันช่วย ----------------
def clean_string(s):
    s = s.lower().strip()
    s = re.sub(r'[\s\-\_\.\,\/]', '', s)
    return s

def find_best_match(text, folder_path):
    text_clean = clean_string(text)
    best_match = None
    max_ratio = 0
    for filename in os.listdir(folder_path):
        file_clean = clean_string(os.path.splitext(filename)[0])
        ratio = SequenceMatcher(None, text_clean, file_clean).ratio()
        if ratio > max_ratio:
            max_ratio = ratio
            best_match = filename
    return best_match, max_ratio

# ---------------- ดึงข้อมูลจากหน้า ----------------
def extract_page_data():
    data_page = []
    seen_names = set()  # เก็บชื่อ h3 ที่เจอแล้ว

    try:
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul > li > div.cell-thumbnail img")))
        time.sleep(1.5)
        items = driver.find_elements(By.CSS_SELECTOR, "ul > li")

        for li in items:
            try:
                img_el = li.find_element(By.CSS_SELECTOR, "div.cell-thumbnail img")
                img_url = img_el.get_attribute("src")

                h3_el = li.find_element(By.CSS_SELECTOR, "div:nth-child(3) h3")
                h3_text = h3_el.text.strip()

                # ใช้ h3 เป็น name เสมอ
                name_text = h3_text
                code_text = ""

                # ถ้า h3 ซ้ำ ตัวที่ซ้ำใช้ <p> เป็น code และเช็ค folder ด้วย code แทน
                if h3_text in seen_names:
                    try:
                        p_el = li.find_element(By.CSS_SELECTOR, "div:nth-child(3) p")
                        code_text = p_el.text.strip()
                        match_text = code_text  # ใช้ code แทน h3 สำหรับเช็ค folder
                    except NoSuchElementException:
                        code_text = ""
                        match_text = name_text  # fallback เป็น h3
                else:
                    seen_names.add(h3_text)
                    match_text = name_text  # ใช้ h3 ปกติ

                matched_file, match_ratio = find_best_match(match_text, folder_path)

                print(f"ชื่อ: {name_text}")
                print(f"รูป: {img_url}")
                print(f"Code: {code_text}")
                if matched_file:
                    print(f"📂 ไฟล์ที่ตรงกันมากที่สุด: {matched_file} (ความใกล้เคียง {match_ratio:.2f})")
                else:
                    print("⚠️ ไม่พบไฟล์ที่ตรงกัน")
                print("-" * 80)

                data_page.append({
                    "name": name_text,
                    "code": code_text,
                    "image_url": img_url,
                    "matched_file": matched_file,
                    "match_ratio": match_ratio
                })
            except Exception:
                continue

        logging.info(f"✅ ดึงข้อมูลจากหน้านี้ได้ {len(data_page)} รายการ")
    except TimeoutException:
        logging.warning("⚠️ ไม่มีรายการในหน้านี้")
    return data_page

# ---------------- เริ่มลูปดึงข้อมูลทุกหน้า ----------------
all_data = []
page = 1

while True:
    logging.info(f"📄 กำลังดึงข้อมูลหน้า {page} ...")
    all_data.extend(extract_page_data())

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
            time.sleep(2)
        else:
            logging.info("✅ ไม่มีหน้าถัดไปแล้ว — หยุดลูป")
            break

    except (TimeoutException, ElementClickInterceptedException, NoSuchElementException):
        logging.info("✅ ไม่มีปุ่ม Next — หยุดลูป")
        break

# ---------------- บันทึก Excel ----------------
if all_data:
    df = pd.DataFrame(all_data)
    df.to_excel(output_path, index=False)
    logging.info(f"💾 บันทึกข้อมูลลง Excel: {output_path}")

# ---------------- ปิด Browser ----------------
driver.quit()
logging.info("🚪 ปิด Browser แล้วเสร็จสิ้น")
