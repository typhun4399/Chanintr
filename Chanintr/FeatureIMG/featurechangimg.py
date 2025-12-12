# -*- coding: utf-8 -*-
import os
import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ---------------- CONFIG ----------------
excel_path = (
    r"C:\Users\tanapat\Downloads\base_feature_images_all_pages_Match_Finish.xlsx"
)
folder_path = r"D:\HIC Feture\Finish"
print("GOOGLE_EMAIL")
GOOGLE_EMAIL = input()
print("GOOGLE_PASSWORD")
GOOGLE_PASSWORD = input()

# ---------------- Chrome Options ----------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# ---------------- Logging ----------------
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# ---------------- STEP 1: Google Login ----------------
try:
    driver.get("https://accounts.google.com/signin/v2/identifier")

    email_input = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='email']"))
    )
    email_input.clear()
    email_input.send_keys(GOOGLE_EMAIL)
    driver.find_element(By.ID, "identifierNext").click()

    password_input = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='password']"))
    )
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
    google_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Sign in with Google')]")
        )
    )
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

# ---------------- STEP 4: โหลด Excel ----------------
try:
    df = pd.read_excel(excel_path)
    if "name" not in df.columns or "matched_file" not in df.columns:
        logging.error("❌ ไม่พบคอลัมน์ 'name' หรือ 'matched_file' ในไฟล์ Excel")
        driver.quit()
        exit()

    records = (
        df[["name", "matched_file"]].dropna(subset=["name"]).to_dict(orient="records")
    )
    logging.info(f"📘 โหลดข้อมูลทั้งหมด {len(records)} รายการจาก Excel")
except Exception as e:
    logging.error(f"❌ โหลด Excel ไม่สำเร็จ: {e}")
    driver.quit()
    exit()

# ---------------- STEP 5: วนพิมพ์ name และอัปโหลด matched_file ----------------
for idx, row in enumerate(records, start=1):
    name = row["name"]
    matched_file = row["matched_file"]
    file_path = os.path.join(folder_path, matched_file) if matched_file else None

    try:
        # 🔹 ใส่ชื่อในช่อง input
        input_box = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[3]/div[3]/div[1]/div[1]/input",
                )
            )
        )
        input_box.clear()
        input_box.send_keys(name)
        logging.info(f"✅ ({idx}/{len(records)}) พิมพ์ชื่อ '{name}' สำเร็จ")

        time.sleep(2)  # รอให้ผลลัพธ์โหลดก่อน

        # 🔹 หา input[type=file] และอัปโหลดไฟล์
        if file_path and os.path.exists(file_path):
            try:
                file_input = wait.until(
                    EC.presence_of_element_located(
                        (
                            By.CSS_SELECTOR,
                            "body > div > div > section > section > section.wrapper-container.brand-features-wrapper "
                            "> div:nth-child(1) > section > div > div.collection-section > div "
                            "> div.collection-result-container > ul > li:nth-child(1) "
                            "> div.cell-thumbnail > div > div:nth-child(1) > input[type=file]",
                        )
                    )
                )
                file_input.send_keys(file_path)
                logging.info(f"📤 อัปโหลดไฟล์ '{matched_file}' สำเร็จ")
            except TimeoutException:
                logging.warning(f"⚠️ ({idx}) ไม่พบช่องอัปโหลดไฟล์ในหน้านี้")
        else:
            logging.warning(f"⚠️ ({idx}) ไม่พบไฟล์ในโฟลเดอร์: {file_path}")

        time.sleep(1)
    except TimeoutException:
        logging.warning(f"⚠️ ({idx}) ไม่พบช่อง input ในหน้านี้")
        break
    except Exception as e:
        logging.warning(f"⚠️ ({idx}) เกิดข้อผิดพลาด: {e}")
        continue

# ---------------- STEP 6: ปิด Browser ----------------
driver.quit()
logging.info("🚪 ปิด Browser แล้วเสร็จสิ้น")
