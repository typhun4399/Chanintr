# -*- coding: utf-8 -*-
import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
)

# ---------------- CONFIG ----------------
print("GOOGLE_EMAIL")
GOOGLE_EMAIL = input()
print("GOOGLE_PASSWORD")
GOOGLE_PASSWORD = input()
INPUT_FILE = r"C:\Users\tanapat\Desktop\base_products.xlsx"

# ---------------- Chrome Options ----------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

try:
    # ---------- STEP 1: Google Login ----------
    logging.info("🌐 เริ่ม Login Google...")
    driver.get("https://accounts.google.com/signin/v2/identifier")
    email_input = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//input[@type='email' or @id='identifierId']")
        )
    )
    email_input.send_keys(GOOGLE_EMAIL)
    driver.find_element(By.ID, "identifierNext").click()

    password_input = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='password']"))
    )
    time.sleep(0.5)
    password_input.send_keys(GOOGLE_PASSWORD)
    driver.find_element(By.ID, "passwordNext").click()
    time.sleep(3)

    # ---------- STEP 2: Base Login ----------
    logging.info("🔐 Login Base ด้วย Google...")
    driver.get("https://base.chanintr.com/login")
    google_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Sign in with Google')]")
        )
    )
    google_btn.click()
    time.sleep(5)

    # ---------- STEP 3: Load Excel ----------
    logging.info("📘 โหลด Excel...")
    df = pd.read_excel(INPUT_FILE)
    if "url" not in df.columns:
        logging.error("❌ ไม่มี column 'url' ใน Excel")
        driver.quit()
        raise SystemExit(1)

    collections = []
    designers = []

    # ---------- STEP 4: Iterate URLs ----------
    for index, row in df.iterrows():
        url = row["url"]
        logging.info(f"\n========================================")
        logging.info(f"🔗 เปิด URL ({index+1}): {url}")
        logging.info(f"========================================")

        driver.get(url)
        time.sleep(2)

        # ------------------ GET COLLECTION ------------------
        try:
            elem = driver.find_element(
                By.XPATH,
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[1]/div[3]/div[2]/ul/li/div",
            )
            collection = elem.text.strip()
            logging.info(f"📦 Collection: {collection}")
        except NoSuchElementException:
            collection = ""
            logging.warning("⚠️ Collection ไม่พบ (ใส่ว่าง)")

        collections.append(collection)

        # ------------------ GET DESIGNER ------------------
        try:
            elem = driver.find_element(
                By.XPATH,
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[3]/div[2]/span",
            )
            designer = elem.text.strip()
            logging.info(f"🎨 Designer: {designer}")
        except NoSuchElementException:
            designer = ""
            logging.warning("⚠️ Designer ไม่พบ (ใส่ว่าง)")

        designers.append(designer)

    # ---------- STEP 5: Save Output ----------
    df["Collection"] = collections
    df["Designer"] = designers

    OUTPUT_FILE = INPUT_FILE.replace(".xlsx", "_with_collection_designer.xlsx")
    df.to_excel(OUTPUT_FILE, index=False)

    logging.info(f"\n✅ เสร็จสมบูรณ์! บันทึกไฟล์ที่: {OUTPUT_FILE}")

except Exception as e:
    logging.error(f"🚨 ERROR: {e}")

finally:
    driver.quit()
    logging.info("👋 ปิด Browser แล้ว")
