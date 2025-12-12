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

    # ---------- STEP 2: BASE Login ----------
    driver.get("https://base.chanintr.com/login")
    google_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Sign in with Google')]")
        )
    )
    google_btn.click()
    time.sleep(5)

    # ---------- STEP 3: อ่าน Excel ----------
    df = pd.read_excel(INPUT_FILE)
    required_cols = ["url", "name"]

    for col in required_cols:
        if col not in df.columns:
            logging.error(f"❌ ไม่มี column '{col}' ใน Excel")
            driver.quit()
            raise SystemExit(1)

    # ---------- STEP 4: Loop URL ----------
    for index, row in df.iterrows():
        url = row["url"]
        new_name = str(row["name"]).strip()

        logging.info(f"🔗 เปิด URL: {url}")
        logging.info(f"📝 ชื่อใหม่: {new_name}")

        driver.get(url)
        time.sleep(2)

        max_retries = 5

        for attempt in range(max_retries):
            try:
                # ---------- STEP 4C: คลิกปุ่ม info-header-buttons ----------
                try:
                    btn = driver.find_element(
                        By.CSS_SELECTOR,
                        "body > div > div > section > section > section.info-header-section.brand-info-header > div > div.info-header-buttons > button",
                    )
                    btn.click()
                    logging.info("✅ กดปุ่ม info-header-buttons สำเร็จ")
                    time.sleep(1)
                except Exception as e:
                    logging.warning(f"⚠️ กดปุ่ม info-header-buttons ไม่ได้: {e}")
                    continue

                # ---------- STEP 4D: เคลียร์ชื่อ และใส่ชื่อใหม่ ----------
                try:
                    name_input = wait.until(
                        EC.element_to_be_clickable(
                            (
                                By.XPATH,
                                "/html/body/div/div/section/section/div[1]/div/div[2]/div/div[2]/div[1]/div[2]/div/input",
                            )
                        )
                    )

                    name_input.clear()
                    time.sleep(0.3)
                    name_input.send_keys(new_name)

                    logging.info("✏️ เปลี่ยนชื่อสินค้าเรียบร้อย")
                except Exception as e:
                    logging.warning(f"⚠️ ใส่ชื่อใหม่ไม่สำเร็จ: {e}")
                    continue

                # ---------- STEP 4E: กด Submit ----------
                try:
                    submit_btn = driver.find_element(
                        By.CSS_SELECTOR,
                        "body > div > div > section > section > div.v--modal-overlay.scrollable.modal.modal-product-create-edit > div > div.v--modal-box.v--modal > div > div.modal-footer > button.btn",
                    )
                    submit_btn.click()
                    logging.info("💾 บันทึกข้อมูลสำเร็จ (Submit)")
                except Exception as e:
                    logging.warning(f"⚠️ Submit modal ไม่สำเร็จ: {e}")
                    continue

                break  # จบ retry loop ได้แล้ว

            except Exception as e:
                logging.warning(f"⚠️ Attempt {attempt+1}/{max_retries} error: {e}")
                time.sleep(1)

        time.sleep(1)

    driver.quit()

except Exception:
    logging.exception("เกิดข้อผิดพลาดระหว่างประมวลผล:")
    try:
        driver.quit()
    except:
        pass
    raise
