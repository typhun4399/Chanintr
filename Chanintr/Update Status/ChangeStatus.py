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

    # ---------- STEP 2: Base Login ----------
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
    if "url" not in df.columns:
        logging.error("❌ ไม่มี column 'url' ใน Excel")
        driver.quit()
        raise SystemExit(1)

    statuses = []

    # ---------- STEP 4: วน URL ----------
    for index, row in df.iterrows():
        url = row["url"]
        logging.info(f"🔗 เปิด URL: {url}")

        driver.get(url)
        time.sleep(2)

        status_text = ""
        max_retries = 5
        for attempt in range(max_retries):
            try:
                elem = wait.until(
                    EC.presence_of_element_located(
                        (
                            By.CSS_SELECTOR,
                            "body > div > div > section > section > section.info-header-section.brand-info-header > div > div.info-header-container-wrapper > div.title-container > div.desc-container > div.desc-info.desc-info-status > div > div",
                        )
                    )
                )
                status_text = elem.text.strip()
                if status_text:
                    if status_text == "Order with condition":
                        logging.info("⚠️ สถานะเป็น 'Order with condition' ข้ามรายการนี้")
                        status_text = ""
                    elif status_text == "Discontinued":
                        logging.info("⚠️ สถานะเป็น 'Discontinued' ข้ามรายการนี้")
                        status_text = ""
                    else:
                        logging.info(f"📄 สถานะสินค้า: {status_text}")

                        # ---------- STEP 4C: กดปุ่ม info-header-buttons ----------
                        try:
                            btn = driver.find_element(
                                By.CSS_SELECTOR,
                                "body > div > div > section > section > section.info-header-section.brand-info-header > div > div.info-header-buttons > button",
                            )
                            btn.click()
                            logging.info("✅ กดปุ่มเรียบร้อย")
                            time.sleep(1)

                            # ---------- STEP 4D: กดปุ่ม modal ----------
                            modal_elem = wait.until(
                                EC.element_to_be_clickable(
                                    (
                                        By.CSS_SELECTOR,
                                        "body > div > div > section > section > div.v--modal-overlay.scrollable.modal.modal-product-create-edit > div > div.v--modal-box.v--modal > div > div.modal-content > div:nth-child(6) > div.col-md-8 > div",
                                    )
                                )
                            )
                            modal_elem.click()
                            logging.info("✅ กดปุ่มใน modal เรียบร้อย")
                            time.sleep(0.5)

                            # ---------- STEP 4E: กด dropdown li:nth-child(5) ----------
                            dropdown_item = wait.until(
                                EC.element_to_be_clickable(
                                    (
                                        By.CSS_SELECTOR,
                                        "body > div > div > section > section > div.v--modal-overlay.scrollable.modal.modal-product-create-edit > div > div.v--modal-box.v--modal > div > div.modal-content > div:nth-child(6) > div.col-md-8 > div > div.dropdown-bubble-container > div > div.dropdown-list-container > ul > li:nth-child(5)",
                                    )
                                )
                            )
                            dropdown_item.click()
                            logging.info("✅ กด dropdown li:nth-child(5) เรียบร้อย")
                            time.sleep(0.5)

                            # ---------- STEP 4F: รอ text เปลี่ยนเป็น 'Order with condition' ----------
                            success = False
                            for retry in range(10):
                                value_elem = driver.find_element(
                                    By.CSS_SELECTOR,
                                    "body > div > div > section > section > div.v--modal-overlay.scrollable.modal.modal-product-create-edit > div > div.v--modal-box.v--modal > div > div.modal-content > div:nth-child(6) > div.col-md-8 > div.dropdown-container > div.dropdown-value > span.dropdown-value-text",
                                )
                                value_text = value_elem.text.strip()
                                if value_text == "Order with condition":
                                    logging.info(
                                        "✅ dropdown text เปลี่ยนเป็น 'Order with condition'"
                                    )
                                    success = True
                                    break
                                else:
                                    logging.info(
                                        f"⏳ รอ text เปลี่ยน, ปัจจุบัน: '{value_text}'"
                                    )
                                    time.sleep(0.5)
                            if not success:
                                logging.warning(
                                    "⚠️ dropdown text ไม่เปลี่ยนเป็น 'Order with condition'"
                                )

                            submit_btn = driver.find_element(
                                By.CSS_SELECTOR,
                                "body > div > div > section > section > div.v--modal-overlay.scrollable.modal.modal-product-create-edit > div > div.v--modal-box.v--modal > div > div.modal-footer > button.btn",
                            )
                            submit_btn.click()
                            logging.info("✅ กด Submit modal เรียบร้อย")

                        except Exception as e:
                            logging.warning(
                                f"⚠️ ปัญหาระหว่างกดปุ่ม/ modal/ dropdown/ checkbox: {e}"
                            )

                    break
                else:
                    logging.warning(
                        f"⚠️ ยังไม่พบ status (ว่าง) ครั้งที่ {attempt+1}, retry ..."
                    )
                    time.sleep(2)
            except TimeoutException:
                logging.warning(f"⚠️ ไม่พบ element สถานะสินค้า ครั้งที่ {attempt+1}, retry ...")
                time.sleep(2)

        if not status_text:
            logging.warning("⚠️ ดึง status ไม่สำเร็จหรือถูกข้าม, บันทึกเป็นว่าง")

        statuses.append(status_text)
        time.sleep(1)

    driver.quit()

except Exception:
    logging.exception("เกิดข้อผิดพลาดระหว่างดึงสถานะสินค้า:")
    try:
        driver.quit()
    except:
        pass
    raise
