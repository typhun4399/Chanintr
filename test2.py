# -*- coding: utf-8 -*-
import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

# ---------------- CONFIG ----------------
GOOGLE_EMAIL = "tanapat@chanintr.com"
GOOGLE_PASSWORD = "Qwerty12345$$"
EXCEL_FILE = r"C:\Users\tanapat\Desktop\vendor_code.xlsx"
OUTPUT_FILE = r"C:\Users\tanapat\Desktop\vendor_output.xlsx"
NEW_VALUE = "Value"

# ---------------- Chrome Options ----------------
CHROME_OPTIONS = Options()
CHROME_OPTIONS.add_argument("--start-maximized")
CHROME_OPTIONS.add_argument("--disable-notifications")
CHROME_OPTIONS.add_argument("--no-sandbox")
CHROME_OPTIONS.add_argument("--disable-dev-shm-usage")
CHROME_OPTIONS.add_experimental_option("excludeSwitches", ["enable-logging"])

# ---------------- SETUP ----------------
driver = webdriver.Chrome(options=CHROME_OPTIONS)
driver.set_page_load_timeout(300)  # รอโหลดหน้าสูงสุด 5 นาที
wait = WebDriverWait(driver, 30)   # รอ element clickable / presence

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def safe_get(url, retries=3, delay=5):
    """โหลด url พร้อม retry กรณี timeout"""
    for attempt in range(1, retries + 1):
        try:
            driver.get(url)
            return True
        except WebDriverException as e:
            print(f"⏳ โหลดหน้า {url} attempt {attempt} failed: {e}")
            time.sleep(delay)
    print(f"❌ โหลดหน้า {url} ไม่สำเร็จหลัง {retries} ครั้ง")
    return False

try:
    # ---------- STEP 1: Google Login ----------
    driver.get("https://accounts.google.com/signin/v2/identifier")
    email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email' or @id='identifierId']")))
    email_input.send_keys(GOOGLE_EMAIL)
    driver.find_element(By.ID, "identifierNext").click()

    password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
    time.sleep(0.5)
    password_input.send_keys(GOOGLE_PASSWORD)
    driver.find_element(By.ID, "passwordNext").click()
    time.sleep(3)

    # ---------- STEP 2: Base Login ----------
    if safe_get("https://base.chanintr.com/login"):
        google_btn = WebDriverWait(driver, 180).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Sign in with Google')]"))
        )
        google_btn.click()
        time.sleep(10)
    else:
        raise Exception("ไม่สามารถโหลดหน้า Base login ได้")

    # ---------- STEP 3: อ่าน Excel ----------
    df = pd.read_excel(EXCEL_FILE)
    if "url" not in df.columns:
        raise ValueError("ไม่พบคอลัมน์ 'url' ใน Excel")
    urls = df["url"].dropna().tolist()

    # ---------- STEP 4: วนเปิดแต่ละ URL ----------
    results = []

    for idx, url in enumerate(urls, start=1):
        logging.info(f"🔹 row {idx}: {url}")

        if not safe_get(url):
            results.append({
                "row": idx,
                "url": url,
                "span_text": "Timeout",
                "input_value_set": ""
            })
            continue

        time.sleep(2)

        # ---------- STEP 4.1: ตรวจ span ----------
        span_text = ""
        try:
            span_element = wait.until(EC.presence_of_element_located(
                (By.XPATH, "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[3]/div/div/div/span")
            ))
            span_text = span_element.text.strip()
            print(f"[Row {idx}] Span Text: {span_text}")
        except TimeoutException:
            print(f"[Row {idx}] ❌ ไม่พบ span element")

        # ---------- STEP 4.2: ถ้า span_text = "No SKU item created." ให้ใส่ค่าใน input ----------
        if span_text.strip().lower() == "no sku item created.":
            try:
                input_blocks = driver.find_elements(By.CSS_SELECTOR, "div.vendor-item-number-value > div > input")
                for i, input_field in enumerate(input_blocks, start=1):
                    input_field.clear()
                    time.sleep(1)
                    input_field.send_keys(NEW_VALUE)
                    print(f"[Row {idx}] 👉 ใส่ค่าใหม่ใน input block {i}: {NEW_VALUE}")
                    time.sleep(10)
            except Exception as e:
                print(f"[Row {idx}] ❌ ไม่สามารถใส่ค่าใน input ได้: {e}")

        # ---------- STEP 5: คลิกปุ่ม ----------
        try:
            btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[3]/div/button")
            ))
            btn.click()
            logging.info(f"[Row {idx}] ✅ กดปุ่มสำเร็จ")
        except TimeoutException:
            logging.warning(f"[Row {idx}] ⚠️ ไม่พบปุ่ม")

        # บันทึกผลลัพธ์
        results.append({
            "row": idx,
            "url": url,
            "span_text": span_text,
            "input_value_set": NEW_VALUE if span_text == "No SKU item created." else ""
        })

        time.sleep(1)

    # ---------- STEP 6: บันทึกผลลัพธ์กลับ Excel ----------
    results_df = pd.DataFrame(results)
    results_df.to_excel(OUTPUT_FILE, index=False)
    logging.info("✅ บันทึกผลลัพธ์เรียบร้อยแล้ว")

    driver.quit()

except Exception as e:
    logging.exception("เกิดข้อผิดพลาด:")
    driver.quit()
