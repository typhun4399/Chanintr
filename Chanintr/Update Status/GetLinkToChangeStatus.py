# -*- coding: utf-8 -*-
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
GOOGLE_EMAIL = "tanapat@chanintr.com"
GOOGLE_PASSWORD = "Qwerty12345$$"
OUTPUT_FILE = r"C:\Users\tanapat\Desktop\base_products.xlsx"
link_prod = "https://base.chanintr.com/brand/95/products?currentPage=1&directionUser=DESC&sortBy=title&direction=ASC&isSearch=false"

# ---------------- Chrome Options ----------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

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
    driver.get("https://base.chanintr.com/login")
    google_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Sign in with Google')]")))
    google_btn.click()
    time.sleep(5)

    # ---------- STEP 3: เปิด target page ----------
    driver.get(link_prod)
    time.sleep(3)

    # ---------- STEP 4: ฟังก์ชันดึงสินค้าจากหน้า ----------
    def extract_page_data(max_retries=5):
        page_data = []
        for attempt in range(max_retries):
            try:
                items = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "section.wrapper-container ul li a"))
                )
                page_data.clear()
                for a in items:
                    # ดึงชื่อสินค้า
                    try:
                        name = a.find_element(By.CSS_SELECTOR, "section div.product-title-container h3").text.strip()
                    except:
                        name = ""

                    # ดึงลิงก์สินค้า
                    href = a.get_attribute("href")

                    # ดึง extra info
                    try:
                        extra_info = a.find_element(By.CSS_SELECTOR, "section div.cell-md > div > div").text.strip()
                    except:
                        extra_info = ""

                    if name:
                        page_data.append({
                            "name": name,
                            "url": href,
                            "status": extra_info
                        })
                        logging.info(f"🔹 เจอสินค้า: {name} | {href} | Status : {extra_info}")

                if len(page_data) > 0:
                    return page_data
                else:
                    logging.warning(f"⚠️ ยังไม่มีสินค้าบนหน้านี้ ลองรอบที่ {attempt+1}")
                    time.sleep(2)
            except TimeoutException:
                logging.warning(f"⏳ รอสินค้าไม่เจอครั้งที่ {attempt+1}")
                time.sleep(2)
        logging.error("❌ ไม่พบสินค้าใด ๆ หลังรอหลายรอบ")
        return page_data

    # ---------- STEP 5: วนทุกหน้า ----------
    all_data = []
    page = 1

    while True:
        logging.info(f"📄 กำลังดึงข้อมูลหน้า {page} ...")
        page_links = extract_page_data()
        while len(page_links) == 0:
            logging.info("🔄 จำนวนสินค้า=0, รอและลองดึงใหม่")
            time.sleep(2)
            page_links = extract_page_data()
        logging.info(f"จำนวนสินค้าในหน้านี้: {len(page_links)}")
        all_data.extend(page_links)

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
                time.sleep(2)
            else:
                logging.info("✅ ไม่มีหน้าถัดไปแล้ว — หยุดลูป")
                break
        except Exception:
            logging.info("✅ ไม่มีปุ่ม Next — หยุดลูป")
            break

    # ---------- STEP 6: Export Excel ----------
    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel(OUTPUT_FILE, index=False)
        logging.info(f"✅ บันทึกข้อมูลลง Excel เสร็จ: {OUTPUT_FILE}")
    else:
        logging.warning("⚠️ ไม่พบสินค้าใด ๆ สำหรับบันทึก")

    driver.quit()

except Exception:
    logging.exception("เกิดข้อผิดพลาด:")
    try:
        driver.quit()
    except:
        pass
    raise
