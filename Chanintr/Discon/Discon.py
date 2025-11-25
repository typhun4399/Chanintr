import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import logging

# --- CONFIG ---
excel_path = r"C:\Users\tanapat\Downloads\to check BAK & MGC discontinuation 090925.xlsx"
id_column_name = "id"
wait_time = 3

GOOGLE_EMAIL = "tanapat@chanintr.com"
GOOGLE_PASSWORD = "Qwerty12345$$"

# --- อ่าน Excel ---
df = pd.read_excel(excel_path)

# --- ตั้งค่า Chrome ---
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 15)

# --- ตั้งค่า logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- การล็อกอิน Google แบบ Hardcode ---
try:
    driver.get("https://accounts.google.com/signin/v2/identifier")
    
    email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
    email_input.clear()
    email_input.send_keys(GOOGLE_EMAIL)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='identifierNext']"))).click()
    
    password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
    password_input.clear()
    password_input.send_keys(GOOGLE_PASSWORD)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='passwordNext']"))).click()
    
    logging.info("ล็อกอินสำเร็จ... รอ 5 วินาทีเพื่อเข้าสู่ระบบ")
    time.sleep(5)
except TimeoutException:
    logging.error("ล็อกอินไม่สำเร็จ! ไม่พบช่องกรอกข้อมูลหรือหมดเวลา.")
    driver.quit()
    exit()

# --- เปิดหน้า base.chanintr.com/login และกดปุ่ม ---
try:
    driver.get("https://base.chanintr.com/login")
    button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/section/div[3]/button")))
    button.click()
    logging.info("กดปุ่มบนหน้า base.chanintr.com/login เรียบร้อยแล้ว")
    time.sleep(3)

    # กดปุ่มเลือก account
    second_login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div/div/div[2]/div/div/div[1]/form/span/section/div/div/div/div/ul/li[1]")))
    second_login_button.click()
    logging.info("กดปุ่มเลือก account สำเร็จแล้ว")
    time.sleep(5)

except TimeoutException:
    logging.error("ไม่พบปุ่มบนหน้า login หรือหมดเวลา")
    driver.quit()
    exit()
    
# --- วนลูปเปิด URL ตาม id ---
for idx, row in df.iterrows():
    product_id = row[id_column_name]
    url = f"https://base.chanintr.com/brand/10/product/{product_id}/overview"
    logging.info(f"Opening: {url}")
    driver.get(url)

    try:
        # กดปุ่มแรกใน product
        product_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/section/section/section[2]/div/div[2]/button")))
        product_button.click()
        logging.info(f"กดปุ่ม product {product_id} เรียบร้อยแล้ว")
        time.sleep(2)

        # กดปุ่มถัดไป
        second_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/section/section/div[1]/div/div[2]/div/div[2]/div[6]/div[2]/div/div[1]")))
        second_button.click()
        logging.info(f"กดปุ่มถัดไปใน product {product_id} เรียบร้อยแล้ว")
        time.sleep(2)

        # กด Discontinued (li[4])
        discontinued_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/section/section/div[1]/div/div[2]/div/div[2]/div[6]/div[2]/div/div[2]/div/div[2]/ul/li[4]")))
        discontinued_button.click()
        logging.info(f"กด Discontinued ใน product {product_id} เรียบร้อยแล้ว")
        time.sleep(2)

        # กดปุ่ม button[2]
        final_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/section/section/div[1]/div/div[2]/div/div[3]/button[2]")))
        final_button.click()
        logging.info(f"กดปุ่ม button[2] ใน product {product_id} เรียบร้อยแล้ว")

    except TimeoutException:
        logging.warning(f"ไม่พบปุ่มใน product {product_id}")

    time.sleep(wait_time)

# --- ปิด Browser ---
driver.quit()
