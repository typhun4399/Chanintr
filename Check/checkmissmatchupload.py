# -*- coding: utf-8 -*-
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import logging

# --- CONFIG ---
EXCEL_PATH = r"C:\Users\tanapat\Downloads\1_MUU active model id_27Jun25 - Done_updated.xlsx"
BASE_FOLDER = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\11_MUU_done uploaded"
BASE_URL = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects?inv=1&invt=Ab5Vew&prefix=&forceOnObjectsSortingFiltering=false"
GCS_TABLE_BODY_XPATH = "//cfc-table//table/tbody"
MAX_UPLOAD_WAIT_SECONDS = 3600  # 1 ชั่วโมง

# --- Login ---
GOOGLE_EMAIL = "tanapat@chanintr.com"
GOOGLE_PASSWORD = "Qwerty12345$$"

# --- Logging ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()]
)

# --- Helper Functions ---
def get_row_count(driver: webdriver.Chrome, table_xpath: str, timeout: int = 25) -> int:
    """นับจำนวนแถวในตาราง GCS โดยเลื่อนหน้าจอจนกว่าจำนวนจะแน่นอน"""
    wait_local = WebDriverWait(driver, timeout)
    try:
        table_body = wait_local.until(EC.presence_of_element_located((By.XPATH, table_xpath)))
        last_count, stable_count = -1, 0

        while stable_count < 3:
            rows = table_body.find_elements(By.TAG_NAME, "tr")
            current_count = len(rows)

            if current_count == last_count:
                stable_count += 1
            else:
                stable_count = 0
                last_count = current_count

            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", table_body)
            time.sleep(0.5)
        return last_count
    except TimeoutException:
        logging.warning("ไม่พบตารางข้อมูลในเวลา, คืนค่าเป็น 0")
        return 0


def _handle_upload_dialog(driver: webdriver.Chrome, wait: WebDriverWait, upload_type: str):
    """จัดการ popup หลังสั่งอัปโหลด"""
    try:
        clicked_checkbox = False

        start_time = time.time()
        while time.time() - start_time < 15:  # loop รอไม่เกิน 15 วินาที
            # หา checkbox
            checkboxes = driver.find_elements(By.XPATH, "//mat-dialog-container//mat-checkbox")
            if checkboxes and not clicked_checkbox:
                try:
                    ActionChains(driver).move_to_element(checkboxes[0]).click().perform()
                    logging.info("เลือก checkbox เรียบร้อยแล้ว")
                    clicked_checkbox = True
                except Exception:
                    pass
            if clicked_checkbox:
                    break
            time.sleep(1)

        # กด Continue Uploading
        continue_btn = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(), 'Continue Uploading')]")
            )
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", continue_btn)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", continue_btn)
        logging.info("กดปุ่ม Continue Uploading เรียบร้อย")

    except Exception as e:
        pass

    # ตรวจสอบ snackbar
    start_time = time.time()
    while time.time() - start_time < MAX_UPLOAD_WAIT_SECONDS:
        try:
            time.sleep(1)
            snackbar = driver.find_element(By.XPATH, "//mat-snack-bar-container")
            snackbar_text = snackbar.text.lower()
            if "successfully uploaded" in snackbar_text:
                logging.info(f"✅ อัปโหลดเสร็จสิ้น: {snackbar.text}")
                return
            elif "error" in snackbar_text or "failed" in snackbar_text:
                logging.error(f"❌ อัปโหลดล้มเหลว: {snackbar.text}")
                return
        except NoSuchElementException:
            pass
        time.sleep(2)

    logging.error("⚠️ Timeout: ไม่มี snackbar แจ้งสถานะอัปโหลด")


def upload_folder(folder_path: str, driver: webdriver.Chrome, wait: WebDriverWait):
    """อัปโหลดทั้งโฟลเดอร์"""
    logging.info(f"กำลังอัปโหลดโฟลเดอร์: {folder_path}")
    try:
        upload_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'][webkitdirectory]"))
        )
        upload_input.send_keys(folder_path)
        _handle_upload_dialog(driver, wait, upload_type='folder')
    except TimeoutException:
        logging.error("ไม่พบ input สำหรับอัปโหลดโฟลเดอร์")


def upload_files_only(folder_path: str, driver: webdriver.Chrome, wait: WebDriverWait):
    """อัปโหลดไฟล์ในโฟลเดอร์ (ไม่รวมโฟลเดอร์ย่อย)"""
    files = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
             if os.path.isfile(os.path.join(folder_path, f))]
    if not files:
        logging.warning(f"ไม่มีไฟล์ใน: {folder_path}")
        return

    logging.info(f"อัปโหลดไฟล์ {len(files)} ไฟล์ จาก: {folder_path}")
    try:
        upload_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'][multiple]"))
        )
        upload_input.send_keys("\n".join(files))
        _handle_upload_dialog(driver, wait, upload_type='files')
    except TimeoutException:
        logging.error("ไม่พบ input สำหรับอัปโหลดไฟล์")


def process_single_id(driver: webdriver.Chrome, wait: WebDriverWait, id_value: str, txt_file):
    """ตรวจสอบ ID เดียว แล้วอัปโหลดไฟล์ที่ขาด"""
    target_folders = [f for f in os.listdir(BASE_FOLDER)
                      if f.startswith(id_value) and os.path.isdir(os.path.join(BASE_FOLDER, f))]
    if not target_folders:
        logging.warning(f"ไม่พบโฟลเดอร์สำหรับ ID {id_value}, บันทึกลงไฟล์ mismatch")
        txt_file.write(f"{id_value}\n")
        return

    folder_path = os.path.join(BASE_FOLDER, target_folders[0])

    # --- START: ตรวจสอบโฟลเดอร์ 2D, 3D, Datasheet ก่อนอัปโหลด ---
    required_folders = ["2D", "3D", "Datasheet"]
    has_files_to_upload = False
    for subfolder_name in required_folders:
        subfolder_path = os.path.join(folder_path, subfolder_name)
        if os.path.isdir(subfolder_path):
            if any(os.path.isfile(os.path.join(subfolder_path, f)) for f in os.listdir(subfolder_path)):
                has_files_to_upload = True
                break

    if not has_files_to_upload:
        logging.info(f"ID {id_value}: ข้ามการทำงาน เนื่องจากโฟลเดอร์ 2D, 3D และ Datasheet ไม่มีไฟล์อยู่เลย")
        return
    # --- END: สิ้นสุดการตรวจสอบ ---

    local_subfolders = [f for f in os.listdir(folder_path)
                        if os.path.isdir(os.path.join(folder_path, f))]
    logging.info(f"ID {id_value}: พบ {len(local_subfolders)} โฟลเดอร์ย่อย, เริ่มการซิงค์...")

    url = BASE_URL.format(id_value)
    driver.get(url)

    try:
        wait.until(EC.presence_of_element_located((By.XPATH, GCS_TABLE_BODY_XPATH)))

        for subfolder_name in local_subfolders:
            subfolder_path = os.path.join(folder_path, subfolder_name)
            local_file_count = len([f for f in os.listdir(subfolder_path)
                                    if os.path.isfile(os.path.join(subfolder_path, f))])
            
            # ถ้าโฟลเดอร์ใน local ไม่มีไฟล์เลย ให้ข้ามไปโฟลเดอร์ถัดไป
            if local_file_count == 0:
                logging.info(f"  -> '{subfolder_name}' ใน local ไม่มีไฟล์, ข้ามการตรวจสอบ")
                continue

            try:
                link_xpath = f"//cfc-table//table//a[contains(text(), '{subfolder_name}')]"
                link_element = driver.find_element(By.XPATH, link_xpath)

                logging.info(f"  -> '{subfolder_name}' มีบน GCS, ตรวจสอบจำนวนไฟล์...")
                link_element.click()

                gcs_file_count = get_row_count(driver, GCS_TABLE_BODY_XPATH, timeout=25)
                logging.info(f"     '{subfolder_name}': local={local_file_count}, GCS={gcs_file_count}")

                if gcs_file_count != local_file_count:
                    logging.warning(f"     จำนวนไฟล์ไม่ตรง! อัปโหลดใหม่...")
                    upload_files_only(subfolder_path, driver, wait)

                driver.back()
                wait.until(EC.presence_of_element_located((By.XPATH, GCS_TABLE_BODY_XPATH)))
                time.sleep(1)

            except NoSuchElementException:
                logging.warning(f"  -> '{subfolder_name}' ไม่พบใน GCS, อัปโหลดใหม่ทั้งโฟลเดอร์")
                upload_folder(subfolder_path, driver, wait)

    except TimeoutException:
        logging.error(f"โหลดตาราง GCS ของ ID {id_value} ไม่สำเร็จ")


def main():
    """ฟังก์ชันหลัก"""
    try:
        df = pd.read_excel(EXCEL_PATH)
        ids = df['id'].dropna().astype(int).astype(str).tolist()
        logging.info(f"โหลด {len(ids)} IDs จาก Excel")
    except FileNotFoundError:
        logging.error(f"ไม่พบไฟล์ Excel: {EXCEL_PATH}")
        return

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    # Login Google
    driver.get("https://accounts.google.com/signin/v2/identifier")
    try:
        email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
        email_input.send_keys(GOOGLE_EMAIL)
        wait.until(EC.element_to_be_clickable((By.ID, "identifierNext"))).click()

        password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
        password_input.send_keys(GOOGLE_PASSWORD)
        wait.until(EC.element_to_be_clickable((By.ID, "passwordNext"))).click()

        logging.info("ล็อกอิน Google สำเร็จ")
        time.sleep(5)
    except TimeoutException:
        logging.error("ล็อกอิน Google ล้มเหลว")
        driver.quit()
        return

    logging.info("=== เสร็จสิ้น ===")
    driver.quit()


if __name__ == "__main__":
    main()

