import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import shutil
import logging # 👈 เพิ่มเข้ามา
from datetime import datetime # 👈 เพิ่มเข้ามา

# --- ‼️ ส่วนของการตั้งค่า Logger (เพิ่มเข้ามาใหม่ทั้งหมด) ‼️ ---
# สร้างชื่อไฟล์ log ที่ไม่ซ้ำกันตามวันเวลาที่รัน
log_filename = f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"

# สร้าง logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# สร้าง formatter เพื่อกำหนดรูปแบบของ log
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

# สร้าง File Handler เพื่อบันทึก log ลงไฟล์
file_handler = logging.FileHandler(log_filename, encoding='utf-8')
file_handler.setFormatter(formatter)

# สร้าง Console Handler เพื่อแสดง log บนหน้าจอ
console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)

# เพิ่ม handler ทั้งสองเข้าไปใน logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)


# --- ฟังก์ชันหลัก: สำหรับประมวลผล 1 Product Number ---
def process_one_product(driver, wait, product_number, excel_title, destination_folder_path, download_dir):
    try:
        search_url = f"https://www.bakerfurniture.com/search-results?q={product_number}"
        logger.info(f"> กำลังค้นหา Product Number: {product_number}")
        driver.get(search_url)
        time.sleep(3)

        results_li_xpath = "/html/body/div[1]/div[2]/div[3]/div/div/div/div[2]/div/div[2]/ul/li"
        search_results_list = driver.find_elements(By.XPATH, results_li_xpath)
        logger.info(f"  └── พบผลการค้นหา {len(search_results_list)} รายการ")
        best_matches = []
        max_score = -1
        for result_li in search_results_list:
            try:
                num_element = result_li.find_element(By.XPATH, ".//div[1]/div[2]/div[4]/a")
                number_only = num_element.text.replace("No. ", "")
                if number_only == str(product_number):
                    title_element = result_li.find_element(By.XPATH, ".//div[1]/div[2]/div[3]/a")
                    web_title = title_element.text
                    excel_title_words = set(str(excel_title).lower().split())
                    web_title_words = set(web_title.lower().split())
                    score = len(excel_title_words.intersection(web_title_words))
                    candidate_info = {'element': num_element, 'title': web_title, 'number': number_only}
                    if score > max_score:
                        max_score = score
                        best_matches = [candidate_info]
                    elif score == max_score and max_score != -1:
                        best_matches.append(candidate_info)
            except Exception:
                continue
        
        if best_matches:
            if len(best_matches) > 1:
                logger.warning(f"  -> !! คำเตือน: พบ Best Match คะแนนสูงสุด ({max_score}) เท่ากัน {len(best_matches)} รายการ:")
                for item in best_matches: logger.warning(f"     - รหัส: {item['number']}, ชื่อ: {item['title']}")
            
            winner = best_matches[0]
            logger.info(f"  -> กำลังคลิก Best Match (รายการแรก): '{winner['number']}' (คะแนน {max_score})")
            winner['element'].click()

            # --- เริ่มกระบวนการดาวน์โหลดปุ่มที่ 1 ---
            try:
                logger.info("  --- [ขั้นตอน 1/2]: กำลังดาวน์โหลด Tearsheet/Print File ---")
                time.sleep(1)
                button_pdf_xpath = ".btn.mix-btn_blocked"
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, button_pdf_xpath))).click()
                time.sleep(1)
                button_pdf_load_xpath = "//button[contains(text(), 'Download PDF')]"
                button_pdfload = wait.until(EC.element_to_be_clickable((By.XPATH, button_pdf_load_xpath)))
                files_before = os.listdir(download_dir)
                button_pdfload.click()
                logger.info("  -> คลิกปุ่ม 'Download PDF' แล้ว")
                
                timeout, start_time, is_complete, has_retried = 600, time.time(), False, False
                while time.time() - start_time < timeout:
                    elapsed = time.time() - start_time
                    if elapsed > 10 and not has_retried and len(os.listdir(download_dir)) == len(files_before):
                        logger.warning("  -> !! ไม่พบการเริ่มดาวน์โหลด, ลองคลิกอีกครั้ง...")
                        button_pdfload.click()
                        has_retried = True
                    is_downloading = any([f.endswith((".crdownload", ".tmp")) for f in os.listdir(download_dir)])
                    if not is_downloading:
                        files_after = os.listdir(download_dir)
                        if len(files_after) > len(files_before):
                            is_complete = True
                            break
                        time.sleep(1)
                if is_complete:
                    new_file = list(set(files_after) - set(files_before))[0]
                    logger.info(f"  -> ดาวน์โหลดเสร็จสิ้น! ไฟล์ใหม่คือ: {new_file}")
                    source_path = os.path.join(download_dir, new_file)
                    os.makedirs(destination_folder_path, exist_ok=True)
                    shutil.move(source_path, destination_folder_path)
                    logger.info(f"  -> ย้ายไฟล์ '{new_file}' ไปที่ '{destination_folder_path}' เรียบร้อยแล้ว")
                    logger.info("  -> กำลังปิดหน้าต่างดาวน์โหลด...")
                    time.sleep(1)
                    close_popup_button = "//button[contains(text(), 'Cancel')]"
                    wait.until(EC.element_to_be_clickable((By.XPATH, close_popup_button))).click()
                    time.sleep(1)
                else: logger.error("  -> การดาวน์โหลด Tearsheet/Print ไม่เสร็จสิ้นภายในเวลาที่กำหนด")
            except Exception: logger.exception("  -> !! เกิดข้อผิดพลาดกับปุ่ม Tearsheet/Print")

            # --- เริ่มกระบวนการดาวน์โหลดปุ่มที่ 2 ---
            try:
                time.sleep(1)
                logger.info("  --- [ขั้นตอน 2/2]: กำลังดาวน์โหลด Project File ---")
                button_2_xpath = "//a[contains(text(), 'Download Design Package')]"
                wait.until(EC.element_to_be_clickable((By.XPATH, button_2_xpath))).click()
                files_before = os.listdir(download_dir)
                logger.info("  -> คลิกปุ่ม 'ADD TO PROJECT' แล้ว")
                
                timeout, start_time, is_complete, has_retried = 600, time.time(), False, False
                while True:
                    is_downloading = any([f.endswith((".crdownload", ".tmp")) for f in os.listdir(download_dir)])
                    if not is_downloading:
                        files_after = os.listdir(download_dir)
                        if len(files_after) >= len(files_before):
                            is_complete = True
                            break                   
                if is_complete:
                    new_file = list(set(files_after) - set(files_before))[0]
                    logger.info(f"  -> ดาวน์โหลดเสร็จสิ้น! ไฟล์ใหม่คือ: {new_file}")
                    source_path = os.path.join(download_dir, new_file)
                    os.makedirs(destination_folder_path, exist_ok=True)
                    shutil.move(source_path, destination_folder_path)
                    logger.info(f"  -> ย้ายไฟล์ '{new_file}' ไปที่ '{destination_folder_path}' เรียบร้อยแล้ว")
                else: logger.error("  -> การดาวน์โหลด 'ADD TO PROJECT' ไม่เสร็จสิ้นภายในเวลาที่กำหนด")
            except Exception: logger.exception("  -> !! เกิดข้อผิดพลาดกับปุ่ม 'ADD TO PROJECT'")
        else:
            logger.warning("-> ไม่พบรายการที่ตรงกับรหัสสินค้าบนหน้านี้")
    except Exception:
        logger.exception(f"-> !! เกิดข้อผิดพลาดร้ายแรงระหว่างประมวลผล Product Number {product_number}")

# --- ส่วนของการตั้งค่า (Options) และโฟลเดอร์ดาวน์โหลด ---
download_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(download_dir): os.makedirs(download_dir)
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
prefs = {"download.default_directory": download_dir, "download.prompt_for_download": False, "download.directory_upgrade": True}
chrome_options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# --- ‼️ แก้ไขตรงนี้: ระบุข้อมูลไฟล์ Excel และชื่อคอลัมน์ ---
excel_file_path = r"C:\Users\tanapat\Downloads\1_BAK active model id_25Jun25 (3).xlsx"
vendor_column_names = ["product_vendor_number1", "product_vendor_number2", "product_vendor_number3"]
column_name_with_titles = "product_title"
column_name_with_id = "id"

try:
    logger.info("===== เริ่มต้นการทำงานของสคริปต์ =====")
    driver.get("https://www.bakerfurniture.com/")
    logger.info("กำลังเปิดเว็บ...")
    account_button_xpath = "/html/body/div[1]/div[1]/div[2]/div/div[2]/ul/li[3]/a"
    wait.until(EC.element_to_be_clickable((By.XPATH, account_button_xpath))).click()
    logger.info("กำลังล็อกอิน...")
    email_input_xpath = "/html/body/div[1]/div[3]/div[1]/div/div/div[2]/div/div[1]/div/div[2]/form/div[4]/div[2]/input"
    wait.until(EC.visibility_of_element_located((By.XPATH, email_input_xpath))).send_keys("info@chanintr.com")
    password_input_xpath = "/html/body/div[1]/div[3]/div[1]/div/div/div[2]/div/div[1]/div/div[2]/form/div[5]/div[2]/input"
    wait.until(EC.visibility_of_element_located((By.XPATH, password_input_xpath))).send_keys("clv123456")
    signin_button_xpath = "/html/body/div[1]/div[3]/div[1]/div/div/div[2]/div/div[1]/div/div[2]/form/div[7]/button"
    wait.until(EC.element_to_be_clickable((By.XPATH, signin_button_xpath))).click()
    logger.info("ทำการ Sign In เรียบร้อยแล้ว")
    time.sleep(3)

    logger.info(f"กำลังอ่านข้อมูลจากไฟล์: {excel_file_path}")
    df = pd.read_excel(excel_file_path)
    
    for index, row in df.iterrows():
        excel_title = row[column_name_with_titles]
        folder_id = row[column_name_with_id]
        valid_vendor_numbers = []
        for col_name in vendor_column_names:
            if col_name in row and pd.notna(row[col_name]):
                valid_vendor_numbers.append(row[col_name])
        
        if not valid_vendor_numbers: continue

        logger.info(f"===== เริ่มทำงานกับ ID: {folder_id} (พบ {len(valid_vendor_numbers)} vendor number) =====")

        for vendor_num in valid_vendor_numbers:
            if len(valid_vendor_numbers) > 1:
                destination_path = os.path.join(r"D:\BAK\2D&3D", str(folder_id), str(vendor_num))
            else:
                destination_path = os.path.join(r"D:\BAK\2D&3D", str(folder_id))
            
            process_one_product(driver, wait, vendor_num, excel_title, destination_path, download_dir)
            time.sleep(3)
except Exception:
    logger.exception("เกิดข้อผิดพลาดร้ายแรงที่ไม่สามารถจัดการได้ใน Loop หลัก")
finally:
    logger.info("===== ทำงานเสร็จสิ้นทั้งหมดแล้ว =====")
    time.sleep(10)
    driver.quit()