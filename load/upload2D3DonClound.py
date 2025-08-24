import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime

# --- ตั้งค่า ---
excel_path = r"C:\Users\tanapat\Desktop\topBAK.xlsx"
base_folder = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\4_BAK"
base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects?inv=1&invt=Ab5Vew&prefix=&forceOnObjectsSortingFiltering=false"

# โหลด id จาก Excel
df = pd.read_excel(excel_path)
ids = df['id'].dropna().astype(str).tolist()

# ตั้งค่า selenium
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# --- ล็อกอิน Google ---
driver.get(
    "https://accounts.google.com/v3/signin/accountchooser?continue=https%3A%2F%2Fwww.google.com%2F&ec=futura_exp_og_so_72776762_e&hl=en&ifkv=AdBytiMGIAttUZ1i34PUHaH6pAPdQX_zyXswMLfvRU4wKbXzwS2dw7_xaUs92mDAWPzAUA_TWbDLoA&passive=true&flowName=GlifWebSignIn&flowEntry=ServiceLogin&dsh=S1812811471%3A1755056943534734"
)

email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
email_input.clear()
email_input.send_keys("tanapat@chanintr.com")

next_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='identifierNext']")))
next_btn.click()

password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
password_input.clear()
password_input.send_keys("Qwerty12345$$")

password_next = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='passwordNext']")))
password_next.click()

time.sleep(5)  # รอโหลดหน้า GCS

# --- เริ่มลูปตาม id ---
for id_value in ids:
    # หาโฟลเดอร์หลัก
    target_folders = [
        f
        for f in os.listdir(base_folder)
        if f.startswith(id_value) and os.path.isdir(os.path.join(base_folder, f))
    ]
    if not target_folders:
        print(f"ไม่พบโฟลเดอร์หลักสำหรับ id {id_value}")
        continue

    # เปิดเว็บ GCS ตาม id
    url = base_url.format(id_value)
    print(f"เปิด URL: {url}")
    driver.get(url)
    time.sleep(5)  # รอโหลดหน้า

    # ตรวจสอบข้อความ "No rows to display"
    try:
        no_rows_xpath = "/html/body/div[1]/div[3]/div[3]/div/pan-shell/pcc-shell/cfc-panel-container/div/div/cfc-panel/div/div/div[3]/cfc-panel-container/div/div/cfc-panel/div/div/cfc-panel-container/div/div/cfc-panel/div/div/cfc-panel-container/div/div/cfc-panel[2]/div/div/central-page-area/div/div/pcc-content-viewport/div/div/pangolin-home-wrapper/pangolin-home/cfc-router-outlet/div/ng-component/cfc-single-panel-layout/cfc-panel-container/div/div/cfc-panel/div/div/cfc-panel-body/cfc-virtual-viewport/div[1]/div/mat-tab-group/div/mat-tab-body[1]/div/storage-bucket-details-objects/cfc-left-panel-layout/cfc-panel-container/div/div/cfc-panel[2]/div/div/cfc-main-panel-content/cfc-panel-body/cfc-virtual-viewport/div[1]/div/storage-objects-table/storage-drop-target/div[2]/cfc-table/div[2]/cfc-table-columns-presenter-v2/div/div[3]/table/tbody/tr/td"
        cell = wait.until(EC.presence_of_element_located((By.XPATH, no_rows_xpath)))
        cell_text = cell.text.strip()
    except Exception as e:
        print(f"ไม่พบ element เช็คข้อความ หรือเกิดข้อผิดพลาด: {e}")
        cell_text = ""

    if cell_text == "No rows to display":
        print(f"Bucket id {id_value} ยังไม่มีไฟล์, เริ่มอัปโหลดได้")
        # เริ่มอัปโหลดโฟลเดอร์ย่อยตามที่มีในเครื่อง
        for folder_name in target_folders:
            folder_path = os.path.join(base_folder, folder_name)
            subfolders = [
                sf
                for sf in os.listdir(folder_path)
                if os.path.isdir(os.path.join(folder_path, sf))
            ]
            if not subfolders:
                print(f"ไม่มีโฟลเดอร์ย่อยใน {folder_path}")
                continue

            for sf in subfolders:
                full_path = os.path.join(folder_path, sf)
                print(f"กำลังอัปโหลดโฟลเดอร์: {full_path}")

                # เลือก input สำหรับอัปโหลด
                upload_input = wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "input[type='file'][webkitdirectory]")
                    )
                )
                upload_input.send_keys(full_path)

                # --- รอข้อความ success (ไม่จำกัดเวลา + กัน loop ค้าง) ---
            success_xpath = "//mat-snack-bar-container//div[contains(text(),'successfully uploaded')]"
            start_time = datetime.datetime.now()

            while True:
                try:
                    msg = driver.find_element(By.XPATH, success_xpath).text.strip().lower()
                    if "successfully uploaded" in msg:
                        print(f"✅ อัปโหลดโฟลเดอร์ {sf} เสร็จสมบูรณ์")
                        break
                except:
                    pass

                time.sleep(2)  # เช็คทุก 2 วิ

                # กัน loop ค้างเกิน 1 ชั่วโมง
                if (datetime.datetime.now() - start_time).seconds > 3600:
                    print(f"⚠️ อัปโหลดโฟลเดอร์ {sf} ใช้เวลาเกิน 1 ชั่วโมง หยุดรอ")
                    break

    else:
        print(f"Bucket id {id_value} มีไฟล์อยู่แล้ว ({cell_text}) ไม่ต้องอัปโหลด")

print("เสร็จสิ้นการประมวลผลทั้งหมด")
driver.quit()
