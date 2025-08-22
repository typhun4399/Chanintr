import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import datetime

# --- ตั้งค่า ---
excel_path = r"C:\Users\tanapat\Desktop\missmatch.xlsx"
base_folder = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\14_KCS_done uploaded"
base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects?inv=1&invt=Ab5Vew&prefix=&forceOnObjectsSortingFiltering=false"

# โหลด id จาก Excel
df = pd.read_excel(excel_path)
ids = df['id'].dropna().astype(str).tolist()

# ตั้งค่า selenium
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--log-level=3")
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# --- ล็อกอิน Google ---
driver.get(
    "https://accounts.google.com/v3/signin/identifier?continue=https%3A%2F%2Fwww.google.com%2F%3Fhl%3Dth&ec=futura_exp_og_so_72776762_e&hl=th&ifkv=AdBytiMKqEEqMrt9R1KnjuPXmFPiu0nFjZWeUoSErpWjXrqtDmTX09g5-xwn-YhNshimTArNTEDcqA&passive=true&flowName=GlifWebSignIn&flowEntry=ServiceLogin&dsh=S-1123946678%3A1755283467407267"
)

email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
email_input.clear()
email_input.send_keys("tanapat@chanintr.com")
wait.until(EC.element_to_be_clickable((By.ID, "identifierNext"))).click()

password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
password_input.clear()
password_input.send_keys("Qwerty12345$$")
wait.until(EC.element_to_be_clickable((By.ID, "passwordNext"))).click()

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

    # อัปโหลดโฟลเดอร์ย่อย
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

            # --- dialog หลังอัปโหลด ---
            try:
                dialog_timeout = 30  # กันค้าง (วินาที)
                start_dialog = datetime.datetime.now()

                clicked_radio = False
                clicked_checkbox = False

                while (datetime.datetime.now() - start_dialog).seconds < dialog_timeout:
                    
                    # หา checkbox
                    checkboxes = driver.find_elements(By.XPATH, "//mat-dialog-container//mat-checkbox")
                    if checkboxes and not clicked_checkbox:
                        try:
                            ActionChains(driver).move_to_element(checkboxes[0]).click().perform()
                            print("เลือก checkbox แล้ว")
                            clicked_checkbox = True
                        except Exception:
                            pass

                    # ถ้าเลือกครบแล้ว -> ออกจาก loop
                    if clicked_radio or clicked_checkbox:
                        break

                    time.sleep(2)  # รอแล้ว retry

                # กด Continue Uploading (แบบ flexible)
                continue_btn = wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(), 'Continue Uploading')]")
                    )
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", continue_btn)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", continue_btn)
                print("กดปุ่ม Continue Uploading เรียบร้อย")

            except Exception as e:
                print(f"⚠️ ไม่สามารถกด radio/checkbox/Continue Uploading ได้: {e}")

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

print("เสร็จสิ้นการประมวลผลทั้งหมด")
driver.quit()