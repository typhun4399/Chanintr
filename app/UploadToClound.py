import os
import pandas as pd
import time
import datetime
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

log_desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
log_file_path = os.path.join(log_desktop_path, 'upload_automation_log.txt')

def get_inputs():
    def browse_excel():
        path = filedialog.askopenfilename(
            filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, path)

    def browse_folder():
        path = filedialog.askdirectory()
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, path)

    def submit():
        nonlocal email, password, excel_path, folder_path
        # --- คงข้อมูล Hardcode ไว้ตามที่ผู้ใช้ต้องการ ---
        email = email_entry
        password = password_entry
        excel_path = excel_entry
        folder_path = folder_entry
        root.destroy()

    email = password = excel_path = folder_path = ""

    root = tk.Tk()
    root.title("ตั้งค่า Selenium Upload")

    tk.Label(root, text="Email:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    email_entry = tk.Entry(root, width=40)
    email_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(root, text="Password:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    password_entry = tk.Entry(root, width=40, show="*")
    password_entry.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(root, text="Excel/CSV Path:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    excel_entry = tk.Entry(root, width=40)
    excel_entry.grid(row=2, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=browse_excel).grid(row=2, column=2, padx=5, pady=5)

    tk.Label(root, text="Folder Path:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
    folder_entry = tk.Entry(root, width=40)
    folder_entry.grid(row=3, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=browse_folder).grid(row=3, column=2, padx=5, pady=5)

    tk.Button(root, text="เริ่มทำงาน", command=submit).grid(row=4, column=0, columnspan=3, pady=10)

    root.mainloop()
    return email, password, excel_path, folder_path

email, password, excel_path, base_folder = get_inputs()

if not all([email, password, excel_path, base_folder]):
    print("❌ใส่ข้อมูลไม่ครบถ้วน")
    exit()

if excel_path.lower().endswith(".csv"):
    df = pd.read_csv(excel_path)
else:
    df = pd.read_excel(excel_path)
ids = df['id'].dropna().astype(int).astype(str).tolist()

# ตั้งค่า Selenium
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# ล็อกอิน Google
try:
    driver.get("https://accounts.google.com/v3/signin/identifier?continue=https%3A%2F%2Fwww.google.com%2F%3Fhl%3Dth&ec=futura_exp_og_so_72776762_e&hl=th&ifkv=AdBytiPWkA-lXJsnK3T4TFbRSkqmZxItIQFbyepCsUhuk_btQR3u5Qa1JFOnV4NX_lT1FiQ7KM9JyQ&passive=true&flowName=GlifWebSignIn&flowEntry=ServiceLogin&dsh=S539092489%3A1755489428343508")
    email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
    email_input.clear()
    email_input.send_keys(email)
    driver.find_element(By.ID, "identifierNext").click()
    password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
    password_input.clear()
    password_input.send_keys(password)
    driver.find_element(By.ID, "passwordNext").click()
    print("✅ ล็อกอินสำเร็จ")
    time.sleep(5)
except Exception as e:
    print(f"❌ เกิดข้อผิดพลาดระหว่างการล็อกอิน: {e}")
    driver.quit()
    exit()

base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects"
for id_value in ids:
    
    log_messages_for_this_id = []

    target_folders = [f for f in os.listdir(base_folder) if f.split('_')[0] == id_value and os.path.isdir(os.path.join(base_folder, f))]
    if not target_folders:
        message = f"❌ ไม่พบโฟลเดอร์หลักสำหรับ id {id_value}"
        print(message)
        log_messages_for_this_id.append(message)
        continue

    url = base_url.format(id_value)
    message = f"\n🌐 เปิด URL: {url}"
    print(message)
    log_messages_for_this_id.append(message)
    driver.get(url)
    
    time.sleep(5) 

    try:
        no_rows_xpath = "//td[contains(text(),'No rows to display')]"
        cell = wait.until(EC.presence_of_element_located((By.XPATH, no_rows_xpath)))
        cell_text = cell.text.strip()
    except:
        cell_text = ""

    if cell_text == "No rows to display":
        message = f"📂 Bucket {id_value} ว่าง เริ่มอัปโหลด"
        print(message)
        log_messages_for_this_id.append(message)

        for folder_name in target_folders:
            folder_path = os.path.join(base_folder, folder_name)
            subfolders = [sf for sf in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, sf))]
            
            if not subfolders:
                message = f"⚠️ ไม่มีโฟลเดอร์ย่อยใน {folder_path}"
                print(message)
                log_messages_for_this_id.append(message) 
                continue

            for sf in subfolders:
                time.sleep(1)
                full_path = os.path.join(folder_path, sf)
                message = f"⬆️ กำลังอัปโหลด: {full_path}"
                print(message)
                log_messages_for_this_id.append(message)

                try:
                    upload_input = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'][webkitdirectory]"))
                    )
                    upload_input.send_keys(full_path)
                    
                    # --- ⚙️ ตรรกะการรอแบบใหม่ที่ฉลาดขึ้น ---
                    started_xpath = "//div[contains(text(),'Upload started')]"
                    success_xpath = "//div[contains(text(),'successfully uploaded')]"

                    try:
                        # 1. รอให้ข้อความ "Upload started" แสดงขึ้นมาก่อน (รอไม่เกิน 10 วินาที)
                        upload_started_wait = WebDriverWait(driver, 20)
                        upload_started_wait.until(EC.presence_of_element_located((By.XPATH, started_xpath)))
                        
                        # 2. เมื่อการอัปโหลดเริ่มแล้ว ให้รอจนกว่าข้อความ "Upload started" จะ "หายไป" (รอสูงสุด 30 นาที)
                        upload_finished_wait = WebDriverWait(driver, 1800)
                        upload_finished_wait.until(EC.invisibility_of_element_located((By.XPATH, started_xpath)))
                        time.sleep(2)

                        try:
                            upload_sucsess_wait = WebDriverWait(driver, 20)
                            upload_sucsess_wait.until(EC.presence_of_element_located((By.XPATH, success_xpath)))
                            message = f"✅ อัปโหลดโฟลเดอร์ {sf} เสร็จแล้ว"
                            print(message)
                            log_messages_for_this_id.append(message)
                            time.sleep(2)

                        except:
                            message = f"🟡 โฟลเดอร์ '{sf}' ว่างหรือไม่สำเร็จ (ไม่มีข้อความยืนยัน) กำลังข้าม..."
                            print(message)
                            log_messages_for_this_id.append(message)

                    except TimeoutException:
                        message = f"🟡 โฟลเดอร์ '{sf}' ว่าง (ไม่พบการแจ้งเตือนเริ่มต้น) กำลังข้าม..."
                        print(message)
                        log_messages_for_this_id.append(message)

                except Exception as e:
                    message = f"❌ เกิดข้อผิดพลาดระหว่างอัปโหลด '{sf}': {e}"
                    print(message)
                    log_messages_for_this_id.append(message)

    else:
        message = f"⏩ Bucket {id_value} มีไฟล์อยู่แล้ว"
        print(message)
        log_messages_for_this_id.append(message)

    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"--- [END OF ROW] ID: {id_value} | Time: {timestamp} ---\n")
        for msg in log_messages_for_this_id:
            log_file.write(msg.strip() + "\n")
        log_file.write("-" * 70 + "\n\n")

print("\n🎉 การทำงานทั้งหมดเสร็จสิ้น")
driver.quit()