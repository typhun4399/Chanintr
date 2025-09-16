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

# --- เพิ่มโค้ดส่วนที่ 1: กำหนดที่อยู่และชื่อไฟล์ Log ---
log_desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
log_file_path = os.path.join(log_desktop_path, 'upload_automation_log.txt')
# ----------------------------------------------------


# ------------------- ฟังก์ชัน GUI -------------------
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
        email = email_entry.get()
        password = password_entry.get()
        excel_path = excel_entry.get()
        folder_path = folder_entry.get()
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


# ------------------- เริ่มโปรแกรมหลัก -------------------
email, password, excel_path, base_folder = get_inputs()

# โหลด id จาก Excel หรือ CSV
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
driver.get("https://accounts.google.com/v3/signin/identifier?continue=https%3A%2F%2Fwww.google.com%2F%3Fhl%3Dth&ec=futura_exp_og_so_72776762_e&hl=th&ifkv=AdBytiPWkA-lXJsnK3T4TFbRSkqmZxItIQFbyepCsUhuk_btQR3u5Qa1JFOnV4NX_lT1FiQ7KM9JyQ&passive=true&flowName=GlifWebSignIn&flowEntry=ServiceLogin&dsh=S539092489%3A1755489428343508")
email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
email_input.clear()
email_input.send_keys(email)

driver.find_element(By.ID, "identifierNext").click()

password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
password_input.clear()
password_input.send_keys(password)

driver.find_element(By.ID, "passwordNext").click()

time.sleep(5)  # รอโหลดหน้า GCS

# เริ่มลูปตาม id
base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects"
for id_value in ids:
    # --- เพิ่มโค้ดส่วนที่ 2: สร้าง List เพื่อเก็บ Log ของแถวนี้ ---
    log_messages_for_this_id = []
    # --------------------------------------------------------

    target_folders = [f for f in os.listdir(base_folder) if f.split('_')[0] == id_value and os.path.isdir(os.path.join(base_folder, f))]
    if not target_folders:
        message = f"❌ ไม่พบโฟลเดอร์หลักสำหรับ id {id_value}"
        print(message)
        log_messages_for_this_id.append(message) # แก้ไข print เดิม
        continue

    # เปิด GCS
    url = base_url.format(id_value)
    message = f"\n🌐 เปิด URL: {url}"
    print(message)
    log_messages_for_this_id.append(message) # แก้ไข print เดิม
    driver.get(url)
    time.sleep(5)

    # เช็ค "No rows to display"
    try:
        no_rows_xpath = "//td[contains(text(),'No rows to display')]"
        cell = wait.until(EC.presence_of_element_located((By.XPATH, no_rows_xpath)))
        cell_text = cell.text.strip()
    except:
        cell_text = ""

    if cell_text == "No rows to display":
        message = f"📂 Bucket {id_value} ว่าง เริ่มอัปโหลด"
        print(message)
        log_messages_for_this_id.append(message) # แก้ไข print เดิม

        for folder_name in target_folders:
            folder_path = os.path.join(base_folder, folder_name)
            subfolders = [sf for sf in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, sf))]
            if not subfolders:
                message = f"⚠️ ไม่มีโฟลเดอร์ย่อยใน {folder_path}"
                print(message)
                log_messages_for_this_id.append(message) # แก้ไข print เดิม
                continue

            for sf in subfolders:
                time.sleep(1)
                full_path = os.path.join(folder_path, sf)
                message = f"⬆️ กำลังอัปโหลด: {full_path}"
                print(message)
                log_messages_for_this_id.append(message)

                # เลือกปุ่ม input[type=file][webkitdirectory]
                upload_input = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'][webkitdirectory]"))
                )
                upload_input.send_keys(full_path)

                success_xpath = "//mat-snack-bar-container//div[contains(text(),'successfully uploaded')]"
                started_xpath = "//mat-snack-bar-container//div[contains(text(),'Upload started')]"
                start_time = datetime.datetime.now()
                
                try:
                    # ใช้ WebDriverWait ที่มี timeout สั้นๆ (5 วินาที) เพื่อเช็คโดยเฉพาะ
                    short_wait = WebDriverWait(driver, 5)
                    short_wait.until(EC.presence_of_element_located((By.XPATH, started_xpath)))
                    # ถ้าเจอ popup ก็ให้ทำงานใน while loop ต่อไปตามปกติ
                except Exception:
                    # ถ้าไม่เจอ popup ภายใน 5 วินาที (เกิด TimeoutException)
                    message = f"🟡 ไม่พบการแจ้งเตือนเริ่มต้นอัปโหลดสำหรับ '{sf}' (อาจเป็นโฟลเดอร์ว่าง) กำลังข้าม..."
                    print(message)
                    log_messages_for_this_id.append(message)
                    time.sleep(1) 
                    continue # ข้ามไปยัง subfolder ถัดไปในลูปทันที
                # ========== END: โค้ดส่วนที่แก้ไข ==========
                
                upload_successful = False
                while True:
                    try:
                        msg = driver.find_element(By.XPATH, success_xpath).text.strip().lower()
                        if "successfully uploaded" in msg:
                            message = f"✅ อัปโหลดโฟลเดอร์ {sf} เสร็จแล้ว"
                            print(message)
                            log_messages_for_this_id.append(message)
                            upload_successful = True
                            break
                    except:
                        pass

                    time.sleep(2)
                    if (datetime.datetime.now() - start_time).seconds > 1800:
                        message = f"⚠️ อัปโหลดโฟลเดอร์ {sf} ใช้เวลาเกิน 30 นาที ข้าม"
                        print(message)
                        log_messages_for_this_id.append(message)
                        break
                if upload_successful:
                    time.sleep(3)
    else:
        message = f"⏩ Bucket {id_value} มีไฟล์อยู่แล้ว ({cell_text})"
        print(message)
        log_messages_for_this_id.append(message)

    # --- เพิ่มโค้ดส่วนที่ 3: บันทึก Log ของแถวนี้ลงไฟล์ ---
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"--- [END OF ROW] ID: {id_value} | Time: {timestamp} ---\n")
        for msg in log_messages_for_this_id:
            log_file.write(msg.strip() + "\n")
        log_file.write("-" * 70 + "\n\n")
    # ----------------------------------------------------

driver.quit()