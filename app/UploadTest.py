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

# ------------------- ฟังก์ชันช่วย log -------------------
def write_log(log_path, msg):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")
    print(msg)

# ------------------- เริ่มโปรแกรมหลัก -------------------
email, password, excel_path, base_folder = get_inputs()

# path สำหรับ log
log_file = os.path.join(base_folder, "upload_log.txt")

# โหลด id จาก Excel/CSV
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

# ------------------- ล็อกอิน Google -------------------
driver.get("https://accounts.google.com/signin")
email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
email_input.clear()
email_input.send_keys(email)
driver.find_element(By.ID, "identifierNext").click()

password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
password_input.clear()
password_input.send_keys(password)
driver.find_element(By.ID, "passwordNext").click()

time.sleep(5)  # รอโหลดหน้า GCS

# ------------------- เริ่มลูปตาม id -------------------
base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects"

for id_value in ids:
    target_folders = [f for f in os.listdir(base_folder) if f.startswith(id_value) and os.path.isdir(os.path.join(base_folder, f))]
    if not target_folders:
        write_log(log_file, f"❌ ไม่พบโฟลเดอร์หลักสำหรับ id {id_value}")
        continue

    url = base_url.format(id_value)
    write_log(log_file, f"🌐 เปิด URL: {url}")
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
        write_log(log_file, f"📂 Bucket {id_value} ว่าง เริ่มอัปโหลด")

        for folder_name in target_folders:
            folder_path = os.path.join(base_folder, folder_name)
            subfolders = [sf for sf in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, sf))]
            if not subfolders:
                write_log(log_file, f"⚠️ ไม่มีโฟลเดอร์ย่อยใน {folder_path}")
                continue

            for sf in subfolders:
                full_path = os.path.join(folder_path, sf)
                write_log(log_file, f"⬆️ กำลังอัปโหลด: {full_path}")

                upload_input = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'][webkitdirectory]"))
                )

                retry = 0
                max_retry = 3
                while retry < max_retry:
                    upload_input.send_keys(full_path)
                    write_log(log_file, f"🔄 Upload attempt {retry+1} สำหรับ {sf}")

                    success_xpath = "//mat-snack-bar-container//div[contains(text(),'successfully uploaded')]"
                    started_xpath = "//mat-snack-bar-container//div[contains(text(),'Upload started')]"
                    start_time = datetime.datetime.now()

                    while True:
                        try:
                            success_elems = driver.find_elements(By.XPATH, success_xpath)
                            started_elems = driver.find_elements(By.XPATH, started_xpath)

                            if success_elems:
                                write_log(log_file, f"✅ อัปโหลดโฟลเดอร์ {sf} เสร็จแล้ว")
                                break

                            if not started_elems:
                                write_log(log_file, f"⚠️ Upload folder {sf} อาจล้มเหลว")
                                retry += 1
                                time.sleep(2)
                                break

                        except Exception as e:
                            write_log(log_file, f"❌ เกิดข้อผิดพลาดระหว่างรอ upload: {e}")

                        time.sleep(2)
                        if (datetime.datetime.now() - start_time).seconds > 1800:
                            write_log(log_file, f"⚠️ อัปโหลดโฟลเดอร์ {sf} ใช้เวลาเกิน 30 นาที ข้าม")
                            retry = max_retry
                            break

                    if driver.find_elements(By.XPATH, success_xpath):
                        break

                    if retry >= max_retry:
                        write_log(log_file, f"❌ อัปโหลดโฟลเดอร์ {sf} ล้มเหลวหลัง {max_retry} attempts")
                        break

    else:
        write_log(log_file, f"⏩ Bucket {id_value} มีไฟล์อยู่แล้ว ({cell_text})")

driver.quit()
write_log(log_file, "🏁 จบการทำงานทั้งหมด")
