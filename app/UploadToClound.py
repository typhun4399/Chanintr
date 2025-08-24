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


# ------------------- เริ่มโปรแกรมหลัก -------------------
email, password, excel_path, base_folder = get_inputs()

# โหลด id จาก Excel หรือ CSV
if excel_path.lower().endswith(".csv"):
    df = pd.read_csv(excel_path)
else:
    df = pd.read_excel(excel_path)
ids = df['id'].dropna().astype(str).tolist()

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
    target_folders = [f for f in os.listdir(base_folder) if f.startswith(id_value) and os.path.isdir(os.path.join(base_folder, f))]
    if not target_folders:
        print(f"❌ ไม่พบโฟลเดอร์หลักสำหรับ id {id_value}")
        continue

    # เปิด GCS
    url = base_url.format(id_value)
    print(f"\n🌐 เปิด URL: {url}")
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
        print(f"📂 Bucket {id_value} ว่าง เริ่มอัปโหลด")

        for folder_name in target_folders:
            folder_path = os.path.join(base_folder, folder_name)
            subfolders = [sf for sf in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, sf))]
            if not subfolders:
                print(f"⚠️ ไม่มีโฟลเดอร์ย่อยใน {folder_path}")
                continue

            for sf in subfolders:
                full_path = os.path.join(folder_path, sf)
                print(f"⬆️ กำลังอัปโหลด: {full_path}")

                # เลือกปุ่ม input[type=file][webkitdirectory]
                upload_input = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'][webkitdirectory]"))
                )
                upload_input.send_keys(full_path)

                # รอข้อความ success สำหรับโฟลเดอร์นี้
                success_xpath = "//mat-snack-bar-container//div[contains(text(),'successfully uploaded')]"
                start_time = datetime.datetime.now()

                while True:
                    try:
                        msg = driver.find_element(By.XPATH, success_xpath).text.strip().lower()
                        if "successfully uploaded" in msg:
                            print(f"✅ อัปโหลดโฟลเดอร์ {sf} เสร็จแล้ว")
                            break
                    except:
                        pass

                    time.sleep(2)
                    if (datetime.datetime.now() - start_time).seconds > 1800:  # 30 นาที
                        print(f"⚠️ อัปโหลดโฟลเดอร์ {sf} ใช้เวลาเกิน 30 นาที ข้าม")
                        break

    else:
        print(f"⏩ Bucket {id_value} มีไฟล์อยู่แล้ว ({cell_text})")

driver.quit()
