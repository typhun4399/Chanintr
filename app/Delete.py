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
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


# ------------------- ฟังก์ชัน GUI -------------------
def get_inputs():
    def browse_excel():
        path = filedialog.askopenfilename(
            filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, path)

    def submit():
        nonlocal email, password, excel_path, folder_path
        email = email_entry.get()
        password = password_entry.get()
        excel_path = excel_entry.get()
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

    url = base_url.format(id_value)
    print(f"\n🌐 เปิด URL: {url}")
    driver.get(url)
    time.sleep(5)

    try:
        # ------------------- Delete existing files/folders -------------------
        try:
            select_all_checkbox = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//mat-pseudo-checkbox")
            ))
            select_all_checkbox.click()

            delete_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Delete')]")
            ))
            delete_button.click()
            time.sleep(1)
            confirm_input = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[@type='text']")
            ))
            confirm_input.send_keys("DELETE")
            time.sleep(1)

            actions = ActionChains(driver)
            actions.send_keys(Keys.ENTER).perform()
            print("🗑️ Confirm delete button")

            time.sleep(4)
            
            # รอข้อความ success
            success_xpath = "//mat-snack-bar-container//div[contains(text(),'deleted')]"
            wait.until(EC.visibility_of_element_located((By.XPATH, success_xpath)))
            print(f"🗑️ Bucket {id_value} ลบไฟล์/โฟลเดอร์เก่าเรียบร้อย")
        except:
            print(f"⚠️ Bucket {id_value} ไม่มีไฟล์/โฟลเดอร์ให้ลบ")

    except Exception as e:
        print(f"⏩ Bucket {id_value} มีปัญหา: {e}")

driver.quit()
