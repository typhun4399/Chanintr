# -*- coding: utf-8 -*-
import os
import pandas as pd
import time
import datetime
import tkinter as tk
from tkinter import filedialog
import ttkbootstrap as tb
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap import ttk

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

# --- 1. ส่วน UI และการรับข้อมูล ---
def get_inputs_with_mode_selection():
    # --- ปุ่ม Browse ---
    def browse_excel():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        if path:
            excel_entry.delete(0, tk.END)
            excel_entry.insert(0, path)

    def browse_folder():
        path = filedialog.askdirectory()
        if path:
            folder_entry.delete(0, tk.END)
            folder_entry.insert(0, path)

    # --- เปิด/ปิดช่อง Folder ตามโหมด ---
    def toggle_folder_input():
        if mode_var.get() == "upload":
            folder_entry.config(state="normal")
            folder_browse_button.config(state="normal")
        else:
            folder_entry.config(state="disabled")
            folder_browse_button.config(state="disabled")

    # --- Load Columns ---
    def load_columns():
        path = excel_entry.get()
        if not path:
            Messagebox.show_error(title="Error", message="กรุณาเลือกไฟล์ Excel ก่อน")
            return
        try:
            df = pd.read_excel(path) if path.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(path)
            columns = df.columns.tolist()
            column_combobox['values'] = columns
            if 'id' in columns:
                column_var.set('id')
            else:
                column_var.set(columns[0])
            Messagebox.show_info(title="Success", message="โหลดคอลัมน์สำเร็จ")
        except Exception as e:
            Messagebox.show_error(title="Error", message=f"ไม่สามารถอ่านไฟล์ Excel: {e}")

    # --- Submit ---
    def submit():
        nonlocal selected_mode, email, password, excel_path, folder_path, selected_column
        
        selected_mode = mode_var.get()
        email = email_entry.get()
        password = password_entry.get()
        excel_path = excel_entry.get()
        folder_path = folder_entry.get() if selected_mode == "upload" else ""
        selected_column = column_var.get()

        # ตรวจสอบข้อมูล
        if not all([email, password, excel_path, selected_column]):
            Messagebox.show_error(title="ข้อมูลไม่ครบ", message="กรุณากรอก Email, Password, เลือกไฟล์ Excel และคอลัมน์")
            return
        if selected_mode == "upload" and not folder_path:
            Messagebox.show_error(title="ข้อมูลไม่ครบ", message="กรุณาเลือกโฟลเดอร์สำหรับโหมด Upload")
            return
        
        root.destroy()

    # --- ตัวแปร UI ---
    selected_mode = email = password = excel_path = folder_path = selected_column = ""

    root = tb.Window(themename="litera") 
    root.title("Cloud Storage Automation")
    root.resizable(False, False)

    # --- Frame สำหรับเลือกโหมด ---
    mode_frame = tb.Frame(root, padding=(10, 10))
    mode_frame.pack(fill="x")
    
    mode_var = tk.StringVar(value="upload")
    
    tb.Label(mode_frame, text="เลือกโหมดการทำงาน:").pack(side="left", padx=(0, 10))
    
    tb.Radiobutton(mode_frame, text="⬆️ Upload", variable=mode_var, value="upload", command=toggle_folder_input, bootstyle="primary-toolbutton").pack(side="left", fill="x", expand=True, padx=2)
    tb.Radiobutton(mode_frame, text="🗑️ Delete", variable=mode_var, value="delete", command=toggle_folder_input, bootstyle="primary-toolbutton").pack(side="left", fill="x", expand=True, padx=2)

    # --- Frame สำหรับกรอกข้อมูล ---
    main_frame = tb.Frame(root, padding=(10, 5))
    main_frame.pack(fill="x", expand=True)
    main_frame.columnconfigure(1, weight=1) 

    # Email
    tb.Label(main_frame, text="Email:").grid(row=0, column=0, sticky="e", padx=5, pady=8)
    email_entry = tb.Entry(main_frame)
    email_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=8)

    # Password
    tb.Label(main_frame, text="Password:").grid(row=1, column=0, sticky="e", padx=5, pady=8)
    password_entry = tb.Entry(main_frame, show="*")
    password_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=8)

    # Excel Path
    tb.Label(main_frame, text="Excel Path:").grid(row=2, column=0, sticky="e", padx=5, pady=8)
    excel_entry = tb.Entry(main_frame)
    excel_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=8)
    tb.Button(main_frame, text="Browse...", command=browse_excel, bootstyle="outline-secondary").grid(row=2, column=2, padx=5, pady=8)

    # Folder Path (สำหรับ Upload)
    tb.Label(main_frame, text="Folder Path:").grid(row=3, column=0, sticky="e", padx=5, pady=8)
    folder_entry = tb.Entry(main_frame)
    folder_entry.grid(row=3, column=1, sticky="ew", padx=5, pady=8)
    folder_browse_button = tb.Button(main_frame, text="Browse...", command=browse_folder, bootstyle="outline-secondary")
    folder_browse_button.grid(row=3, column=2, padx=5, pady=8)

    # Combobox เลือกคอลัมน์
    tb.Label(main_frame, text="เลือกคอลัมน์ ID:").grid(row=4, column=0, sticky="e", padx=5, pady=8)
    column_var = tk.StringVar()
    column_combobox = ttk.Combobox(main_frame, textvariable=column_var, state="readonly")
    column_combobox.grid(row=4, column=1, sticky="ew", padx=5, pady=8)
    tb.Button(main_frame, text="Load Columns", command=load_columns, bootstyle="outline-primary").grid(row=4, column=2, padx=5, pady=8)

    # ปุ่ม Submit
    submit_button = tb.Button(root, text="🚀 เริ่มทำงาน", command=submit, bootstyle="success", padding=(10, 10))
    submit_button.pack(pady=15, padx=10, fill="x")
    
    toggle_folder_input()
    root.mainloop()
    
    return selected_mode, email, password, excel_path, folder_path, selected_column

# --- 2. ฟังก์ชัน Selenium ---
def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def google_login(driver, wait, email, password):
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
        return True
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดระหว่างการล็อกอิน: {e}")
        return False

def write_log(id_value, messages):
    log_desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    log_file_path = os.path.join(log_desktop_path, 'Cloud Storage Automation.txt')
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"--- [ID: {id_value}] | Time: {timestamp} ---\n")
        for msg in messages:
            log_file.write(msg.strip() + "\n")
        log_file.write("-" * 70 + "\n\n")

# --- Upload Mode ---
def run_upload_mode(email, password, excel_path, base_folder, selected_column):
    print("\n--- ⬆️ เริ่มโหมด UPLOAD ---")
    driver = setup_driver()
    wait = WebDriverWait(driver, 20)

    if not google_login(driver, wait, email, password):
        driver.quit()
        return

    df = pd.read_excel(excel_path) if excel_path.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(excel_path)
    ids = df[selected_column].dropna().astype(str).tolist()
    
    base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects"

    for id_value in ids:
        log_messages = []
        target_folders = [f for f in os.listdir(base_folder) if f.startswith(id_value + '_') and os.path.isdir(os.path.join(base_folder, f))]
        
        if not target_folders:
            msg = f"❌ ไม่พบโฟลเดอร์สำหรับ id {id_value}"
            print(msg)
            log_messages.append(msg)
            write_log(id_value, log_messages)
            continue
            
        url = base_url.format(id_value)
        msg = f"\n🌐 กำลังตรวจสอบ URL: {url}"
        print(msg)
        log_messages.append(msg)
        driver.get(url)
        time.sleep(5)

        try:
            driver.find_element(By.XPATH, "//td[contains(text(),'No rows to display')]")
            is_empty = True
        except NoSuchElementException:
            is_empty = False

        if is_empty:
            msg = f"📂 Bucket {id_value} ว่าง เริ่มอัปโหลด"
            print(msg)
            log_messages.append(msg)
            for folder_name in target_folders:
                folder_path = os.path.join(base_folder, folder_name)
                subfolders = [sf for sf in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, sf))]
                if not subfolders:
                    msg = f"⚠️ ไม่มีโฟลเดอร์ย่อยใน {folder_path}"
                    print(msg)
                    log_messages.append(msg)
                    continue
                for sf in subfolders:
                    full_path = os.path.join(folder_path, sf)
                    msg = f"⬆️ กำลังอัปโหลด: {full_path}"
                    print(msg)
                    log_messages.append(msg)
                    try:
                        upload_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'][webkitdirectory]")))
                        upload_input.send_keys(full_path)
                        
                        started_xpath = "//div[contains(text(),'Upload started')]"
                        success_xpath = "//div[contains(text(),'successfully uploaded')]"

                        try:
                            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, started_xpath)))
                            WebDriverWait(driver, 1800).until(EC.invisibility_of_element_located((By.XPATH, started_xpath)))
                            time.sleep(2)
                            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, success_xpath)))
                            msg = f"✅ อัปโหลด {sf} เสร็จแล้ว"
                            print(msg)
                            log_messages.append(msg)
                        except TimeoutException:
                            msg = f"🟡 โฟลเดอร์ '{sf}' ว่าง หรือไม่พบการยืนยัน"
                            print(msg)
                            log_messages.append(msg)
                    except Exception as e:
                        msg = f"❌ เกิดข้อผิดพลาดระหว่างอัปโหลด '{sf}': {e}"
                        print(msg)
                        log_messages.append(msg)
        else:
            msg = f"⏩ Bucket {id_value} มีไฟล์อยู่แล้ว"
            print(msg)
            log_messages.append(msg)
            
        write_log(id_value, log_messages)

    print("\n🎉 โหมด Upload ทำงานเสร็จสิ้น")
    driver.quit()

# --- Delete Mode ---
def run_delete_mode(email, password, excel_path, selected_column):
    print("\n--- 🗑️ เริ่มโหมด DELETE ---")
    driver = setup_driver()
    wait = WebDriverWait(driver, 20)

    if not google_login(driver, wait, email, password):
        driver.quit()
        return

    df = pd.read_excel(excel_path) if excel_path.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(excel_path)
    ids = df[selected_column].dropna().astype(str).tolist()
    
    base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects"

    for id_value in ids:
        log_messages = []
        url = base_url.format(id_value)
        msg = f"\n🌐 กำลังตรวจสอบ URL: {url}"
        print(msg)
        log_messages.append(msg)
        driver.get(url)
        time.sleep(5)

        try:
            driver.find_element(By.XPATH, "//td[contains(text(),'No rows to display')]")
            is_empty = True
        except NoSuchElementException:
            is_empty = False

        if not is_empty:
            try:
                msg = f"🗑️ พบไฟล์ใน Bucket {id_value} กำลังดำเนินการลบ..."
                print(msg)
                log_messages.append(msg)

                select_all_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-pseudo-checkbox")))
                select_all_checkbox.click()

                delete_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Delete')]")))
                delete_button.click()
                time.sleep(1)
                confirm_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='text']")))
                confirm_input.send_keys("DELETE")
                time.sleep(1)

                actions = ActionChains(driver)
                actions.send_keys(Keys.ENTER).perform()
                print("🗑️ Confirm delete button")

                time.sleep(4)
                success_xpath = "//mat-snack-bar-container//div[contains(text(),'deleted')]"
                wait.until(EC.visibility_of_element_located((By.XPATH, success_xpath)))
                msg = f"✅ ลบไฟล์ทั้งหมดใน Bucket {id_value} สำเร็จ"
                print(msg)
                log_messages.append(msg)

            except Exception as e:
                msg = f"❌ เกิดข้อผิดพลาดระหว่างลบไฟล์ใน Bucket {id_value}: {e}"
                print(msg)
                log_messages.append(msg)
        else:
            msg = f"⏩ Bucket {id_value} ว่างอยู่แล้ว ไม่ต้องดำเนินการ"
            print(msg)
            log_messages.append(msg)
        
        write_log(id_value, log_messages)

    print("\n🎉 โหมด Delete ทำงานเสร็จสิ้น")
    driver.quit()

# --- 3. เริ่มต้นโปรแกรม ---
if __name__ == "__main__":
    mode, email, password, excel_path, folder_path, selected_column = get_inputs_with_mode_selection()

    if mode:
        if mode == 'upload':
            run_upload_mode(email, password, excel_path, folder_path, selected_column)
        elif mode == 'delete':
            run_delete_mode(email, password, excel_path, selected_column)
