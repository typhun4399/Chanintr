import os
import pandas as pd
import time
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- 1. ส่วน UI และการรับข้อมูล ---

def get_inputs_with_mode_selection():
    
    # ฟังก์ชันสำหรับปุ่ม Browse
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

    # ฟังก์ชันเปิด/ปิดช่อง Folder Path ตามโหมดที่เลือก
    def toggle_folder_input():
        if mode_var.get() == "upload":
            folder_entry.config(state=tk.NORMAL)
            folder_browse_button.config(state=tk.NORMAL)
        elif mode_var.get() == "delete":
            folder_entry.config(state=tk.DISABLED)
            folder_browse_button.config(state=tk.DISABLED)

    # ฟังก์ชันเมื่อกดปุ่ม "เริ่มทำงาน"
    def submit():
        nonlocal selected_mode, email, password, excel_path, folder_path
        
        selected_mode = mode_var.get()
        email = email_entry.get()
        password = password_entry.get()
        excel_path = excel_entry.get()
        folder_path = folder_entry.get() if selected_mode == "upload" else ""

        # ตรวจสอบข้อมูล
        if not all([email, password, excel_path]):
            messagebox.showerror("ข้อมูลไม่ครบ", "กรุณากรอก Email, Password, และเลือกไฟล์ Excel")
            return
        if selected_mode == "upload" and not folder_path:
            messagebox.showerror("ข้อมูลไม่ครบ", "กรุณาเลือกโฟลเดอร์สำหรับโหมด Upload")
            return
            
        root.destroy()

    # ตั้งค่าตัวแปรสำหรับ UI
    selected_mode = email = password = excel_path = folder_path = ""

    # สร้างหน้าต่างหลัก
    root = tk.Tk()
    root.title("Cloud Storage Automation")
    root.resizable(False, False)

    # --- Frame สำหรับเลือกโหมด ---
    mode_frame = tk.Frame(root, padx=10, pady=10)
    mode_frame.pack(fill="x")
    
    mode_var = tk.StringVar(value="upload") # ค่าเริ่มต้นคือ upload
    
    tk.Label(mode_frame, text="เลือกโหมดการทำงาน:").pack(side="left", padx=5)
    tk.Radiobutton(mode_frame, text="⬆️ Upload", variable=mode_var, value="upload", command=toggle_folder_input).pack(side="left")
    tk.Radiobutton(mode_frame, text="🗑️ Delete", variable=mode_var, value="delete", command=toggle_folder_input).pack(side="left")

    # --- Frame สำหรับกรอกข้อมูล ---
    main_frame = tk.Frame(root, padx=10, pady=5)
    main_frame.pack()

    # Email
    tk.Label(main_frame, text="Email:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    email_entry = tk.Entry(main_frame, width=50)
    email_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5)

    # Password
    tk.Label(main_frame, text="Password:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    password_entry = tk.Entry(main_frame, width=50, show="*")
    password_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=5)

    # Excel Path
    tk.Label(main_frame, text="Excel Path:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    excel_entry = tk.Entry(main_frame, width=50)
    excel_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5)
    tk.Button(main_frame, text="Browse...", command=browse_excel).grid(row=2, column=3, padx=5, pady=5)

    # Folder Path (สำหรับ Upload)
    tk.Label(main_frame, text="Folder Path:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
    folder_entry = tk.Entry(main_frame, width=50)
    folder_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=5)
    folder_browse_button = tk.Button(main_frame, text="Browse...", command=browse_folder)
    folder_browse_button.grid(row=3, column=3, padx=5, pady=5)

    # ปุ่ม Submit
    tk.Button(root, text="🚀 เริ่มทำงาน", command=submit, font=("Segoe UI", 10, "bold"), bg="#4CAF50", fg="white").pack(pady=15)
    
    toggle_folder_input() # เรียกครั้งแรกเพื่อตั้งค่าสถานะของช่อง Folder
    root.mainloop()
    
    return selected_mode, email, password, excel_path, folder_path

# --- 2. ส่วนฟังก์ชันการทำงานหลัก (แยกตามโหมด) ---

def setup_driver():
    """ตั้งค่าและคืนค่า Selenium WebDriver"""
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def google_login(driver, wait, email, password):
    """ฟังก์ชันสำหรับล็อกอิน Google"""
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
    """ฟังก์ชันสำหรับบันทึก Log"""
    log_desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    log_file_path = os.path.join(log_desktop_path, 'Cloud Storage Automation.txt')
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"--- [ID: {id_value}] | Time: {timestamp} ---\n")
        for msg in messages:
            log_file.write(msg.strip() + "\n")
        log_file.write("-" * 70 + "\n\n")

# --- โหมด UPLOAD ---
def run_upload_mode(email, password, excel_path, base_folder):
    print("\n--- ⬆️ เริ่มโหมด UPLOAD ---")
    driver = setup_driver()
    wait = WebDriverWait(driver, 20)

    if not google_login(driver, wait, email, password):
        driver.quit()
        return

    df = pd.read_excel(excel_path) if excel_path.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(excel_path)
    ids = df['id'].dropna().astype(int).astype(str).tolist()
    
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

# --- โหมด DELETE ---
def run_delete_mode(email, password, excel_path):
    print("\n--- 🗑️ เริ่มโหมด DELETE ---")
    driver = setup_driver()
    wait = WebDriverWait(driver, 20)

    if not google_login(driver, wait, email, password):
        driver.quit()
        return

    df = pd.read_excel(excel_path) if excel_path.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(excel_path)
    ids = df['id'].dropna().astype(int).astype(str).tolist()
    
    base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects"

    for id_value in ids:
        log_messages = [] # << บรรทัดที่ 1: เตรียมลิสต์สำหรับเก็บ Log ของ ID นี้
        url = base_url.format(id_value)
        msg = f"\n🌐 กำลังตรวจสอบ URL: {url}"
        print(msg)
        log_messages.append(msg) # << บรรทัดที่ 2: เพิ่มข้อความแรกลงใน Log
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
                log_messages.append(msg) # << บรรทัดที่ 3: เพิ่มข้อความลงใน Log

                # 1. คลิก Checkbox เพื่อเลือกทั้งหมด
                select_all_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//th//mat-checkbox")))
                select_all_checkbox.click()
                time.sleep(1)

                # 2. คลิกปุ่ม Delete
                delete_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label, 'Delete')]")))
                delete_button.click()
                time.sleep(1)

                # 3. ยืนยันการลบในหน้าต่าง pop-up
                confirm_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@aria-label='Type DELETE to confirm']")))
                confirm_input.send_keys("DELETE")
                time.sleep(1)

                confirm_delete_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button/span[contains(text(), 'Delete')]/..")))
                confirm_delete_button.click()

                # รอจนกว่าจะกลับมาหน้าว่าง
                WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, "//td[contains(text(),'No rows to display')]")))
                
                msg = f"✅ ลบไฟล์ทั้งหมดใน Bucket {id_value} สำเร็จ"
                print(msg)
                log_messages.append(msg) # << บรรทัดที่ 4: เพิ่มข้อความลงใน Log

            except Exception as e:
                msg = f"❌ เกิดข้อผิดพลาดระหว่างลบไฟล์ใน Bucket {id_value}: {e}"
                print(msg)
                log_messages.append(msg) # << บรรทัดที่ 5: เพิ่มข้อความ Error ลงใน Log
        else:
            msg = f"⏩ Bucket {id_value} ว่างอยู่แล้ว ไม่ต้องดำเนินการ"
            print(msg)
            log_messages.append(msg) # << บรรทัดที่ 6: เพิ่มข้อความลงใน Log
        
        write_log(id_value, log_messages) # << บรรทัดที่ 7: บันทึก Log ทั้งหมดของ ID นี้ลงไฟล์

    print("\n🎉 โหมด Delete ทำงานเสร็จสิ้น")
    driver.quit()

# --- 3. ส่วนเริ่มต้นโปรแกรม ---
if __name__ == "__main__":
    mode, email, password, excel_path, folder_path = get_inputs_with_mode_selection()

    if mode: # ตรวจสอบว่าผู้ใช้กด "เริ่มทำงาน" ไม่ใช่ปิดหน้าต่างไปเฉยๆ
        if mode == 'upload':
            run_upload_mode(email, password, excel_path, folder_path)
        elif mode == 'delete':
            run_delete_mode(email, password, excel_path)