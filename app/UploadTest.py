import os
import pandas as pd
import datetime
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- ฟังก์ชัน GUI (ไม่เปลี่ยนแปลง) ---
def get_inputs():
    # ... (โค้ดส่วนนี้เหมือนกับต้นฉบับ) ...
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

# --- ฟังก์ชันช่วย log (ไม่เปลี่ยนแปลง) ---
def write_log(log_path, msg):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")
    print(msg)

# --- ฟังก์ชันหลัก ---
def main():
    email, password, excel_path, base_folder = get_inputs()

    # ตรวจสอบว่าผู้ใช้ป้อนข้อมูลครบหรือไม่
    if not all([email, password, excel_path, base_folder]):
        print("❌ การทำงานถูกยกเลิก ผู้ใช้ไม่ได้ป้อนข้อมูล")
        return

    log_file = os.path.join(base_folder, "upload_log.txt")

    try:
        if excel_path.lower().endswith(".csv"):
            df = pd.read_csv(excel_path)
        else:
            df = pd.read_excel(excel_path)
        ids = df['id'].dropna().astype(int).astype(str).tolist()
    except Exception as e:
        write_log(log_file, f"❌ ไม่สามารถอ่านไฟล์ Excel/CSV ได้: {e}")
        return

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    
    # IMPROVEMENT: เพิ่ม Timeout สำหรับการรอที่นานขึ้น สำหรับการอัปโหลดโดยเฉพาะ
    long_wait = WebDriverWait(driver, 1800) # 30 นาที สำหรับรออัปโหลด
    short_wait = WebDriverWait(driver, 30)  # 30 วินาที สำหรับรอ element ทั่วไปในหน้าเว็บ

    try:
        # --- ล็อกอิน Google ---
        write_log(log_file, "🔐 กำลังล็อกอินเข้าสู่ระบบ Google...")
        driver.get("https://accounts.google.com/signin")
        email_input = short_wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
        email_input.clear()
        email_input.send_keys(email)
        driver.find_element(By.ID, "identifierNext").click()

        password_input = short_wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
        password_input.clear()
        password_input.send_keys(password)
        driver.find_element(By.ID, "passwordNext").click()
        
        # IMPROVEMENT: รอ element ที่มีเฉพาะในหน้า GCS แทนการใช้ time.sleep()
        write_log(log_file, "...รอหน้า Google Cloud Console โหลด...")
        # หมายเหตุ: Selector นี้อาจเปลี่ยนแปลงได้ในอนาคต เป็น Selector ของส่วนเนื้อหาหลักในหน้า GCS
        short_wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "c-wiz[data-og-display-name='Cloud Storage']")))
        write_log(log_file, "✅ ล็อกอินสำเร็จ")

        # --- เริ่มลูปตาม id ---
        base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects"

        for id_value in ids:
            target_folders = [f for f in os.listdir(base_folder) if f.startswith(id_value) and os.path.isdir(os.path.join(base_folder, f))]
            if not target_folders:
                write_log(log_file, f"❌ ไม่พบโฟลเดอร์หลักสำหรับ id {id_value}")
                continue

            url = base_url.format(id_value)
            write_log(log_file, f"🌐 เปิด URL: {url}")
            driver.get(url)

            # IMPROVEMENT: รอให้ตารางพร้อมใช้งานแทนการใช้ time.sleep()
            try:
                # รออย่างใดอย่างหนึ่งระหว่าง "ข้อความไม่พบข้อมูล" หรือ "หัวตารางข้อมูล"
                WebDriverWait(driver, 15).until(
                    EC.any_of(
                        EC.presence_of_element_located((By.XPATH, "//td[contains(text(),'No rows to display')]")),
                        EC.presence_of_element_located((By.XPATH, "//div[@role='columnheader' and contains(., 'Name')]"))
                    )
                )
            except TimeoutException:
                 write_log(log_file, f"⚠️ ไม่สามารถโหลดเนื้อหาสำหรับ bucket {id_value} ได้ในเวลาที่กำหนด ข้าม...")
                 continue

            # IMPROVEMENT: ใช้ find_elements เพื่อตรวจสอบการมีอยู่ของ element โดยไม่ทำให้เกิด error
            no_rows_element = driver.find_elements(By.XPATH, "//td[contains(text(),'No rows to display')]")

            if no_rows_element:
                write_log(log_file, f"📂 Bucket {id_value} ว่าง เริ่มอัปโหลด")
                
                # ค้นหาปุ่ม upload แค่ครั้งเดียวหลังจากโหลดหน้าเสร็จ
                try:
                    upload_input = short_wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file'][webkitdirectory]"))
                    )
                except TimeoutException:
                    write_log(log_file, f"❌ ไม่พบปุ่มอัปโหลดสำหรับ bucket {id_value} ข้าม...")
                    continue

                for folder_name in target_folders:
                    folder_path = os.path.join(base_folder, folder_name)
                    subfolders = [sf for sf in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, sf))]
                    if not subfolders:
                        write_log(log_file, f"⚠️ ไม่มีโฟลเดอร์ย่อยใน {folder_path}")
                        continue

                    for sf in subfolders:
                        full_path = os.path.join(folder_path, sf)
                        write_log(log_file, f"⬆️ กำลังอัปโหลด: {full_path}")

                        try:
                            upload_input.send_keys(full_path)
                            
                            # IMPROVEMENT: ปรับปรุง Logic การติดตามผลการอัปโหลดให้ง่ายขึ้น
                            success_xpath = "//mat-snack-bar-container//span[contains(text(),'successfully uploaded')]"
                            long_wait.until(EC.presence_of_element_located((By.XPATH, success_xpath)))
                            write_log(log_file, f"✅ อัปโหลดโฟลเดอร์ {sf} เสร็จแล้ว")
                            
                            # ควรรอให้ข้อความแจ้งเตือนสำเร็จหายไปก่อนเริ่มอัปโหลดไฟล์ถัดไป
                            short_wait.until(EC.invisibility_of_element_located((By.XPATH, success_xpath)))

                        except TimeoutException:
                            write_log(log_file, f"❌ อัปโหลดโฟลเดอร์ {sf} ล้มเหลว (หมดเวลา)")
                        except Exception as e:
                            write_log(log_file, f"❌ เกิดข้อผิดพลาดที่ไม่คาดคิดระหว่างอัปโหลด {sf}: {e}")
            else:
                write_log(log_file, f"⏩ Bucket {id_value} มีไฟล์อยู่แล้ว ข้าม...")

    finally:
        driver.quit()
        write_log(log_file, "🏁 จบการทำงานทั้งหมด")

if __name__ == "__main__":
    main()