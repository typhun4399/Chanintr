# -*- coding: utf-8 -*-
import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ---------------- CONFIG ----------------
excel_path = r"C:\Users\tanapat\Downloads\1_PTH_model id to get datasheet_17Oct25_matched_rows.xlsx"
base_download_path = r"D:\PTH\2D&3D"

# โหลด Excel
df = pd.read_excel(excel_path)
if "link" not in df.columns or "product_title" not in df.columns:
    raise Exception("❌ Excel ต้องมีคอลัมน์ 'link' และ 'product_title'")

# ---------------- Selenium Setup ----------------
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-notifications")
options.add_argument("--disable-extensions")
options.add_experimental_option("prefs", {
    "plugins.always_open_pdf_externally": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
})
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
driver.maximize_window()

# ---------------- ฟังก์ชันรอดาวน์โหลด ----------------
def wait_for_downloads(download_dir, timeout=600):
    """รอจนกว่าการดาวน์โหลดทั้งหมดจะเสร็จ"""
    seconds = 0
    while seconds < timeout:
        if any(fname.endswith((".crdownload", ".tmp")) for fname in os.listdir(download_dir)):
            time.sleep(1)
            seconds += 1
        else:
            break
    return not any(fname.endswith((".crdownload", ".tmp")) for fname in os.listdir(download_dir))

# ---------------- เปิดลิงก์ใน Excel ทีละลิงก์ ----------------
for idx, row in df.iterrows():
    link = str(row.get("link", "")).strip()
    product_title = str(row.get("product_title", "")).strip()

    if not link or link.lower() == "nan":
        print(f"⚪ แถวที่ {idx+1} ไม่มีลิงก์ — ข้าม")
        continue

    # สร้างโฟลเดอร์สำหรับ product_title
    folder_path = os.path.join(base_download_path, product_title)
    os.makedirs(folder_path, exist_ok=True)

    # ---------------- ตั้งค่า download path แบบ dynamic ----------------
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": folder_path
    })

    try:
        print(f"🔗 กำลังเปิดลิงก์ที่ {idx+1}: {link}")
        driver.get(link)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)

        # ---------------- ดึง <a> tag ----------------
        a_elements = driver.find_elements(By.CSS_SELECTOR,
            "#products > main > section:nth-child(3) > div > div.relative.w-100.w-50-s.bb.bn-s.br-s > div > div > a"
        )
        print(f"🔹 พบ <a> จำนวน {len(a_elements)} แถว")

        for i, a in enumerate(a_elements, start=1):
            href = a.get_attribute("href")
            text = a.text.strip()
            print(f"  {i}. {text} → {href}")

            if href and href.startswith("http"):
                driver.get(href)  # เปิดลิงก์ดาวน์โหลด
                time.sleep(2)
                wait_for_downloads(folder_path)
                print(f"✅ ดาวน์โหลดเสร็จสำหรับ: {text}")

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดที่แถว {idx+1}: {e}")

driver.quit()
print("✅ เปิดและดาวน์โหลดครบทุกลิงก์แล้ว")
