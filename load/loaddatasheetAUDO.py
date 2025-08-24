import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor

def download_for_vendor(vendor, id_value):
    print(f"▶️ เริ่มดาวน์โหลด vendor: {vendor}, id: {id_value}")
    
    # ให้ไฟล์ไปอยู่ที่ id/Datasheet
    download_dir = os.path.join(r"D:\AUDO\2D&3D", id_value, "Datasheet")
    os.makedirs(download_dir, exist_ok=True)

    options = Options()
    # รันแบบ headless
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")  # กำหนดขนาดจอเสมือน
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(), options=options)

    try:
        driver.get("https://factsheet.audocph.com/")
        wait = WebDriverWait(driver, 15)

        input_box = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div/div/div[2]/div/div[2]/div[1]/div/input")))
        input_box.send_keys(vendor)

        dropdown_item = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div/div/div[2]/div/div[2]/div[1]/div/div/div")))
        dropdown_item.click()

        search_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div/div/div[2]/div/div[2]/button")))
        search_button.click()

        download_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//a[contains(@href, ".pdf") or contains(text(), "Download")]'))
        )
        download_button.click()
        print(f"✅ คลิกดาวน์โหลดแล้ว vendor: {vendor}")

        return driver, download_dir, vendor

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดกับ vendor {vendor}: {e}")
        driver.quit()
        return None, None, vendor

def main():
    excel_path = r"C:\Users\tanapat\Downloads\AUD New model id_21Aug25 (7).xlsx"
    df = pd.read_excel(excel_path)

    max_workers = 1
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for _, row in df.iterrows():
            vendor = str(row['vendor'])
            id_value = str(row['id'])
            futures.append(executor.submit(download_for_vendor, vendor, id_value))

        results = [f.result() for f in futures]

    def download_finished(folder_path, timeout=90):
        for _ in range(timeout):
            if not any(fname.endswith(".crdownload") for fname in os.listdir(folder_path)):
                return True
            time.sleep(1)
        return False

    for driver, folder, vendor in results:
        if driver is None:
            print(f"❌ ข้าม vendor {vendor} เนื่องจากเกิดข้อผิดพลาด")
            continue
        if download_finished(folder):
            print(f"✅ ดาวน์โหลดเสร็จสมบูรณ์สำหรับ vendor {vendor} ที่: {folder}")
        else:
            print(f"⚠️ ดาวน์โหลดไม่เสร็จสมบูรณ์สำหรับ vendor {vendor} ที่: {folder}")
        driver.quit()

if __name__ == "__main__":
    main()
