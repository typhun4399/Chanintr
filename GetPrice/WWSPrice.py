import time
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ---------------- CONFIG ----------------
excel_input = r"C:\Users\tanapat\Downloads\WWS_model id to get 2D-3D_20Aug25_Test.xlsx"

# อ่านข้อมูลจากไฟล์ Excel
df = pd.read_excel(excel_input)
search_list = df['style'].dropna().astype(str).tolist()  # รายการ 'style' สำหรับค้นหา

# เพิ่มคอลัมน์ Price ถ้ายังไม่มี
if "Price" not in df.columns:
    df["Price"] = ""

# --- ตั้งค่า Chrome Driver ---
options = uc.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-notifications")
options.add_argument("--disable-extensions")

driver = uc.Chrome(options=options)
wait = WebDriverWait(driver, 20)

try:
    for idx, vid_search in enumerate(search_list):
        driver.get("https://www.waterworks.com/us_en/")
        time.sleep(2)

        # --- ค้นหาสินค้าและใช้ autocomplete ---
        search_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]")))
        search_box.click()
        search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]/input")))
        search_input.clear()
        search_input.send_keys(vid_search)
        time.sleep(1)

        # คลิกรายการแนะนำแรก
        try:
            first_item = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//ol/li[1]/div/div[1]"))
            )
            first_item.click()
            time.sleep(2)
        except:
            pass
            time.sleep(3)

        # --- ตรวจสอบว่ากลับมาที่หน้า homepage หรือไม่ ---
        if driver.current_url.rstrip("/") == "https://www.waterworks.com/us_en":
            print(f"⚠️ {vid_search}: ไม่มีสินค้า ข้ามไปอันถัดไป")
            df.loc[idx, "Price"] = ""
            continue

        # --- ดึงราคา (วนจนกว่าจะได้ค่าไม่ว่าง) ---
        price = ""
        max_retries = 10  # จำนวนครั้งสูงสุดในการลอง
        retry_count = 0

        while (price == "") and (retry_count < max_retries):
            try:
                price_element = driver.find_element(
                    By.XPATH,
                    "/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[1]/div/span/span[1]"
                )
                price = price_element.text.strip()
                if price:
                    print(f"💰 {vid_search}: ราคา = {price}")
                    break
            except:
                pass

            retry_count += 1
            time.sleep(1)  # รอแล้วลองใหม่

        if price == "":
            print(f"⚠️ {vid_search}: ไม่พบราคาหลังจากลอง {max_retries} ครั้ง")

        df.loc[idx, "Price"] = price

    # --- บันทึก Excel ---
    df.to_excel(excel_input, index=False)
    print("✅ อัปเดตราคาลง Excel เรียบร้อย")

finally:
    driver.quit()
