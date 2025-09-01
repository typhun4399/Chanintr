import time
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# ---------------- CONFIG ----------------
excel_input = r"C:\Users\tanapat\Downloads\1_WWS_model id to get 2D-3D_20Aug25_updated style no.xlsx"
log_file = r"C:\Users\tanapat\Downloads\waterworks_price_log.txt"
batch_save = 10  # บันทึก Excel ทุก 10 รายการ

# ฟังก์ชันเขียน log ทั้ง console และไฟล์
def log(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(line + "\n")

# อ่านข้อมูลจากไฟล์ Excel
df = pd.read_excel(excel_input)
search_list = df['Style No.'].dropna().astype(str).tolist()

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
wait = WebDriverWait(driver, 10)

def search_and_open(vid_search):
    """ ฟังก์ชันค้นหาและคลิกรายการแรก """
    try:
        search_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", search_box)
        search_box.click()
        search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]/input")))
        search_input.clear()
        search_input.send_keys(vid_search)
        search_input.send_keys(Keys.RETURN)
    except:
        return False

    try:
        first_item = wait.until(EC.element_to_be_clickable((By.XPATH, "//ol/li[1]/div/div[1]")))
        first_item.click()
        time.sleep(2)
        return True
    except:
        return False

try:
    for idx, vid_search in enumerate(search_list):

        # ถ้ามีราคาอยู่แล้ว -> ข้ามไปเลย
        if pd.notna(df.loc[idx, "Price"]) and str(df.loc[idx, "Price"]).strip() != "":
            log(f"⏭️ {vid_search}: มีราคาที่บันทึกไว้แล้ว ข้ามไป")
            continue

        driver.get("https://www.waterworks.com/us_en/")
        time.sleep(2)
        
        # ค้นหาครั้งแรก
        search_and_open(vid_search)

        # --- ตรวจสอบ Style No. (รอบแรก) ---
        try:
            style_element = wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[3]/div[1]/div")
                )
            )
            style_text = style_element.text.strip()
        except:
            log(f"⚠️ {vid_search}: ไม่พบ element Style No.")
            df.loc[idx, "Price"] = "NO PRICE"
            continue

        if style_text != vid_search:
            log(f"⚠️ {vid_search}: ไม่ตรงกับบนเว็บ")
            df.loc[idx, "Price"] = "NO PRICE"

        # --- ดึงราคา ---
        price = ""
        try:
            price_element = driver.find_element(
                By.XPATH,
                    "/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[1]/div/span/span[1]"
                )
            price = price_element.text.strip()
            if price:
                log(f"💰 {vid_search}: ราคา = {price}")
            else:
                log(f"⚠️ {vid_search}: ไม่พบราคา")
                price = "NO PRICE"
        except:
            log(f"⚠️ {vid_search}: ไม่พบ element ราคา")
            price = "NO PRICE"

        df.loc[idx, "Price"] = price

        # --- บันทึก Excel เป็น batch ---
        if (idx + 1) % batch_save == 0:
            try:
                df.to_excel(excel_input, index=False)
                log(f"💾 บันทึก Excel batch (ถึงรายการที่ {idx + 1})")
            except Exception as e:
                log(f"❌ บันทึก Excel ล้มเหลว: {e}")

    # บันทึกตอนจบ
    df.to_excel(excel_input, index=False)
    log("✅ อัปเดตราคาลง Excel เรียบร้อย")

finally:
    driver.quit()
    log("🚪 ปิด Browser แล้ว")
