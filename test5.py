import os
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
id_list = df['id'].dropna().astype(str).tolist()            # รายการ 'id' สำหรับตั้งชื่อโฟลเดอร์

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
        max_search_retries = 2
        current_search_retry = 0
        search_successful = False

        while not search_successful and current_search_retry < max_search_retries:
            try:
                driver.get("https://www.waterworks.com/us_en/")
                time.sleep(3)

                # --- ค้นหาสินค้า ---
                search_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]")))
                driver.execute_script("arguments[0].scrollIntoView(true);", search_box)
                search_box.click()
                search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]/input")))
                search_input.clear()
                search_input.send_keys(vid_search)
                search_input.send_keys(Keys.RETURN)
                time.sleep(2)

                # --- คลิกรายการแนะนำ (autocomplete) ---
                try:
                    first_item = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//ol/li[1]/div/div[1]"))
                    )
                    first_item.click()
                except:
                    pass
                time.sleep(2)

                # --- ตรวจสอบว่าพบสินค้า ---
                li_list = []
                for i in range(7, 10):
                    try:
                        container = driver.find_element(By.XPATH,
                            f"/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[{i}]/div[1]/div[2]/div")
                        ul_element = container.find_element(By.XPATH, ".//ul")
                        li_list = ul_element.find_elements(By.TAG_NAME, "li")
                        if li_list:
                            print(f"ℹ️ {vid_search} (ลองครั้งที่ {current_search_retry+1}): พบรายการสินค้าใน div[{i}]")
                            break
                    except:
                        continue

                if li_list:
                    search_successful = True
                else:
                    print(f"❌หา{vid_search}ไม่เจอ")
                    break

            except Exception as e:
                current_search_retry += 1
                print(f"❌ Error {vid_search} (ลองใหม่ครั้งที่ {current_search_retry}): {e}")
                if current_search_retry == max_search_retries:
                    print(f"❌ {vid_search}: เกิดข้อผิดพลาดซ้ำๆ หลังจาก {max_search_retries} ครั้ง, ข้ามไปสินค้ารายการถัดไป")
                    break
                time.sleep(2)

        if not search_successful:
            continue

        # --- ดึงราคา ---
        price = ""
        try:
            price_element = driver.find_element(By.XPATH, "/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[4]/div[1]/div/span/span[1]")
            price = price_element.text.strip()
            print(f"💰 {vid_search}: ราคา = {price}")
        except:
            print(f"⚠️ {vid_search}: ไม่พบราคาสินค้า")
        df.loc[idx, "Price"] = price

    # --- บันทึก Excel ---
    df.to_excel(excel_input, index=False)
    print("✅ อัปเดตราคาลง Excel เรียบร้อย")

finally:
    try:
        driver.quit()
    except:
        pass
