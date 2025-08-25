import os
import time
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ---------------- CONFIG ----------------
excel_input = r"C:\Users\phunk\Downloads\WWS_model id to get 2D-3D_20Aug25.xlsx"
excel_output = r"C:\Users\phunk\Downloads\WWS_model id to get 2D-3D_20Aug25.xlsx"
base_folder = r"C:\Users\phunk\OneDrive\Desktop\WWS\2D&3D"

df = pd.read_excel(excel_input)
search_list = df['style'].dropna().astype(str).tolist()
id_list = df['id'].dropna().astype(str).tolist()

prices = []
links_all = []  # เก็บ links ทั้งหมดต่อ row

# --- Setup Chrome (ครั้งเดียว) ---
options = uc.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-notifications")
options.add_argument("--disable-extensions")

driver = uc.Chrome(options=options)
wait = WebDriverWait(driver, 20)

try:
    # ---------------- Main Loop ----------------
    for idx, vid_search in enumerate(search_list):
        try:
            vid_folder = id_list[idx]
            id_folder = os.path.join(base_folder, vid_folder)

            # ---------------- เริ่มทำงาน 1 row ----------------
            driver.get("https://www.waterworks.com/us_en/")
            time.sleep(3)

            # --- Search ---
            search_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]")))
            driver.execute_script("arguments[0].scrollIntoView(true);", search_box)
            search_box.click()
            search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//form/div[2]/input")))
            search_input.clear()
            search_input.send_keys(vid_search)
            search_input.send_keys(Keys.RETURN)
            time.sleep(2)

            # --- คลิก autocomplete ---
            try:
                first_item = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//ol/li[1]/div/div[1]"))
                )
                first_item.click()
            except:
                print(f"❌ {vid_search}: ไม่เจอ autocomplete, ข้ามไป")

            time.sleep(2)

            # --- หา container และ loop li ---
            li_list = []

            # เริ่มจาก div[7]
            try:
                container = driver.find_element(By.XPATH,
                    "/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[7]/div[1]/div[2]/div")
                ul_element = container.find_element(By.XPATH, ".//ul")
                li_list = ul_element.find_elements(By.TAG_NAME, "li")
            except:
                li_list = []

            # ถ้าไม่เจอ li ใน div[7] → ลอง div[8]-div[10]
            if not li_list:
                for i in range(8, 11):
                    try:
                        container = driver.find_element(By.XPATH,
                            f"/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[{i}]/div[1]/div[2]/div")
                        ul_element = container.find_element(By.XPATH, ".//ul")
                        li_list = ul_element.find_elements(By.TAG_NAME, "li")
                        if li_list:
                            break
                    except:
                        continue

            # --- เก็บ log ลงไฟล์ .txt ---
            datasheet_folder = os.path.join(id_folder, "Datasheet")
            log_path = os.path.join(datasheet_folder, "links.txt")
            links_data = []  # เก็บข้อมูลแต่ละ a สำหรับ Excel
            with open(log_path, "w", encoding="utf-8") as log_file:
                log_file.write(f"🔎 Links for {vid_search} ({vid_folder})\n")
                log_file.write("="*60 + "\n\n")

                if li_list:
                    for li in li_list:
                        li_text = li.text.strip()
                        try:
                            # ดึง a ทั้งหมดใน li
                            a_tags = li.find_elements(By.TAG_NAME, "a")
                            if not a_tags:  # ถ้าไม่มี a เลย
                                log_file.write(f"- li text: {li_text}\n  (no <a> found)\n\n")
                                links_data.append({
                                    "li_text": li_text,
                                    "a_text": "",
                                    "href": ""
                                })
                            else:
                                for a_tag in a_tags:
                                    href = a_tag.get_attribute("href")
                                    a_text = driver.execute_script("return arguments[0].innerText;", a_tag).strip()
                                    if not a_text:
                                        a_text = li_text  # fallback
                                    log_file.write(f"- li text: {li_text}\n  a text: {a_text}\n  URL: {href}\n\n")
                                    links_data.append({
                                        "li_text": li_text,
                                        "a_text": a_text,
                                        "href": href
                                    })
                        except:
                            log_file.write(f"- li text: {li_text}\n  ⚠️ Error reading <a>\n\n")
                            links_data.append({
                                "li_text": li_text,
                                "a_text": "",
                                "href": "(error)"
                            })
                else:
                    log_file.write("⚠️ ไม่พบ <ul><li> ในหน้านี้\n")

            print(f"✅ {vid_search}: บันทึก links ทั้งหมด {len(links_data)} รายการไปที่ {log_path}")
            links_all.append(str(links_data))  # แปลงเป็น str เพื่อใส่ Excel

            # --- ดึง Price ---
            try:
                price_el = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//div[4]/div[1]/div/span/span[1]/span"))
                )
                price_text = price_el.text.strip()
            except:
                price_text = ""
            prices.append(price_text)

        except Exception as e:
            print(f"❌ Error {vid_search}: {e}")
            prices.append("")
            links_all.append("")

finally:
    # ปิด Browser หลังทำงานทั้งหมดเสร็จ
    try:
        driver.quit()
    except:
        pass

# ---------------- Save Excel ----------------
df['Price'] = prices
df['Links'] = links_all
df.to_excel(excel_output, index=False)
print(f"✅ บันทึกเสร็จ: {excel_output}")
