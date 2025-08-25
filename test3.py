import os
import time
import pandas as pd
import shutil
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ---------------- CONFIG ----------------
excel_input = r"C:\Users\tanapat\Downloads\WWS_model id to get 2D-3D_20Aug25_Test.xlsx"
excel_output = r"C:\Users\tanapat\Downloads\WWS_model id to get 2D-3D_20Aug25_Test.xlsx"
base_folder = r"D:\WWS\2D&3D"

df = pd.read_excel(excel_input)
search_list = df['style'].dropna().astype(str).tolist()
id_list = df['id'].dropna().astype(str).tolist()

prices = []

# ---------------- Main Loop ----------------
for idx, vid_search in enumerate(search_list):
    try:
        vid_folder = id_list[idx]
        id_folder = os.path.join(base_folder, vid_folder)

        # เตรียม subfolder
        datasheet_folder = os.path.join(id_folder, "Datasheet")
        folder_2d = os.path.join(id_folder, "2D")
        folder_3d = os.path.join(id_folder, "3D")
        os.makedirs(datasheet_folder, exist_ok=True)
        os.makedirs(folder_2d, exist_ok=True)
        os.makedirs(folder_3d, exist_ok=True)

        # --- Setup Chrome (ใหม่ทุกครั้ง) ---
        options = uc.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-extensions")

        prefs = {
            "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local"}],"selectedDestinationId":"Save as PDF","version":2}',
            "savefile.default_directory": base_folder
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--kiosk-printing")

        driver = uc.Chrome(options=options)
        wait = WebDriverWait(driver, 20)

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
        items = driver.find_elements(By.XPATH, "//ol/li[1]/div/div[1]")
        if items:
            try:
                items[0].click()
            except:
                print(f"⚠️ {vid_search}: element เจอแต่คลิกไม่ได้")

        time.sleep(2)

        # --- หา ul และ loop li ---
        ul_element = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[2]/main/div[3]/div/div/div[1]/div[2]/div/div/div[7]/div[1]/div[2]/div/ul")
        ))
        li_list = ul_element.find_elements(By.TAG_NAME, "li")

        for li in li_list:
            text_li = li.text.strip().lower()
            a_tag = li.find_element(By.TAG_NAME, "a")
            href = a_tag.get_attribute("href")

                # โหลดไฟล์แบบ TearSheet → ไป Datasheet
            try:
                driver.execute_script("window.open(arguments[0]);", href)
                driver.switch_to.window(driver.window_handles[-1])
                time.sleep(2)
                driver.execute_script("window.print();")
                time.sleep(5)

                downloaded_files = [f for f in os.listdir(base_folder) if f.endswith(".pdf")]
                if downloaded_files:
                    downloaded_file = downloaded_files[0]
                    source_path = os.path.join(base_folder, downloaded_file)
                    save_path = os.path.join(datasheet_folder, downloaded_file)
                    shutil.move(source_path, save_path)
                    print(f"✅ {vid_search}: บันทึก Datasheet '{downloaded_file}' เสร็จแล้ว")
                else:
                        print(f"⚠️ {vid_search}: ไม่พบไฟล์ PDF Datasheet ที่ดาวน์โหลด")
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            except Exception as e:
                print(f"⚠️ {vid_search}: โหลดไฟล์ '{text_li}' ผิดพลาด: {e}")

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
    finally:
        try:
            driver.quit()
        except:
            pass

# ---------------- Save Excel ----------------
df['Price'] = prices
df.to_excel(excel_output, index=False)
print(f"✅ บันทึกเสร็จ: {excel_output}")