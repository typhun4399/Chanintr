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
excel_input = r"C:\Users\phunk\Downloads\WWS_model id to get 2D-3D_20Aug25.xlsx"
excel_output = r"C:\Users\phunk\Downloads\WWS_model id to get 2D-3D_20Aug25_price.xlsx"
base_folder = r"C:\Users\phunk\OneDrive\Desktop\WWS\2D&3D"

# ---------------- Read Excel ----------------
df = pd.read_excel(excel_input)
search_list = df['style'].dropna().astype(str).tolist()
id_list = df['id'].dropna().astype(str).tolist()

# ---------------- Setup Chrome (undetected) ----------------
options = uc.ChromeOptions()
options.add_argument("--start-maximized")

prefs = {
    "download.default_directory": base_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

driver = uc.Chrome(options=options)
wait = WebDriverWait(driver, 20)

driver.get("https://www.waterworks.com/us_en/")
time.sleep(5)

prices = []

# ---------------- Main Loop ----------------
for idx, vid_search in enumerate(search_list):
    try:
        vid_folder = id_list[idx]
        id_folder = os.path.join(base_folder, vid_folder)

        # --- Search ---
        search_box = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//form/div[2]"))
        )
        search_box.click()
        search_input = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//form/div[2]/input"))
        )
        search_input.clear()
        search_input.send_keys(vid_search)
        search_input.send_keys(Keys.RETURN)

        # --- คลิก autocomplete ถ้ามี ---
        try:
            first_item = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//ol/li[1]/div/div[1]"))
            )
            first_item.click()
        except:
            print(f"❌ {vid_search}: ไม่เจอ autocomplete, ข้ามไป")

        # --- รอ Technical Documents ---
        try:
            tech_doc = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='maincontent']//span[contains(text(), 'Technical Documents')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", tech_doc)
            time.sleep(1)
            tech_doc.click()
        except:
            print(f"❌ {vid_search}: ไม่มี Technical Documents")
            prices.append("")
            continue

        # --- Download Datasheet ---
        try:
            ds = driver.find_element(By.XPATH, "//*[@id='maincontent']//a[contains(text(), 'TearSheet')]")
            ds_link = ds.get_attribute("href")
            driver.get(ds_link)
            time.sleep(3)
            latest = max([os.path.join(base_folder, f) for f in os.listdir(base_folder)], key=os.path.getctime)
            shutil.move(latest, os.path.join(id_folder, "Datasheet", os.path.basename(latest)))
            driver.back()
        except:
            print(f"⚠️ {vid_search}: ไม่มี Datasheet")

        # --- Download 2D ---
        try:
            twod_links = driver.find_elements(By.XPATH, "//ul/li[6]/a")
            for link in twod_links:
                href = link.get_attribute("href")
                driver.get(href)
                time.sleep(3)
                latest = max([os.path.join(base_folder, f) for f in os.listdir(base_folder)], key=os.path.getctime)
                shutil.move(latest, os.path.join(id_folder, "2D", os.path.basename(latest)))
                driver.back()
        except:
            print(f"⚠️ {vid_search}: ไม่มี 2D")

        # --- Download 3D ---
        try:
            threeD_parent = driver.find_element(By.XPATH, "//ul/li[7]")
            threeD_links = threeD_parent.find_elements(By.TAG_NAME, "a")
            for link in threeD_links:
                href = link.get_attribute("href")
                driver.get(href)
                time.sleep(3)
                latest = max([os.path.join(base_folder, f) for f in os.listdir(base_folder)], key=os.path.getctime)
                shutil.move(latest, os.path.join(id_folder, "3D", os.path.basename(latest)))
                driver.back()
        except:
            print(f"⚠️ {vid_search}: ไม่มี 3D")

        # --- Price ---
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
        continue

# ---------------- Save Excel ----------------
df['Price'] = prices
df.to_excel(excel_output, index=False)
print(f"✅ บันทึกเสร็จ: {excel_output}")

driver.quit()
