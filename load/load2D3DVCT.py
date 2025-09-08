import os
import time
import re
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# ---------------- CONFIG ----------------
excel_input = r"D:\VCT\AllNull.xlsx"
base_download = r"D:\VCT\2D&3D"
log_file = "visualcomfort_log.txt"

# ---------------- โหลด Excel ----------------
df = pd.read_excel(excel_input)

# ---------------- Setup Chrome ----------------
options = uc.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "plugins.always_open_pdf_externally": True  # ให้เปิด PDF เป็นไฟล์ดาวน์โหลด
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-notifications")
options.add_argument("--disable-extensions")
driver = uc.Chrome(version_main=139, options=options)
wait = WebDriverWait(driver, 20)

# ---------------- Helper log ----------------
def log_message(msg):
    print(msg)
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {msg}\n")

# ---------------- Helper set download dir ----------------
def set_download_dir(path):
    """เปลี่ยน download directory ของ Chrome ระหว่าง runtime"""
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": path
    })

# ---------------- Helper รอไฟล์ดาวน์โหลดเสร็จ ----------------
def wait_for_download(directory, timeout=60):
    """รอจนกว่าไฟล์ .crdownload จะหายไป"""
    end_time = time.time() + timeout
    while time.time() < end_time:
        downloading = [f for f in os.listdir(directory) if f.endswith(".crdownload")]
        if not downloading:
            return True
        time.sleep(1)
    return False

# ---------------- วนลูป Model No. ----------------
for idx, row in df.iterrows():
    model_no_raw = str(row["Model No."]).strip()
    id_value = str(row["id"]).strip() if "id" in df.columns else model_no_raw
    if not model_no_raw or model_no_raw.lower() == "nan":
        continue

    # ---------------- สำหรับ URL search ----------------
    query = re.sub(r"\s+", "_", model_no_raw)
    url = f"https://www.visualcomfort.com/us/search?sort=relevance_desc&q={query}"
    driver.get(url)
    time.sleep(2)

    # ---------------- สำหรับ match SKU ----------------
    model_no_clean = re.sub(r"\s+", "", model_no_raw)
    try:
        product_items = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                "#main > div.section.product-list-page-custom-container > div > div > div > div.products > div.list > ol > li"
            ))
        )

        clicked = False
        for li in product_items:
            try:
                sku_p = li.find_element(By.CSS_SELECTOR, "div > div.sku > p")
                sku_text = sku_p.text.strip().replace(" ", "")
                if sku_text == model_no_clean:
                    log_message(f"✅ Found matching SKU for {model_no_raw}: {sku_text}")
                    a_tag = li.find_element(By.CSS_SELECTOR, "div > div:first-child > a")
                    driver.execute_script("arguments[0].click();", a_tag)
                    clicked = True
                    time.sleep(5)
                    break
            except:
                continue

        if not clicked:
            log_message(f"⚠️ No matching SKU found for {model_no_raw}, skipping...")
            continue

    except Exception as e:
        log_message(f"❌ Error on {model_no_raw}: {e}")
        continue

    # ---------------- หลังจากเข้า product ----------------
    try:
        files_grid_css = (
            "#additional-specs > div.additional_block_inner > div.specs-flex > "
            "div.block.block_right > div.additional-data > div > div > div > div.block.files-grid"
        )
        grids = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, files_grid_css))
        )

        for grid in grids:
            links = grid.find_elements(By.TAG_NAME, "a")
            for link in links:
                link_text = link.text.strip()
                href = link.get_attribute("href")

                if not href or not link_text:
                    continue
                if link_text.lower() == "all":
                    continue

                # --------- เลือกโฟลเดอร์เก็บ ---------
                if href.lower().endswith(".pdf"):
                    subfolder = "Datasheet"
                elif link_text == "CAD Block":
                    subfolder = "2D"
                else:
                    subfolder = "3D"

                save_path = os.path.join(base_download, id_value, subfolder)
                os.makedirs(save_path, exist_ok=True)

                # ตั้งค่า download directory
                set_download_dir(save_path)

                log_message(f"⬇️ Downloading {link_text} ({href}) -> {save_path}")
                try:
                    driver.execute_script("arguments[0].click();", link)
                    wait_for_download(save_path, timeout=60)
                except Exception as e:
                    log_message(f"   ❌ Error clicking link {link_text}: {e}")
                time.sleep(2)
                
    except Exception as e:
        log_message(f"❌ Error processing product page for {model_no_raw}: {e}")

# ---------------- ปิด Browser ----------------
log_message("✅ Done")
time.sleep(30)
driver.quit()
