# -*- coding: utf-8 -*-
import os
import time
import re
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from difflib import SequenceMatcher

# ---------------- CONFIG ----------------
excel_input = r"C:\Users\tanapat\Downloads\VCT SKU_to check New USD_9Oct25.xlsx"
log_file = "visualcomfort_log.txt"

# ---------------- โหลด Excel ----------------
df = pd.read_excel(excel_input)

# ✅ สร้างคอลัมน์ถ้ายังไม่มี
for col in ["New USD Price", "match", "%"]:
    if col not in df.columns:
        df[col] = ""

# ---------------- Setup Chrome ----------------
options = uc.ChromeOptions()
options.add_argument("--start-maximized")
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

# ---------------- Helper: คำนวณ % match ----------------
def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

# ---------------- Helper: แยกส่วนของ SKU ----------------
def split_sku_parts(sku):
    sku = sku.strip().replace(" ", "")
    match = re.match(r"([A-Za-z]+)(\d+)(.*)", sku)
    if match:
        prefix = match.group(1)
        num = match.group(2)
        detail = match.group(3)
        return prefix, num, detail
    return sku, "", ""

# ---------------- เริ่มทำงาน ----------------
driver.get("https://www.visualcomfort.com/")
time.sleep(5)

for idx, row in df.iterrows():
    raw_value = str(row["Vendor No. for Price Check"]).strip()
    if not raw_value or raw_value.lower() == "nan":
        continue

    model_no_raw = raw_value.replace(" ", "")
    prefix_ref, num_ref, _ = split_sku_parts(model_no_raw)
    if not num_ref:
        log_message(f"⚠️ Cannot parse SKU: {model_no_raw}")
        continue

    query = re.sub(r"\s+", "_", model_no_raw)
    url = f"https://www.visualcomfort.com/us/search?sort=relevance_desc&q={query}"
    driver.get(url)
    time.sleep(2)

    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//header//picture/img")
        ))
        log_message(f"🌐 Page loaded for {model_no_raw}")
    except:
        log_message(f"⚠️ Timeout waiting for page to load: {model_no_raw}")

    try:
        product_items = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                "#main div.products div.list > ol > li"
            ))
        )

        best_match = None
        best_ratio = 0.0
        best_li = None

        # 🔍 หา SKU ที่ตรงที่สุด
        for li in product_items:
            try:
                sku_p = li.find_element(By.CSS_SELECTOR, "div > div.sku > p")
                sku_text = sku_p.text.strip().replace(" ", "")
                prefix_site, num_site, _ = split_sku_parts(sku_text)

                if prefix_site == prefix_ref and num_site == num_ref:
                    ratio = 1.0
                else:
                    ratio = similarity(f"{prefix_ref}{num_ref}", f"{prefix_site}{num_site}")

                if ratio > best_ratio:
                    best_ratio = ratio
                    best_match = sku_text
                    best_li = li
            except:
                continue

        # ✅ เก็บเฉพาะ match 100% เท่านั้น
        if best_match and round(best_ratio * 100, 2) == 100.0:
            percent = 100.0
            log_message(f"✅ Perfect match for {model_no_raw}: {best_match}")

            try:
                info_div = best_li.find_element(By.CSS_SELECTOR, "div > div:nth-child(6)")
                print_text = info_div.text.strip()
                df.at[idx, "%"] = percent
                df.at[idx, "match"] = best_match
                df.at[idx, "New USD Price"] = print_text

                # ตรวจสอบว่ามีช่วงราคา (เช่น $1,199.00 - $1,799.00)
                if re.search(r"\$\d[\d,]*\.\d{2}\s*-\s*\$\d", print_text):
                    try:
                        link = best_li.find_element(By.CSS_SELECTOR, "div.name > a")
                        driver.execute_script("arguments[0].click();", link)
                        log_message(f"🖱 Clicked product link for {model_no_raw} (price range found)")
                        time.sleep(3)

                        variations = driver.find_elements(By.CSS_SELECTOR,
                            "#product_addtocart_form > div.product-item-variation-carousel-wrapper > div > div.owl-stage-outer > div > div > div > div > a"
                        )

                        for var in variations:
                            try:
                                driver.execute_script("arguments[0].click();", var)
                                time.sleep(2)

                                # ✅ ใช้ XPath ที่ยืดหยุ่น หา SKU ที่อยู่ในหน้าสินค้า
                                try:
                                    sku_element = WebDriverWait(driver, 5).until(
                                        EC.presence_of_element_located((By.XPATH,
                                            "//div[contains(@class, 'product-sku') or contains(@class, 'product-number') or @data-product-code]"
                                        ))
                                    )
                                    sku_text_page = sku_element.text.strip().replace(" ", "")
                                except:
                                    log_message(f"⚠️ SKU element not found after clicking variation for {model_no_raw}")
                                    continue

                                # ตรวจสอบว่า SKU ตรงกับที่ค้นหาไหม
                                if sku_text_page != model_no_raw:
                                    log_message(f"⏭ SKU mismatch ({sku_text_page} != {model_no_raw}), trying next variation...")
                                    continue

                                # ✅ หากตรงกัน — ดึงราคา
                                try:
                                    price_element = WebDriverWait(driver, 5).until(
                                        EC.presence_of_element_located((By.XPATH,
                                            "//span[contains(@class, 'price') or contains(@class, 'product-price')]//span[contains(text(),'$')]"
                                        ))
                                    )
                                    df.at[idx, "match"] = sku_text_page
                                    df.at[idx, "%"] = 100
                                    df.at[idx, "New USD Price"] = price_element.text.strip()
                                    log_message(f"💲 Price for {model_no_raw} found: {price_element.text.strip()}")
                                    break
                                except:
                                    log_message(f"⚠️ Price element not found for {model_no_raw}")
                                    continue

                            except Exception as e_var:
                                log_message(f"⚠️ Variation click error for {model_no_raw}: {e_var}")

                    except Exception as e:
                        log_message(f"⚠️ Could not click product link for {model_no_raw}: {e}")
                else:
                    df.at[idx, "match"] = best_match

            except Exception as e:
                df.at[idx, "match"] = best_match
                df.at[idx, "New USD Price"] = ""

        else:
            log_message(f"⚠️ No perfect match found for {model_no_raw}")
            df.at[idx, "match"] = ""
            df.at[idx, "%"] = ""
            df.at[idx, "New USD Price"] = ""

    except Exception as e:
        log_message(f"❌ Error on {model_no_raw}: {e}")
        continue

# ---------------- Save Excel ----------------
output_file = os.path.join(
    os.path.dirname(excel_input),
    f"VCT_with_match_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
)
df.to_excel(output_file, index=False)
log_message(f"✅ Done — Saved to {output_file}")

time.sleep(5)
driver.quit()
