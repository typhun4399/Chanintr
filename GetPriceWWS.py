import time
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- อ่าน Excel input ---
excel_path = r"C:\Users\tanapat\Downloads\WWS SKU to check new USD price_11Dec25.xlsx"
df = pd.read_excel(excel_path)
style_list = df["Style No."].dropna().astype(str).tolist()
print(f"Found {len(style_list)} Style No.")

# --- เปิด browser ---
driver = uc.Chrome(headless=False, version_main=142)
wait = WebDriverWait(driver, 20)

# --- เก็บผลลัพธ์ทั้งหมด ---
all_results = []

for style_no in style_list:
    url = f"https://www.waterworks.com/us_en/catalogsearch/result/?q={style_no}"
    print(f"\nOpening URL: {url}")
    driver.get(url)
    time.sleep(3)

    # --- ล้าง array สำหรับ Style No. ใหม่ ---
    results = []

    # --- ตรวจสอบ Style ---
    try:
        page_style = driver.find_element(
            By.CSS_SELECTOR, "#pdp-attribute-style"
        ).text.strip()
        if page_style != style_no:
            print(
                f"Style mismatch: page has '{page_style}', expected '{style_no}'. Skip."
            )
            continue
    except:
        print("No style found on page. Skip.")
        continue

    # --- หา dropdown options ---
    try:
        dropdown_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".dynamic-select"))
        )
        dropdown_button.click()
        time.sleep(1)

        li_list = wait.until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "li.product-option-li")
            )
        )
        total = len(li_list)
        print(f"Found {total} options for Style No. {style_no}")

        for i in range(total):
            # เปิด dropdown ใหม่ก่อนคลิก
            dropdown_button.click()
            time.sleep(1)

            # ดึง li ใหม่ทุกครั้ง
            li_list = wait.until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "li.product-option-li")
                )
            )
            li = li_list[i]
            option_text = li.text.strip()

            # คลิกด้วย JS เพื่อดึงข้อมูล
            driver.execute_script("arguments[0].scrollIntoView(true);", li)
            driver.execute_script("arguments[0].click();", li)
            time.sleep(1.5)

            # --- ดึงข้อมูล Style, SKU, Price ---
            try:
                style = driver.find_element(
                    By.CSS_SELECTOR, "#pdp-attribute-style"
                ).text.strip()
            except:
                style = style_no
            try:
                sku = driver.find_element(
                    By.CSS_SELECTOR, "#pdp-attribute-sku"
                ).text.strip()
            except:
                sku = "N/A"
            try:
                price = driver.find_element(
                    By.CSS_SELECTOR,
                    "#pdp-price-from-pricing-main > span > span.price-wrapper > span",
                ).text.strip()
            except:
                price = "N/A"

            # --- ตรวจสอบซ้ำภายใน array ของ Style No. นี้ ---
            is_duplicate = any(r["sku"] == sku and r["price"] == price for r in results)
            if is_duplicate:
                print(
                    f"[{i+1}] Skip '{option_text}' because SKU+Price already exists in this Style No."
                )
                continue

            # ถ้าไม่ซ้ำ → เก็บลง results
            results.append({"style": style, "sku": sku, "price": price})
            print(f"[{i+1}] Added: Style={style}, SKU={sku}, Price={price}")

    except Exception as e:
        print(f"No options found or error: {e}")
        continue

    # --- เก็บผลลัพธ์ของ Style No. นี้ ลง all_results ---
    all_results.extend(results)

# --- บันทึกลง Excel ---
output_path = r"C:\Users\tanapat\Desktop\waterworks_results.xlsx"
df_output = pd.DataFrame(all_results, columns=["style", "sku", "price"])
df_output.to_excel(output_path, index=False)
print(f"\n✅ Saved all results to {output_path}")

driver.quit()
