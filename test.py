import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re


# โหลด Excel
excel_path = r"C:\Users\tanapat\Downloads\BRN SKU to review USD_PL Feb26_28Jan26.xlsx"
df = pd.read_excel(excel_path)
vendor_numbers = df["Vendor Number"].dropna().astype(str).tolist()

options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

# เข้าเว็บและ login
driver.get("https://bernhardt.com/")

# คลิก Sign In
sign_in_btn = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located(
        (
            By.XPATH,
            "/html/body/div[2]/header/div[1]/div/div[4]/div/ul/li[2]/button",
        )
    )
)
sign_in_btn.click()

# กรอก email
email_input = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (
            By.XPATH,
            "/html/body/div[1]/div/div/form/modal/div[2]/div[2]/modal-body/section/div/div[2]/div/login-fieldset/div[1]/input",
        )
    )
)
email_input.clear()
email_input.send_keys(" import@chanintr.com")

# กรอก password
password_input = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (
            By.XPATH,
            "/html/body/div[1]/div/div/form/modal/div[2]/div[2]/modal-body/section/div/div[2]/div/login-fieldset/div[2]/input",
        )
    )
)
password_input.clear()
password_input.send_keys("chanintr2019")

# กด Login
login_btn = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (
            By.XPATH,
            "/html/body/div[1]/div/div/form/modal/div[2]/div[2]/modal-body/section/div/div[2]/div/div[1]/button",
        )
    )
)
login_btn.click()
print("✅ Login สำเร็จ")
time.sleep(5)

# รอให้หน้าโหลด input ช่องค้นหา
search_input_xpath = "/html/body/div[2]/main/section[2]/div/div/div/div/div/ui-view/shopping-container/div/ui-view/shopping-multi-view/section/div/div/input"
first_autocomplete_xpath = "/html/body/div[2]/main/section[2]/div/div/div/div/div/ui-view/shopping-container/div/ui-view/shopping-multi-view/section[2]/div/shopping-multi-view-cards/div[1]/div/a"

# เตรียม list เก็บราคาสินค้า
prices = []

for idx, vendor_num in enumerate(vendor_numbers):
    try:
        # 1. ใช้ Direct URL ในการค้นหา (แทนการกดปุ่ม Search Icon)
        search_url = f"https://www.bernhardt.com/shop/{vendor_num}?position=-1"
        driver.get(search_url)
        print(f"[{idx+1}/{len(vendor_numbers)}] กำลังค้นหา: {vendor_num}")

        # 2. รอให้รายการสินค้า (Card) ปรากฏในหน้าผลลัพธ์
        # ใช้ XPATH ของ SKU เพื่อเช็คว่ามีของขึ้นมาไหม
        vendor_check_xpath = "/html/body/div[2]/main/section[2]/div/div/div/div/div/ui-view/shopping-container/div/ui-view/shopping-one-up/div[3]/div[1]/div[2]/div[3]/div[1]/span"

        web_vendor_el = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, vendor_check_xpath))
        )
        web_sku_text = web_vendor_el.text.strip()

        # 3. ตรวจสอบว่ารหัสสินค้าตัวแรกที่เจอ ตรงกับที่เราค้นหาหรือไม่
        if web_sku_text.lower() == str(vendor_num).lower():
            print(f"✅ รหัสตรงกัน ({web_sku_text})")

            # 4. รอโหลดหน้า One-up เพื่อดึงราคา
            price_xpath = "/html/body/div[2]/main/section[2]/div/div/div/div/div/ui-view/shopping-container/div/ui-view/shopping-one-up/div[3]/div[1]/div[2]/shopping-one-up-heading/div[3]/div[1]"
            price_el = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, price_xpath))
            )

            price_text = price_el.text.strip()
            # ใช้ Regex ดึงราคา เช่น $1,234.00 -> 1,234.00
            match = re.search(r"\$([\d,]+\.\d{2})", price_text)
            clean_price = match.group(1) if match else "No Price Found"

            print(f"💰 ราคาสินค้า: {clean_price}")
            prices.append(clean_price)
        else:
            print(f"⚠️ รหัสไม่ตรงกัน (พบ: {web_sku_text} / ต้องการ: {vendor_num})")
            prices.append("Mismatch")

    except Exception as e:
        # กรณีหา Element ไม่เจอ หรือ Timeout (ไม่มีสินค้านี้บนเว็บ)
        print(f"❌ ไม่พบสินค้าหรือเกิดข้อผิดพลาดกับ {vendor_num}")
        prices.append("Not Found/Error")

    time.sleep(1)

# เพิ่มคอลัมน์ Price ใน DataFrame
df["MSRP on Web in USD"] = prices

# เขียนไฟล์ Excel ใหม่ หรือจะเขียนทับไฟล์เดิม
output_path = r"C:\Users\tanapat\Desktop\testbernh_Price.xlsx"
df.to_excel(output_path, index=False)
print(f"✅ บันทึกไฟล์ Excel พร้อมราคาไว้ที่: {output_path}")

driver.quit()
