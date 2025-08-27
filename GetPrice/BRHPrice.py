import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re


# โหลด Excel
excel_path = r"C:\Users\tanapat\Downloads\BRN SKU_to search for USD price_6Aug25.xlsx"
df = pd.read_excel(excel_path)
vendor_numbers = df['Vendor Number'].dropna().astype(str).tolist()

options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

# เข้าเว็บและ login
driver.get("https://bernhardt.com/")

# คลิก Sign In
sign_in_btn = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/header/div[1]/div/div[2]/div[2]/div/div/div/nav/ul[2]/li[3]/button"))
)
sign_in_btn.click()

# กรอก email
email_input = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/form/modal/div[2]/div[2]/modal-body/section/div/div[2]/div/login-fieldset/div[1]/input"))
)
email_input.clear()
email_input.send_keys(" import@chanintr.com")

# กรอก password
password_input = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/form/modal/div[2]/div[2]/modal-body/section/div/div[2]/div/login-fieldset/div[2]/input"))
)
password_input.clear()
password_input.send_keys("chanintr2019")

# กด Login
login_btn = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/form/modal/div[2]/div[2]/modal-body/section/div/div[2]/div/div[1]/button"))
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
        
# กดค้นหา
        search_icon = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/header/div[1]/div/div[2]/div[2]/div/div/div/nav/ul[2]/li[4]"))
        )
        search_icon.click()
        
        search_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, search_input_xpath))
        )
        search_input.clear()
        search_input.send_keys(vendor_num)
        print(f"พิมพ์ product_vendor_number: {vendor_num}")

        first_item = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, first_autocomplete_xpath))
        )
        first_item.click()
        print(f"✅ คลิกรายการ autocomplete อันแรกสำหรับ {vendor_num}")

        # รอโหลดหน้า หรือ element ราคา
        price_xpath = "/html/body/div[2]/main/section[2]/div/div/div/div/div/ui-view/shopping-container/div/ui-view/shopping-one-up/div[3]/div[1]/div[2]/shopping-one-up-heading/div[3]/div[1]"
        price_el = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, price_xpath))
        )
        price_text = price_el.text.strip()
        match = re.search(r"\$([\d,]+\.\d{2})", price_text)
        if match:
            clean_price = match.group(1)
        else:
            clean_price = ""

        print(f"💰 ราคาสินค้า (clean): {clean_price}")
        prices.append(clean_price)

        time.sleep(1)

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดกับ {vendor_num}: {e}")
        prices.append("")  # ถ้าผิด ให้ใส่ค่าว่างแทน

# เพิ่มคอลัมน์ Price ใน DataFrame
df['MSRP on Web in USD'] = prices

# เขียนไฟล์ Excel ใหม่ หรือจะเขียนทับไฟล์เดิม
output_path = r"C:\Users\tanapat\Desktop\testbernh_Price.xlsx"
df.to_excel(output_path, index=False)
print(f"✅ บันทึกไฟล์ Excel พร้อมราคาไว้ที่: {output_path}")

driver.quit()