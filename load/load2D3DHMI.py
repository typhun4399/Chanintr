import os
import time
import requests
import pandas as pd
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# โหลด Excel
excel_path = r"C:\Users\tanapat\Downloads\hermanmiller_products_with_cat_checked.xlsx"
df = pd.read_excel(excel_path)

# Base path สำหรับเก็บไฟล์
base_path = r"D:\HMI\loaded"
if not os.path.exists(base_path):
    os.makedirs(base_path)

# ---------------- BROWSER ----------------
options = uc.ChromeOptions()
#options.add_argument("--headless=new")  # ซ่อน browser
driver = uc.Chrome(version_main=141, options=options)
driver.maximize_window()


try:
    for idx, row in df.iterrows():
        product_title = str(row.get("Category", f"Product_{idx+1}")).strip()
        link = str(row.get("LoadLink", "")).strip()

        if not link or link.lower() == "nan":
            print(f"❌ แถวที่ {idx+1} ไม่มีลิงก์")
            continue

        # ทำชื่อโฟลเดอร์ให้ปลอดภัย
        safe_product = "".join(c if c not in '\\/:*?"<>|' else "_" for c in product_title)
        product_folder = os.path.join(base_path, safe_product)

        if not os.path.exists(product_folder):
            os.makedirs(product_folder)
            print(f"\n📁 สร้างโฟลเดอร์ใหญ่: {product_folder}")

        # เปิดลิงก์ของ product นี้
        print(f"\n🔗 เปิดลิงก์: {link}")
        driver.get(link)
        time.sleep(2)
        
        try:
            first_tab = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div/div/div[1]/div/div[2]/div/div/div/div/ul/li[1]/a")
            driver.execute_script("arguments[0].click();", first_tab)
            time.sleep(2)
            print("✅ กด tab แรกเรียบร้อย")
        except Exception as e:
            print(f"⚠️ ไม่สามารถกด tab แรก: {e}")

        # หา resource items
        resources = driver.find_elements(By.CSS_SELECTOR, "div.pro-resource-item.col-lg-10.col-lg-offset-1")
        print(f"✅ พบทั้งหมด {len(resources)} resource items ใน {product_title}")

        for i, res in enumerate(resources, start=1):
            # ชื่อรุ่น
            try:
                title_a = res.find_element(By.CSS_SELECTOR, "div.pro-resource-title a")
                title_text = title_a.text.strip()
            except:
                title_text = f"รุ่นที่_{i}"
            print(f"\n--- รุ่นที่ {i}: {title_text} ---")

            # safe name สำหรับโฟลเดอร์รุ่น
            safe_model = "".join(c if c not in '\\/:*?"<>|' else "_" for c in title_text)
            model_folder = os.path.join(product_folder, safe_model)
            folder_2d = os.path.join(model_folder, "2D")
            folder_3d = os.path.join(model_folder, "3D")
            for path in [model_folder, folder_2d, folder_3d]:
                if not os.path.exists(path):
                    os.makedirs(path)

            # ลิงก์ดาวน์โหลด
            file_links = res.find_elements(By.CSS_SELECTOR, "div.pro-resource-files ul li a")
            if file_links:
                print("พบลิงก์ดาวน์โหลด:")
                for f in file_links:
                    file_text = f.text.strip()
                    href = f.get_attribute("href")
                    if not href:
                        continue

                    # absolute URL
                    full_url = urljoin("https://www.hermanmiller.com", href)

                    # เลือก folder
                    if "2D" in file_text:
                        download_folder = folder_2d
                    else:
                        download_folder = folder_3d

                    local_filename = os.path.join(download_folder, os.path.basename(full_url))
                    print(f"- กำลังดาวน์โหลด {file_text} → {local_filename}")
                    try:
                        r = requests.get(full_url, stream=True)
                        r.raise_for_status()
                        with open(local_filename, 'wb') as f_out:
                            for chunk in r.iter_content(chunk_size=8192):
                                f_out.write(chunk)
                    except Exception as e:
                        print(f"❌ ไม่สามารถดาวน์โหลด {file_text}: {e}")
                    time.sleep(2)
            else:
                print("⚠️ ไม่พบลิงก์ดาวน์โหลด")

except Exception as e:
    print(f"\nเกิดข้อผิดพลาดใหญ่: {e}")

finally:
    driver.quit()
    print("\nเสร็จสิ้น ✅")
