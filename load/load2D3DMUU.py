from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
import time

# ===== โหลด Excel =====
df = pd.read_excel(r"C:\Users\tanapat\Desktop\MUU\2D&3D.xlsx")
df = df.dropna(subset=['product_vendor', 'id'])
vendor_id_pairs = df[['product_vendor', 'id']].astype(str).values.tolist()

# ===== ตั้งค่าเบราว์เซอร์ =====
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

# ===== ลูปทีละ Vendor =====
for vendor, foldername in vendor_id_pairs:
    print(f"\n🟡 เริ่มดาวน์โหลด: {vendor} -> {foldername}")

    driver.get("https://www.muuto.com/search/")
    time.sleep(7)


    # ===== ค้นหา vendor =====
    try:
        search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/section[2]/div/button")))
        driver.execute_script("arguments[0].click();", search_btn)
        time.sleep(1)

        search_input = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.search-input")))
        search_input.clear()
        search_input.send_keys(vendor)
        time.sleep(2)
    except Exception as e:
        print(f"❌ ค้นหา vendor ไม่ได้: {e}")
        continue

    # ===== คลิกผลลัพธ์แรก =====
    try:
        first_product = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/div/div/div[1]/ul/li/article/div/div[1]/div[1]/a[1]")))
        first_product.click()
        wait.until(EC.url_contains("product"))
        time.sleep(2)
    except Exception as e:
        print(f"❌ ไม่พบสินค้า: {e}")
        continue

    # ===== คลิก Accordion "2D, 3D & Revit files" =====
    try:
        accordion_btns = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//button[contains(@class,'accordion__trigger')]")))
        found = False
        for btn in accordion_btns:
            if '2D, 3D & Revit files' in btn.text:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", btn)
                found = True
                break
        if not found:
            print("⚠️ ข้าม: ไม่พบ Accordion ที่มีข้อความ '2D, 3D & Revit files'")
            continue
    except Exception as e:
        print(f"❌ ล้มเหลวตอนเปิด Accordion: {e}")
        continue

    # ===== คลิกลิงก์เปิดแท็บใหม่ =====
    try:
        old_tabs = driver.window_handles

        link_xpath = "//dt[contains(., '2D, 3D & Revit files')]/following-sibling::dd[1]//a[contains(@href, 'muuto.com/MediaLibrary')]"
        try:
            link = wait.until(EC.element_to_be_clickable((By.XPATH, link_xpath)))
        except:
            print("❌ ไม่พบลิงก์ดาวน์โหลด")
            continue

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", link)
        print("➡️ เปิดหน้าดาวน์โหลดใหม่")
        time.sleep(5)

        wait.until(EC.number_of_windows_to_be(len(old_tabs) + 1))
        new_tab = [tab for tab in driver.window_handles if tab not in old_tabs][0]
        driver.switch_to.window(new_tab)

        # ===== รอให้รายการไฟล์โหลด =====
        wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "element-item")))
        time.sleep(1)

        # ===== สร้างโฟลเดอร์ปลายทาง =====
        base_path = os.path.join(r"C:\Users\tanapat\Desktop\MUU\2D&3D", foldername)
        folder_2d = os.path.join(base_path, "2D")
        folder_3d = os.path.join(base_path, "3D")
        os.makedirs(folder_2d, exist_ok=True)
        os.makedirs(folder_3d, exist_ok=True)

        # ===== ดาวน์โหลดไฟล์ =====
        items = driver.find_elements(By.CLASS_NAME, "element-item")
        for item in items:
            try:
                name = item.find_element(By.CLASS_NAME, "media-name").text
                is_2d = "2D" in name.upper()
                is_3d = "3D" in name.upper()
                if not (is_2d or is_3d):
                    continue

                download_path = folder_2d if is_2d else folder_3d
                driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                    "behavior": "allow",
                    "downloadPath": download_path
                })

                view_btn = item.find_element(By.XPATH, ".//button[contains(text(), 'VIEW MORE')]")
                view_btn.click()

                dl_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'action=save') and text()='Download']")))
                dl_btn.click()
                print(f"⬇️ ดาวน์โหลด: {name}")

                time.sleep(6)  # ปรับลดเพื่อเร่งความเร็ว

                close_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.close[data-dismiss='modal']")))
                close_btn.click()
                time.sleep(0.5)

            except Exception as e:
                print(f"⚠️ ข้ามไฟล์หนึ่ง: {e}")
                try:
                    close_btn = driver.find_element(By.CSS_SELECTOR, "button.close[data-dismiss='modal']")
                    if close_btn.is_displayed():
                        close_btn.click()
                except:
                    pass
                continue

        # ✅ ปิดแท็บดาวน์โหลด
        driver.close()
        driver.switch_to.window(old_tabs[0])
        print(f"✅ เสร็จสิ้นสำหรับ {vendor}")

    except Exception as e:
        print(f"❌ ล้มเหลวสำหรับ {vendor}: {e}")
        try:
            driver.close()
            driver.switch_to.window(old_tabs[0])
        except:
            pass
        continue

# ===== ปิด Browser =====
print("\n🎉 งานเสร็จสมบูรณ์")
driver.quit()
