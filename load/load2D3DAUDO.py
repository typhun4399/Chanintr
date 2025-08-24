import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, WebDriverException

excel_path = r"C:\Users\tanapat\Downloads\AUD New model id_21Aug25 (7).xlsx"
df = pd.read_excel(excel_path)

service = Service()

def create_chrome_driver_with_download_path(download_path):
    if not os.path.exists(download_path):
        os.makedirs(download_path, exist_ok=True)
        print(f"📂 สร้างโฟลเดอร์ดาวน์โหลด: {download_path}")

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("detach", False)
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def initial_open_site_and_click_buttons(driver):
    driver.get("https://audocph.com/")
    btn1_xpath = "/html/body/store-selector-popover/div/div/div[2]/a[2]"
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, btn1_xpath))).click()
    time.sleep(2)
    btn2_xpath = "/html/body/div[1]/div/div[4]/div[1]/div/div[2]/button[4]"
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, btn2_xpath))).click()
    time.sleep(2)

def open_search_modal_and_clear_input(driver):
    search_btn_xpath = "//summary[@aria-label='Search']"
    search_btn = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, search_btn_xpath)))
    aria_expanded = search_btn.get_attribute("aria-expanded")
    if aria_expanded != "true":
        driver.execute_script("arguments[0].click();", search_btn)
        time.sleep(1)
    search_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "Search-In-Modal")))
    search_input.clear()
    time.sleep(0.5)
    return search_input

def click_accordion_summary(driver):
    try:
        # รอ h2 ด้วยข้อความ "Additional resources"
        summary_h2 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Additional resources')]"))
        )

        # ไปหา element summary จาก h2
        summary_element = summary_h2.find_element(By.XPATH, "./ancestor::summary")

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", summary_element)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", summary_element)
        print("✅ กด Accordion summary (Additional resources) สำเร็จ")
        time.sleep(1)
        return True
    except TimeoutException:
        print("⚠️ ไม่พบ Accordion summary หรือกดไม่สำเร็จ")
        return False

def click_2d3d_button_and_switch_tab(driver):
    try:
        current_tabs = driver.window_handles

        # 🔍 หา <a> ที่ class มี "tab__additional-resources" และ text เป็น "2D/3D Files"
        file_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//a[contains(@class, 'tab__additional-resources') and normalize-space(text())='2D/3D Files']"
            ))
        )

        # คลิกปุ่ม
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", file_link)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", file_link)
        print("✅ กดลิงก์ '2D/3D FILES' สำเร็จ")

        # รอจนแท็บใหม่เปิด
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(current_tabs))
        new_tab = list(set(driver.window_handles) - set(current_tabs))[0]
        driver.switch_to.window(new_tab)
        print("🆕 สลับไปยังแท็บใหม่สำหรับดาวน์โหลด")
        time.sleep(2)
        return True

    except Exception as e:
        print(f"⚠️ ไม่เจอปุ่ม '2D/3D FILES' หรือสลับแท็บล้มเหลว: {e}")
        return False   

def click_add_to_selection_all_2d3d_files(driver):
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "widget__image")))
        file_blocks = driver.find_elements(By.CLASS_NAME, "widget__image")
        count = 0

        for block in file_blocks:
            try:
                title_el = block.find_element(By.CSS_SELECTOR, "p.file_title.color--light")
                file_name = title_el.text.strip()
                if file_name.lower().endswith("2d.zip") or file_name.lower().endswith("3d.zip"):
                    add_btn = block.find_element(By.XPATH, ".//span[contains(text(), 'ADD TO SELECTION')]/ancestor::a")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", add_btn)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", add_btn)
                    print(f"➕ กด ADD TO SELECTION สำหรับไฟล์: {file_name}")
                    count += 1
                    time.sleep(1)
            except Exception:
                continue

        if count == 0:
            print("⚠️ ไม่พบไฟล์ 2D.zip หรือ 3D.zip เลย")
            return False

        print(f"✅ กด ADD TO SELECTION ทั้งหมด {count} ไฟล์ 2D/3D.zip")
        return True
    except Exception as e:
        print(f"❌ พบปัญหาในการกด ADD TO SELECTION: {e}")
        return False

def click_basket_and_download_all(driver):
    try:
        basket_btn_xpath = "/html/body/div/app-root/div/div[2]/page/div/div/block/div/div/div/column/div/widget[2]/div/basketwidget/div/div/div/div"
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, basket_btn_xpath))).click()
        print("🧺 เปิดตะกร้า")
        time.sleep(2)

        download_all_btn_xpath = "/html/body/div/app-root/div/div[2]/div/div/div/basket/page/div/div/block/div/div/div/column/div/widget[1]/div/fileselectionresultoptionswidget/div/ul/li[2]"
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, download_all_btn_xpath))).click()
        print("⬇️ กดปุ่ม Download All")
        time.sleep(2)

        download_btn_xpath = "/html/body/div/app-root/div/multidownload-modal/div/div/div/div[2]/div/div/span/div[4]/div[1]/a"
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, download_btn_xpath))).click()
        print("💾 เริ่มดาวน์โหลดไฟล์ ZIP")
        return True
    except Exception as e:
        print(f"❌ ไม่สามารถดำเนินการดาวน์โหลดได้: {e}")
        return False

def wait_download_complete(driver):
    try:
        download_status_xpath = "/html/body/div/app-root/div/multidownload-modal/div/div/div/div[2]/div/h3"
        WebDriverWait(driver, 90).until(EC.text_to_be_present_in_element((By.XPATH, download_status_xpath), "Download complete"))
        print("✅ ดาวน์โหลดเสร็จสมบูรณ์ (ข้อความ 'Download complete' แสดงแล้ว)")
        time.sleep(3)
        return True
    except Exception as e:
        print(f"❌ รอข้อความดาวน์โหลดเสร็จล้มเหลว: {e}")
        return False


# ===== MAIN LOOP =====
try:
    for idx, row in df.iterrows():
        product_title = str(row['product_title']).strip()
        id_value = str(row['id']).strip()
        if not product_title or not id_value:
            print(f"⏭️ row {idx} ข้อมูลไม่ครบ (ชื่อหรือ id) ข้ามไป")
            continue

        download_folder = os.path.join(r"D:\AUDO\2D&3D", id_value)

        print(f"🔄 รอบ {idx} สำหรับสินค้า '{product_title}' id: {id_value}")
        print(f"📂 ตั้งค่าโฟลเดอร์ดาวน์โหลดเป็น: {download_folder}")

        try:
            driver = create_chrome_driver_with_download_path(download_folder)
            initial_open_site_and_click_buttons(driver)
        except WebDriverException as e:
            print(f"❌ ไม่สามารถเปิดเบราว์เซอร์หรือเข้าเว็บได้: {e}")
            break

        try:
            search_input = open_search_modal_and_clear_input(driver)
            search_input.send_keys(product_title)
            time.sleep(1)
            search_input.send_keys(Keys.ENTER)
            print(f"🔍 ค้นหา: {product_title}")

            first_card = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div.card.card--standard.card--media"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", first_card)
            time.sleep(0.5)
            first_card.click()
            print("✅ คลิกการ์ดสินค้าชิ้นแรก")

            if click_accordion_summary(driver):
                if click_2d3d_button_and_switch_tab(driver):
                    if not click_add_to_selection_all_2d3d_files(driver):
                        print("⏩ ข้ามสินค้าเนื่องจากไม่พบไฟล์ 2D.zip หรือ 3D.zip")
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        continue  # ข้ามไปสินค้าถัดไป

                    if click_basket_and_download_all(driver):
                        wait_download_complete(driver)

                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    print("🔙 กลับแท็บหลัก")

        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
            print(f"❌ row {idx} '{product_title}' ไม่เจอสินค้าหรือเกิดปัญหา: {e}")

        finally:
            print(f"🔚 ปิดเบราว์เซอร์รอบ {idx}")
            try:
                driver.quit()
            except:
                pass

except Exception as e:
    print("❌ Error ที่ไม่คาดคิด:", e)