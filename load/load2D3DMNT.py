# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

# ------------------ CONFIG ------------------
download_path = r"D:\MNT\Seymour"
www_path = "https://www.minotti.com/en/seymour_2"
os.makedirs(download_path, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-notifications")
options.add_argument("--disable-extensions")
prefs = {
    "download.default_directory": download_path,
    "plugins.always_open_pdf_externally": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
actions = ActionChains(driver)

# ------------------ HELPER FUNCTIONS ------------------
def wait_for_downloads(download_dir, timeout=3600):
    """รอจนกว่าการดาวน์โหลดทั้งหมดจะเสร็จ"""
    seconds = 0
    while seconds < timeout:
        if any(fname.endswith((".crdownload", ".tmp")) for fname in os.listdir(download_dir)):
            time.sleep(1)
            seconds += 1
        else:
            break
    return not any(fname.endswith((".crdownload", ".tmp")) for fname in os.listdir(download_dir))

def click_all_3D_elements():
    """คลิกทุกไฟล์ 3D โดยจัดการ overlay ของ javascript:void(0)"""
    # หา element ใหม่ทุกครั้งก่อนคลิก
    file_links_3D = driver.find_elements(By.CSS_SELECTOR, "#modelli2d3d-overlay a.download")
    total = len(file_links_3D)
    i = 0

    while i < total:
        try:
            # หา element ใหม่ทุกครั้ง
            file_links_3D = driver.find_elements(By.CSS_SELECTOR, "#modelli2d3d-overlay a.download")
            el = file_links_3D[i]
            href = el.get_attribute('href') or el.text

            driver.execute_script("arguments[0].scrollIntoView(true);", el)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", el)
            time.sleep(1)

            if href == 'javascript:void(0)':
                
                print("🔔 เจอ javascript:void(0), รอ overlay โหลด")
                WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                    (By.XPATH, "//div/h5[contains(text(),'Choose the kind of file to download')]")
                ))
                # หา <a> ใน overlay
                a_elements = driver.find_elements(By.CSS_SELECTOR, "#download-overlay ul a")
                print(f"🔗 Found {len(a_elements)} <a> tags in overlay:")
                for a in a_elements:
                    driver.get(a.get_attribute('href'))
                    print(" -", a.get_attribute('href') or a.text)
                    wait_for_downloads(download_path)

                # ปิด overlay
                close_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#download-overlay .overlay-navigation a"))
                )
                close_btn.click()
                print("✅ Closed overlay")

            print(f"➡ Clicked: {href}")
            i += 1  # ขยับไป element ถัดไป

        except StaleElementReferenceException:
            print("⚠ StaleElementReferenceException: หา element ใหม่แล้ว retry")
            time.sleep(1)
            continue
        except Exception as e:
            print(f"⚠ Failed to click element: {e}")
            i += 1  # ข้าม element ถัดไป

# ------------------ MAIN ------------------
try:
    driver.get(www_path)
    time.sleep(2)

    # ------------------ LOGIN ------------------
    login_icn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Login')]"))
    )
    login_icn.click()

    email_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#Email"))
    )
    email_input.clear()
    email_input.send_keys("purchasing@chanintr.com")

    password_input = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#Password"))
    )
    password_input.clear()
    password_input.send_keys("MINOTTI2025")

    login_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Please log in')]"))
    )
    login_btn.click()
    print("✅ Login successful")
    time.sleep(3)

    # ------------------ CLICK MAIN DOWNLOAD ------------------
    download_btn = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//a[text()='Download']"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", download_btn)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", download_btn)
    print("✅ Clicked the 'Download' button")

    # ------------------ DOWNLOAD MAIN PDF ------------------
    target_selector = "div.cta.dropdown ul li a"
    download_link_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, target_selector))
    )
    main_pdf_url = download_link_element.get_attribute("href")
    print("✅ Main PDF URL:", main_pdf_url)
    driver.get(main_pdf_url)
    time.sleep(5)

    # ------------------ CLICK 2D/3D FILES ------------------
    click_2D3D = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//a[text()='2d/3d files']"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", click_2D3D)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", click_2D3D)
    print("✅ Clicked '2D/3D files'")
    time.sleep(3)

    # ------------------ CLICK ALL 3D FILES ------------------
    click_all_3D_elements()
    
    # ------------------ WAIT FOR DOWNLOADS ------------------
    print("⏳ Waiting for downloads to finish...")
    if wait_for_downloads(download_path):
        print("✅ All downloads completed")
    else:
        print("⚠ Timeout waiting for downloads")

except Exception as e:
    print("❌ An error occurred:", e)
finally:
    print("Closing the browser.")
    driver.quit()
