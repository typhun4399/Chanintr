import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ✅ กำหนดโฟลเดอร์ปลายทาง
download_dir = r"D:\AUDO"

# ✅ ตั้งค่า Chrome ให้ดาวน์โหลดไปยังโฟลเดอร์ที่กำหนด
options = Options()
options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=Service(), options=options)

# ✅ เปิดลิงก์
url = "https://audo.presscloud.com/digitalshowroom/#/gallery/Afteroom-Counter-Chair-Plus"
driver.get(url)

# ✅ รอโหลด block ของ 2D.zip (div[6]) และ 3D.zip (div[7])
block_xpaths = [
    "/html/body/div/app-root/div/div[2]/div/div/div/gallery/galleries/div/block/div/div/div/column/div/widget/div/imageswidget/div/div[2]/div[6]/imagecollectionview/imagegridview/div/div",
    "/html/body/div/app-root/div/div[2]/div/div/div/gallery/galleries/div/block/div/div/div/column/div/widget/div/imageswidget/div/div[2]/div[7]/imagecollectionview/imagegridview/div/div"
]

count = 0

for block_xpath in block_xpaths:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, block_xpath)))
    container = driver.find_element(By.XPATH, block_xpath)
    thumbnails = container.find_elements(By.TAG_NAME, "thumbnail")

    for thumb in thumbnails:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", thumb)
            time.sleep(0.5)

            title_els = thumb.find_elements(By.CLASS_NAME, "file_title")
            for title_el in title_els:
                text = title_el.text.strip()
                if text.endswith("2D.zip") or text.endswith("3D.zip"):
                    add_btn = thumb.find_element(By.XPATH, ".//span[contains(text(), 'ADD TO SELECTION')]/ancestor::a")
                    add_btn.click()
                    count += 1
                    print(f"✅ คลิกปุ่ม ADD TO SELECTION สำหรับ: {text}")
                    time.sleep(1)
        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาดใน thumbnail: {e}")

print(f"\n🎉 คลิกปุ่ม ADD TO SELECTION ไปทั้งหมด {count} ปุ่ม")

# ✅ คลิกตะกร้า
basket_btn_xpath = "/html/body/div/app-root/div/div[2]/page/div/div/block/div/div/div/column/div/widget[2]/div/basketwidget/div/div/div/div"
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, basket_btn_xpath))).click()
print("🧺 เปิดตะกร้า")

# ✅ คลิก "Download All"
download_all_btn_xpath = "/html/body/div/app-root/div/div[2]/div/div/div/basket/page/div/div/block/div/div/div/column/div/widget[1]/div/fileselectionresultoptionswidget/div/ul/li[2]"
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, download_all_btn_xpath))).click()
print("⬇️ กดปุ่ม Download All")

# ✅ คลิกปุ่มดาวน์โหลดสุดท้าย
download_btn_xpath = "/html/body/div/app-root/div/multidownload-modal/div/div/div/div[2]/div/div/span/div[4]/div[1]/a"
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, download_btn_xpath))).click()
print("💾 เริ่มดาวน์โหลด")

# ✅ รอจนกว่าจะมีข้อความ "Download complete" ใน <h3>
download_status_xpath = "/html/body/div/app-root/div/multidownload-modal/div/div/div/div[2]/div/h3"
WebDriverWait(driver, 90).until(
    EC.text_to_be_present_in_element((By.XPATH, download_status_xpath), "Download complete")
)
print("✅ มีข้อความ 'Download complete' แสดง")
time.sleep(5)
driver.quit()