import time
import pandas as pd
import os
import shutil  # <--- (มีอยู่แล้ว)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from difflib import SequenceMatcher

# --- ตั้งค่า ---
EXCEL_PATH = r"C:\Users\tanapat\Downloads\1_ARK model id to find 2D3D files_17Oct25.xlsx"
COLUMN_TO_MATCH = "name"
COLUMN_ID = "id"
URL = "https://www.ariakecollection.com/products/"
MATCH_THRESHOLD = 0.5
BASE_DOWNLOAD_PATH = r"D:\ARK\2D&3D"

# --- [ตั้งค่า Chrome Options] ---
os.makedirs(BASE_DOWNLOAD_PATH, exist_ok=True)
print(f"ไฟล์จะถูกดาวน์โหลด (ชั่วคราว) ไปที่: {BASE_DOWNLOAD_PATH}")

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-notifications")
options.add_argument("--disable-extensions")
prefs = {
    "download.default_directory": BASE_DOWNLOAD_PATH,
    "plugins.always_open_pdf_externally": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
}
options.add_experimental_option("prefs", prefs)

# --- ฟังก์ชันเทียบความคล้าย ---
def get_match_ratio(a, b):
    """คำนวณความคล้ายกันของ string สองชุด (0.0 ถึง 1.0)"""
    return SequenceMatcher(None, str(a).lower(), str(b).lower()).ratio()

# --- ฟังก์ชันรอการดาวน์โหลด ---
def wait_for_downloads(download_dir, timeout=3600):
    """รอจนกว่าการดาวน์โหลดทั้งหมดในโฟลเดอร์ที่ระบุจะเสร็จ"""
    print("    ...กำลังรอไฟล์ดาวน์โหลดเสร็จ...")
    seconds = 0
    while seconds < timeout:
        if any(fname.endswith((".crdownload", ".tmp")) for fname in os.listdir(download_dir)):
            time.sleep(1)
            seconds += 1
        else:
            return True  # ไม่มีไฟล์ .crdownload แล้ว
    print("    !! การดาวน์โหลด Timeout !!")
    return False

# --- Setup WebDriver ---
try:
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 10)
except Exception as e:
    print(f"ไม่สามารถ setup Chrome WebDriver ได้: {e}")
    exit()

# --- 1. อ่านข้อมูลจาก Excel ---
try:
    print(f"กำลังอ่านไฟล์ Excel จาก: {EXCEL_PATH}")
    df = pd.read_excel(EXCEL_PATH)
    
    if COLUMN_TO_MATCH not in df.columns or COLUMN_ID not in df.columns:
        print(f"!! ข้อผิดพลาด: ไม่พบคอลัมน์ '{COLUMN_TO_MATCH}' หรือ '{COLUMN_ID}'")
        print(f"  คอลัมน์ที่พบ: {df.columns.tolist()}")
        driver.quit()
        exit()
    
    print(f"พบข้อมูล {len(df)} แถวใน Excel")

except FileNotFoundError:
    print(f"!! ข้อผิดพลาด: ไม่พบไฟล์ Excel ที่: {EXCEL_PATH}")
    driver.quit()
    exit()
except Exception as e:
    print(f"!! ข้อผิดพลาดในการอ่าน Excel: {e}")
    driver.quit()
    exit()

# --- 2. เปิดเว็บและดึงข้อมูลสินค้าทั้งหมด ---
print(f"กำลังเปิด URL: {URL}")
driver.get(URL)
main_window_handle = driver.current_window_handle
web_products = []

try:
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "main")))
    articles = driver.find_elements(By.CSS_SELECTOR, "main > article")
    if not articles:
        print("!! ไม่พบ <article> ใดๆ ในหน้าเว็บ")
        driver.quit()
        exit()

    print(f"พบ <article> ทั้งหมด {len(articles)} รายการในหน้าเว็บ")
    
    for article in articles:
        try:
            title = article.find_element(By.CSS_SELECTOR, "div.item-title").text
            link_element = article.find_element(By.CSS_SELECTOR, "div.item-cover > a")
            web_products.append({"title": title, "link_element": link_element})
        except NoSuchElementException:
            print("  - พบ article ที่ไม่มี title/link ขอข้าม")
            pass
    
    print(f"ดึงข้อมูลสินค้าจากเว็บได้ {len(web_products)} รายการ")

except TimeoutException:
    print("!! หน้าเว็บโหลดไม่สำเร็จ หรือไม่พบ element 'main'")
    driver.quit()
    exit()
except Exception as e:
    print(f"!! เกิดข้อผิดพลาดขณะดึงข้อมูลหน้าเว็บ: {e}")
    driver.quit()
    exit()

print("-" * 30)
print("--- เริ่มกระบวนการจับคู่และคลิก ---")
print("-" * 30)

# --- 3. วนลูปรายชื่อจาก Excel (ทีละแถว) ---
for index, row in df.iterrows():
    
    excel_name = str(row.get(COLUMN_TO_MATCH, "")).strip()
    excel_id = str(row.get(COLUMN_ID, "")).strip()

    if not excel_name or not excel_id:
        print(f"ข้ามแถวที่ {index+2}: ข้อมูลไม่ครบ (name='{excel_name}', id='{excel_id}')")
        print("-" * 20)
        continue
        
    best_match = None
    best_score = 0.0
    
    # 3.1 ค้นหาคู่ที่ "ดีที่สุด"
    for product in web_products:
        web_title = product["title"]
        score = get_match_ratio(excel_name, web_title)
        
        if score > best_score:
            best_score = score
            best_match = product
            
    # 3.2 ตรวจสอบว่าคะแนนดีพอหรือไม่
    if best_match and best_score >= MATCH_THRESHOLD:
        matched_title = best_match['title']
        link_to_click = best_match['link_element']
        
        print(f"จับคู่: '{excel_name}' (Excel) / ID: '{excel_id}'")
        print(f"  -> ตรงกับ: '{matched_title}' (Web) (คะแนน: {best_score:.2f})")
        
        try:
            # เปิดแท็บใหม่
            ActionChains(driver)\
                .key_down(Keys.CONTROL)\
                .click(link_to_click)\
                .key_up(Keys.CONTROL)\
                .perform()

            print("  ...กำลังเปิดในแท็บใหม่...")
            time.sleep(1)
            
            all_windows = driver.window_handles
            new_window_handle = [window for window in all_windows if window != main_window_handle][-1]
            driver.switch_to.window(new_window_handle)
            
            print(f"  ...สลับไปที่ URL: {driver.current_url}")
            
            # --- [โค้ดส่วนที่แก้ไข: สร้างโฟลเดอร์และย้ายไฟล์] ---
            
            # 1. สร้าง Path สำหรับโฟลเดอร์ ID
            safe_folder_name = "".join(c for c in excel_id if c.isalnum() or c in (' ', '_', '-')).rstrip()
            product_folder_path = os.path.join(BASE_DOWNLOAD_PATH, safe_folder_name)
            os.makedirs(product_folder_path, exist_ok=True)
            
            # [เพิ่ม] 1.1 สร้างโฟลเดอร์ย่อย (Datasheet และ 3D)
            datasheet_folder_path = os.path.join(product_folder_path, "Datasheet")
            threed_folder_path = os.path.join(product_folder_path, "3D")
            os.makedirs(datasheet_folder_path, exist_ok=True)
            os.makedirs(threed_folder_path, exist_ok=True)
            
            print(f"    -> เตรียม Path: {product_folder_path}")

            specsheet_container_selector = "div.single-product-specsheet"
            link1_selector = f"{specsheet_container_selector} > a:nth-child(1)"
            link2_selector = f"{specsheet_container_selector} > a:nth-child(2)"
            
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, specsheet_container_selector)))
                time.sleep(1)

                # 2. เก็บรายชื่อไฟล์ "ก่อน" ดาวน์โหลด
                files_before = set(os.listdir(BASE_DOWNLOAD_PATH))
                
                # 3. พยายามคลิกดาวน์โหลดทั้งสองลิงก์
                link1_clicked = False
                link2_clicked = False

                try:
                    link1_element = driver.find_element(By.CSS_SELECTOR, link1_selector)
                    print(f"    -> กำลังคลิก Link 1 ({link1_element.text})...")
                    link1_element.click()
                    link1_clicked = True
                    time.sleep(0.5)
                except NoSuchElementException:
                    print("    -> ไม่พบลิงก์ 1 (a:nth-child(1))")
                
                try:
                    link2_element = driver.find_element(By.CSS_SELECTOR, link2_selector)
                    print(f"    -> กำลังคลิก Link 2 ({link2_element.text})...")
                    link2_element.click()
                    link2_clicked = True
                    time.sleep(0.5)
                except NoSuchElementException:
                    print("    -> ไม่พบลิงก์ 2 (a:nth-child(2))")
                
                # 4. ถ้ารมีการคลิก ให้รอโหลด
                if link1_clicked or link2_clicked:
                    wait_for_downloads(BASE_DOWNLOAD_PATH) # รอโหลดเสร็จในโฟลเดอร์หลัก
                    
                    # 5. หาสไฟล์ใหม่และย้าย
                    files_after = set(os.listdir(BASE_DOWNLOAD_PATH))
                    new_files = files_after - files_before
                    
                    if new_files:
                        print(f"    -> พบ {len(new_files)} ไฟล์ใหม่ กำลังย้ายและจัดประเภท...")
                        for f in new_files:
                            if f.endswith((".crdownload", ".tmp")):
                                continue # ข้ามไฟล์ที่ยังโหลดไม่เสร็จ

                            src_path = os.path.join(BASE_DOWNLOAD_PATH, f)
                            
                            # [แก้ไข] ตรรกะการเลือกปลายทาง
                            if f.lower().endswith('.pdf'):
                                dest_path = os.path.join(datasheet_folder_path, f)
                                folder_type = "Datasheet"
                            elif f.lower().endswith('.zip'):
                                dest_path = os.path.join(threed_folder_path, f)
                                folder_type = "3D"
                            else:
                                # ถ้าไม่ใช่ทั้ง pdf และ zip ให้เก็บไว้ที่โฟลเดอร์ ID หลัก
                                dest_path = os.path.join(product_folder_path, f) 
                                folder_type = "ID-Root"

                            try:
                                # ตรวจสอบเผื่อไฟล์ชื่อซ้ำ (ถ้ามีอยู่แล้ว ให้ลบของเก่าทิ้ง)
                                if os.path.exists(dest_path):
                                    os.remove(dest_path)
                                    
                                shutil.move(src_path, dest_path)
                                print(f"      -> ย้ายไฟล์ '{f}' ไปยังโฟลเดอร์ {folder_type} เรียบร้อย")
                            except Exception as move_e:
                                print(f"      !! ไม่สามารถย้ายไฟล์ '{f}': {move_e}")
                    else:
                        print("    -> ไม่พบไฟล์ใหม่หลังจากการดาวน์โหลด (อาจมีปัญหา)")
                else:
                    print("    -> ไม่มีลิงก์ให้ดาวน์โหลด")
            
            except TimeoutException:
                print("    -> ไม่พบส่วนดาวน์โหลด (div.single-product-specsheet) ในหน้านี้")
            except Exception as e:
                print(f"    -> เกิดข้อผิดพลาดขณะดาวน์โหลด/ย้ายไฟล์: {e}")
            
            time.sleep(1) 
            # --- [จบส่วนที่แก้ไข] ---

            
            # 3.4 ปิดแท็บใหม่ และสลับกลับไปแท็บหลัก
            driver.close()
            driver.switch_to.window(main_window_handle)
            print("  ...ปิดแท็บและกลับมาหน้าหลัก")
            
        except Exception as e:
            print(f"  !! เกิดข้อผิดพลาดตอนพยายามคลิกหรือสลับแท็บ: {e}")
            if driver.current_window_handle != main_window_handle:
                all_windows = driver.window_handles
                for window in all_windows:
                    if window != main_window_handle:
                        driver.switch_to.window(window)
                        driver.close()
                driver.switch_to.window(main_window_handle)


    else:
        print(f"ไม่พบคู่ที่ตรงกันสำหรับ: '{excel_name}' (คะแนนสูงสุด: {best_score:.2f})")

    print("-" * 20) # จบการทำงานของ 1 ชื่อ