import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
import time
import os
import re
from datetime import datetime

# Import module for automatic ChromeDriver management
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

# ฟังก์ชันสำหรับทำความสะอาดชื่อไฟล์/โฟลเดอร์
def clean_filename(filename):
    """
    ทำความสะอาดสตริงเพื่อให้เหมาะสมกับการใช้งานเป็นชื่อไฟล์หรือชื่อโฟลเดอร์
    แทนที่อักขระที่ไม่ถูกต้องด้วยขีดล่างและลบช่องว่างนำหน้า/ต่อท้าย
    """
    cleaned_name = re.sub(r'[<>:"/\\|?*]', '_', filename)
    cleaned_name = cleaned_name.replace(' ', ' ') # แทนที่ non-breaking space ด้วย regular space
    cleaned_name = cleaned_name.strip()
    return cleaned_name

# ฟังก์ชันสำหรับพยายามกลับไปยังหน้าหลักและปิดโอเวอร์เลย์
def try_navigate_back_and_close_overlays(driver, base_url="https://www.carlhansen.com/en/en"):
    """
    พยายามปิดโอเวอร์เลย์/โมดัลและนำทางกลับไปยัง URL หลัก
    """
    try:
        pass # ตามคำขอ ไม่มีการจัดการปุ่มปิดหรือการกด ESCAPE อย่างชัดเจนที่นี่
        
    except Exception as e:
        print(f"    > เกิดข้อผิดพลาดขณะพยายามปิดโอเวอร์เลย์ (แต่ไม่ได้ดำเนินการปิด): {e}")
    finally:
        try:
            print(f"    > กำลังกลับไปยังหน้าหลัก: {base_url}")
            driver.get(base_url)
            # รอให้ปุ่มค้นหาบนหน้าแรกปรากฏขึ้นเพื่อยืนยันว่าโหลดเสร็จแล้ว
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Search']")) 
            )
            print("    > กลับไปยังหน้าหลักแล้ว")
            time.sleep(2) 
        except TimeoutException:
            print("    > ไม่สามารถกลับไปยังหน้าหลักได้ สถานะเบราว์เซอร์อาจไม่สอดคล้องกัน.")
        except Exception as e:
            print(f"    > เกิดข้อผิดพลาดที่ไม่คาดคิดขณะกลับหน้าหลัก: {e}")


# --- การทำงานหลัก ---
if __name__ == "__main__":
    # --- กำหนดพาธไฟล์ของคุณ ---
    # พาธไปยังไฟล์ Excel ที่มีชื่อโฟลเดอร์และ product_vendor
    excel_input_path = r"C:\Users\tanapat\Desktop\CHS\Empty_2D_3D_Folders_Categorized_Report_20250717_113214.xlsx"
    
    # พาธไปยังโฟลเดอร์หลักที่จะจัดเก็บโฟลเดอร์สินค้า
    base_product_folders_path = r"C:\Users\tanapat\Desktop\CHS\Test"

    # พาธสำหรับไฟล์ล็อก
    log_file_name = f"carlhansen_download_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    log_output_path = os.path.join(base_product_folders_path, log_file_name) # บันทึกไฟล์ล็อกในโฟลเดอร์ดาวน์โหลดหลัก
    download_log = [] # ลิสต์สำหรับเก็บรายการบันทึก

    # กำหนดค่า Chrome Options
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--log-level=3") # ระงับข้อความ INFO/WARNING จาก ChromeDriver
    chrome_options.add_argument("--ignore-certificate-errors") # จัดการปัญหา SSL ที่อาจเกิดขึ้น
    
    # กำหนดค่าเริ่มต้นชั่วคราวสำหรับไดเรกทอรีดาวน์โหลด โฟลเดอร์เฉพาะจะถูกกำหนดแบบไดนามิกสำหรับแต่ละผลิตภัณฑ์
    prefs = {
        "download.default_directory": base_product_folders_path, 
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safeBrowse.enabled": False 
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = None
    try:
        print("กำลังเริ่มต้น Chrome WebDriver...")
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.maximize_window()

        base_url = "https://www.carlhansen.com/en/en"
        driver.get(base_url)
        print(f"เปิดเบราว์เซอร์และไปยัง {base_url} แล้ว.")

        print("รอ 5 วินาทีเพื่อให้หน้าเว็บโหลดและจัดการ Cookie Banner...")
        time.sleep(5) 

        # ยอมรับคุกกี้ครั้งเดียวตอนเริ่มต้น
        try:
            accept_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.coi-banner__accept[onclick*='CookieInformation.submitAllCategories']"))
            )
            accept_button.click()
            print("คลิกปุ่ม 'Accept all' Cookie สำเร็จ.")
            time.sleep(2) 
        except TimeoutException:
            print("ไม่พบหรือไม่สามารถคลิกปุ่ม 'Accept all' Cookie (อาจไม่มี banner หรือถูกปิดไปแล้ว).")
        except ElementClickInterceptedException:
            print("การคลิกปุ่ม 'Accept all' Cookie ถูกขัดขวาง. ลองอีกครั้งด้วย JavaScript.")
            driver.execute_script("arguments[0].click();", accept_button)
            print("คลิกปุ่ม 'Accept all' Cookie ด้วย JavaScript สำเร็จ.")
            time.sleep(2)
        except Exception as e:
            print(f"เกิดข้อผิดพลาดขณะพยายามกดปุ่ม 'Accept all' Cookie: {e}")

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Search']"))
        )
        print("หน้าเว็บหลักพร้อมสำหรับการค้นหาแล้ว.")

        # 1. โหลดข้อมูลจาก Excel
        if not os.path.exists(excel_input_path):
            print(f"ข้อผิดพลาด: ไม่พบไฟล์ Excel '{excel_input_path}'.")
        else:
            try:
                df = pd.read_excel(excel_input_path)
                if 'Foldername' not in df.columns or 'product_vendor' not in df.columns:
                    print("ข้อผิดพลาด: ไฟล์ Excel ต้องมีคอลัมน์ 'Foldername' และ 'product_vendor'.")
                else:
                    print(f"โหลดข้อมูลสำเร็จ {len(df)} แถวจาก '{excel_input_path}'.")
                    
                    # วนลูปผ่านแต่ละแถวใน DataFrame เพื่อประมวลผลแต่ละรายการ
                    for index, row in df.iterrows():
                        folder_name_raw = str(row['Foldername']).strip()
                        current_product_vendor = str(row['product_vendor']).strip()
                        
                        if not current_product_vendor:
                            print(f"ข้ามแถวที่ {index+1}: 'product_vendor' ว่างเปล่าสำหรับ Foldername '{folder_name_raw}'.")
                            continue
                        
                        folder_name = clean_filename(folder_name_raw)

                        if not folder_name:
                            print(f"ข้ามแถวที่ {index+1}: Foldername ที่ทำความสะอาดแล้ว '{folder_name_raw}' ทำให้ได้สตริงว่างเปล่า. ไม่สามารถสร้างโฟลเดอร์ได้.")
                            continue

                        print(f"\n--- กำลังประมวลผล Product Vendor: '{current_product_vendor}' (สำหรับโฟลเดอร์: '{folder_name}') ---")

                        # กำหนดพาธโฟลเดอร์สินค้าเป้าหมายสำหรับการดาวน์โหลด
                        target_product_folder_dir = os.path.join(base_product_folders_path, folder_name)
                        if not os.path.exists(target_product_folder_dir):
                            os.makedirs(target_product_folder_dir)
                            print(f"สร้างไดเรกทอรี: {target_product_folder_dir}.")

                        # --- กำหนดไดเรกทอรีดาวน์โหลดสำหรับผลิตภัณฑ์ปัจจุบันแบบไดนามิก ---
                        print(f"    > ตั้งค่าโฟลเดอร์ดาวน์โหลดสำหรับสินค้าปัจจุบันไปที่: {target_product_folder_dir}")
                        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                            'behavior': 'allow',
                            'downloadPath': target_product_folder_dir
                        })
                        # --- สิ้นสุดการตั้งค่าไดเรกทอรีดาวน์โหลด ---

                        current_product_downloads = [] # สำหรับเก็บชื่อไฟล์ที่ดาวน์โหลดสำหรับผลิตภัณฑ์นี้
                        current_product_pdf_links = [] # สำหรับเก็บลิงก์ PDF สำหรับผลิตภัณฑ์นี้

                        try:
                            # คลิกปุ่มไอคอนค้นหาเพื่อแสดงช่องป้อนข้อมูลการค้นหา
                            print("    > กำลังคลิกปุ่มค้นหา...")
                            search_icon_button = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[aria-label='Search']"))
                            )
                            search_icon_button.click()
                            time.sleep(1)

                            # พิมพ์คำค้นหาลงในช่องป้อนข้อมูลการค้นหา
                            print(f"    > กำลังพิมพ์ '{current_product_vendor}' ลงในช่องค้นหา...")
                            search_input_field = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Type to start searching...'][class*='chakra-input']"))
                            )
                            search_input_field.clear() # ล้างข้อความที่มีอยู่ก่อนหน้า
                            search_input_field.send_keys(current_product_vendor)
                            search_input_field.send_keys(Keys.ENTER) # กด Enter เพื่อทำการค้นหา
                            print(f"    > พิมพ์ '{current_product_vendor}' และกด Enter แล้ว.")
                            
                            print("    > รอ 5 วินาทีเพื่อให้ผลการค้นหาโหลด...")
                            time.sleep(5) 

                            # --- ค้นหาและคลิกที่ลิงก์ผลิตภัณฑ์ที่ระบุ ---
                            print(f"    > กำลังค้นหาและคลิกที่ลิงก์สินค้า '{current_product_vendor}'...")
                            # XPath ที่อัปเดตเพื่อจับคู่ข้อความที่ *เริ่มต้น* ด้วย product_vendor ตามด้วย " |"
                            product_link_xpath = (
                                f"//a[contains(@class, 'chakra-link') and contains(@class, 'css-spn4bz') and "
                                f"starts-with(normalize-space(.), '{current_product_vendor} |')]"
                            )
                            
                            try:
                                # ใช้ presence_of_all_elements_located เพื่อหาลิงก์ที่ตรงกันทั้งหมด
                                product_links = WebDriverWait(driver, 20).until(
                                    EC.presence_of_all_elements_located((By.XPATH, product_link_xpath))
                                )

                                if product_links:
                                    if len(product_links) > 1:
                                        print(f"    > พบ {len(product_links)} ลิงก์ที่ตรงกันสำหรับ '{current_product_vendor}'. กำลังคลิกที่ลิงก์แรก.")
                                        log_entry = f"Folder '{folder_name}' ({current_product_vendor}): - พบ {len(product_links)} ลิงก์ที่ตรงกัน. คลิกที่ลิงก์แรก.\n"
                                        download_log.append(log_entry)
                                    
                                    product_link_to_click = product_links[0] # เลือกคลิกที่ลิงก์แรก
                                    driver.execute_script("arguments[0].scrollIntoView(true);", product_link_to_click) # เลื่อนไปยังองค์ประกอบ
                                    product_link_to_click.click()
                                    print(f"    > คลิกที่ลิงก์สินค้า '{current_product_vendor}' สำเร็จ.")
                                    time.sleep(5) # ให้เวลาโหลดหน้าสินค้าหลังจากคลิก
                                else:
                                    # กรณีที่ไม่พบลิงก์ใดๆ เลย (แม้ว่า WebDriverWait ควรจะ Timeout ก่อน)
                                    print(f"    > ไม่พบลิงก์สินค้าที่ตรงกับ '{current_product_vendor}'. กำลังข้ามไปยังรายการถัดไป.")
                                    log_entry = f"Folder '{folder_name}' ({current_product_vendor}): - ไม่พบลิงก์สินค้า\n"
                                    download_log.append(log_entry)
                                    try_navigate_back_and_close_overlays(driver, base_url)
                                    continue # ข้ามไปยังรายการถัดไปใน Excel
                            except TimeoutException:
                                print(f"    > ไม่พบลิงก์สินค้าที่ตรงกับ '{current_product_vendor}'. อาจไม่พบสินค้าหรือ selector ผิด. กำลังข้ามไปยังรายการถัดไป.")
                                log_entry = f"Folder '{folder_name}' ({current_product_vendor}): - ไม่พบลิงก์สินค้า\n"
                                download_log.append(log_entry)
                                try_navigate_back_and_close_overlays(driver, base_url)
                                continue # ข้ามไปยังรายการถัดไปใน Excel
                            except ElementClickInterceptedException:
                                print(f"    > การคลิกลิงก์สินค้า '{current_product_vendor}' ถูกขัดขวาง. กำลังพยายามคลิกด้วย JavaScript.")
                                driver.execute_script("arguments[0].click();", product_link_to_click) # ใช้ product_link_to_click ที่ถูกกำหนดแล้ว
                                print(f"    > คลิกที่ลิงก์สินค้า '{current_product_vendor}' ด้วย JavaScript สำเร็จ.")
                                time.sleep(5)
                            except Exception as e:
                                print(f"    > เกิดข้อผิดพลาดที่ไม่คาดคิดขณะคลิกลิงก์สินค้า '{current_product_vendor}': {e}. กำลังข้ามไปยังรายการถัดไป.")
                                log_entry = f"Folder '{folder_name}' ({current_product_vendor}): - ข้อผิดพลาดการคลิก: {e}\n"
                                download_log.append(log_entry)
                                try_navigate_back_and_close_overlays(driver, base_url)
                                continue # ข้ามไปยังรายการถัดไปใน Excel
                            # --- สิ้นสุดการคลิกลิงก์ผลิตภัณฑ์ ---


                            # --- เริ่มต้น: โค้ดที่เกี่ยวข้องกับการดาวน์โหลดจากสคริปต์ฉบับเต็มของคุณ ---

                            # คลิก "Dimensions & Downloads" accordion
                            print("    > กำลังพยายามคลิกปุ่ม 'Dimensions & Downloads'...")
                            accordion_button_xpath = "//button[contains(@class, 'chakra-accordion__button') and ./h2[contains(text(), 'Dimensions & Downloads')]]"
                            
                            try:
                                dimensions_downloads_button = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.XPATH, accordion_button_xpath))
                                )
                                driver.execute_script("arguments[0].scrollIntoView(true);", dimensions_downloads_button)
                                dimensions_downloads_button.click()
                                print("    > คลิกปุ่ม 'Dimensions & Downloads' สำเร็จ.")
                                time.sleep(3) # ให้เวลา accordion ขยายและโหลดเนื้อหา
                            except TimeoutException:
                                print("    > ไม่พบหรือไม่สามารถคลิกปุ่ม 'Dimensions & Downloads' ภายในเวลาที่กำหนด. ข้ามการดาวน์โหลดสำหรับสินค้านี้.")
                                log_entry = f"Folder '{folder_name}' ({current_product_vendor}): - ไม่พบส่วน 'Dimensions & Downloads'\n"
                                download_log.append(log_entry)
                                try_navigate_back_and_close_overlays(driver, base_url)
                                continue
                            except ElementClickInterceptedException:
                                print(f"    > การคลิกปุ่ม 'Dimensions & Downloads' ถูกขัดขวาง. ลองอีกครั้งด้วย JavaScript.")
                                driver.execute_script("arguments[0].click();", dimensions_downloads_button)
                                print("    > คลิกปุ่ม 'Dimensions & Downloads' ด้วย JavaScript สำเร็จ.")
                                time.sleep(3)
                            except Exception as e:
                                print(f"    > เกิดข้อผิดพลาดขณะพยายามคลิก 'Dimensions & Downloads': {e}. ข้ามการดาวน์โหลดสำหรับสินค้านี้.")
                                log_entry = f"Folder '{folder_name}' ({current_product_vendor}): - ข้อผิดพลาด Dimensions & Downloads: {e}\n"
                                download_log.append(log_entry)
                                try_navigate_back_and_close_overlays(driver, base_url)
                                continue

                            # ค้นหาและประมวลผลลิงก์ดาวน์โหลด ZIP และ PDF
                            print("    > กำลังมองหาลิงก์ดาวน์โหลด ZIP และ PDF เพื่อดำเนินการ...")
                            
                            zip_link_selector = "a.chakra-link.css-tiktpv[href$='.zip']"
                            pdf_link_selector = "a.chakra-link.css-tiktpv[href$='.pdf']"

                            # ประมวลผลไฟล์ ZIP
                            try:
                                zip_links = WebDriverWait(driver, 5).until(
                                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, zip_link_selector))
                                )
                                if zip_links:
                                    print(f"    > พบ {len(zip_links)} ลิงก์ ZIP. กำลังดำเนินการดาวน์โหลด...")
                                    for i, link in enumerate(zip_links):
                                        try:
                                            file_name = link.get_attribute('href').split('/')[-1]
                                            print(f"      - กำลังคลิกดาวน์โหลด ZIP: {file_name} (ลิงก์ที่ {i+1} จาก {len(zip_links)})")
                                            driver.execute_script("arguments[0].click();", link)
                                            current_product_downloads.append(file_name)
                                            time.sleep(1) # ให้เวลา 1 วินาทีสำหรับการเริ่มต้นดาวน์โหลดแต่ละครั้ง
                                        except ElementClickInterceptedException:
                                            print(f"      - การคลิกถูกขัดขวางสำหรับลิงก์ ZIP: {file_name}. พยายามใหม่ด้วย JS.")
                                            driver.execute_script("arguments[0].click();", link)
                                            current_product_downloads.append(file_name)
                                            time.sleep(1)
                                        except Exception as dl_error:
                                            print(f"      - ข้อผิดพลาดในการดาวน์โหลดลิงก์ ZIP: {file_name}: {dl_error}")
                                else:
                                    print("    > ไม่พบลิงก์ดาวน์โหลด ZIP สำหรับสินค้านี้.")
                            except TimeoutException:
                                print("    > ไม่พบลิงก์ดาวน์โหลด ZIP ในส่วน 'Dimensions & Downloads' ภายในเวลาที่กำหนด.")
                            except Exception as e:
                                print(f"    > เกิดข้อผิดพลาดขณะพยายามค้นหา/คลิกลิงก์ ZIP: {e}")

                            # ประมวลผลไฟล์ PDF (เก็บลิงก์ ไม่ดาวน์โหลด)
                            try:
                                pdf_links_elements = WebDriverWait(driver, 5).until(
                                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, pdf_link_selector))
                                )
                                if pdf_links_elements:
                                    print(f"    > พบ {len(pdf_links_elements)} ลิงก์ PDF. กำลังบันทึก URL...")
                                    for i, link_element in enumerate(pdf_links_elements):
                                        pdf_url = link_element.get_attribute('href')
                                        file_name = pdf_url.split('/')[-1]
                                        print(f"      - บันทึก URL PDF: {file_name} -> {pdf_url}")
                                        current_product_pdf_links.append(pdf_url)
                                else:
                                    print("    > ไม่พบลิงก์ PDF สำหรับสินค้านี้.")
                            except TimeoutException:
                                print("    > ไม่พบลิงก์ PDF ในส่วน 'Dimensions & Downloads' ภายในเวลาที่กำหนด.")
                            except Exception as e:
                                print(f"    > เกิดข้อผิดพลาดขณะพยายามค้นหา/บันทึกลิงก์ PDF: {e}")

                            # --- การตัดสินใจบันทึกตามกิจกรรม ---
                            if current_product_downloads or current_product_pdf_links:
                                log_entry = f"Folder '{folder_name}' ({current_product_vendor}):\n"
                                if current_product_downloads:
                                    log_entry += f"      - ดาวน์โหลด ZIP: {', '.join(current_product_downloads)}\n"
                                if current_product_pdf_links:
                                    log_entry += f"      - พบลิงก์ PDF (ไม่ได้ดาวน์โหลด): {', '.join(current_product_pdf_links)}\n"
                                download_log.append(log_entry)
                                print(f"    > บันทึกกิจกรรมการดาวน์โหลด/ลิงก์สำหรับ '{folder_name}'.")
                            else:
                                log_entry = f"Folder '{folder_name}' ({current_product_vendor}): - ไม่มีการดาวน์โหลดหรือพบลิงก์ใดๆ\n"
                                download_log.append(log_entry)
                                print(f"    > ไม่มีไฟล์ ZIP/PDF ที่ดาวน์โหลด/พบลิงก์ในโฟลเดอร์ '{folder_name}'.")
                            
                            time.sleep(1) # หน่วงเวลา 1 วินาทีหลังจากประมวลผลผลิตภัณฑ์นี้เสร็จสิ้น
                            # --- สิ้นสุดโค้ดดาวน์โหลด ---

                        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
                            print(f"เกิดข้อผิดพลาดสำหรับ '{current_product_vendor}': {type(e).__name__} - {e}. กำลังข้าม.")
                            download_log.append(f"Folder '{folder_name}' ({current_product_vendor}): - เกิดข้อผิดพลาด: {type(e).__name__} - {e}. ข้ามการประมวลผล\n")
                        except Exception as e:
                            print(f"เกิดข้อผิดพลาดที่ไม่สามารถจัดการได้สำหรับ '{current_product_vendor}': {e}. กำลังข้าม.")
                            download_log.append(f"Folder '{folder_name}' ({current_product_vendor}): - เกิดข้อผิดพลาดไม่คาดคิด: {e}. ข้ามการประมวลผล\n")
                        finally:
                            # กลับไปยังหน้าแรกสำหรับรายการถัดไปเสมอ
                            try_navigate_back_and_close_overlays(driver, base_url)

            except Exception as e:
                print(f"ข้อผิดพลาดในการประมวลผล Excel หรือวนลูป: {e}")

    except Exception as overall_error:
        print(f"เกิดข้อผิดพลาดโดยรวม: {overall_error}")
        download_log.append(f"เกิดข้อผิดพลาดโดยรวมในสคริปต์: {overall_error}\n")
    finally:
        # ปิดเบราว์เซอร์เมื่อเสร็จสิ้นหรือเกิดข้อผิดพลาด
        if driver:
            print("\nกำลังปิดเบราว์เซอร์...")
            driver.quit()
        
        # --- เขียนบันทึกลงในไฟล์ ---
        with open(log_output_path, 'w', encoding='utf-8') as f:
            f.write(f"--- Carl Hansen & Son Download Log ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---\n\n")
            for entry in download_log:
                f.write(entry)
        print(f"\nบันทึกกิจกรรมการดาวน์โหลดทั้งหมดลงในไฟล์: '{log_output_path}' แล้ว.")
        # --- สิ้นสุดการเขียนล็อก ---

    print("สคริปต์ทำงานเสร็จสิ้น.")
