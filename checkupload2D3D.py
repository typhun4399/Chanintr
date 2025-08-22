import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# --- ตั้งค่า ---
excel_path = r"C:\Users\tanapat\Downloads\1_KCS active model id_27Jun25 (2).xlsx"
base_folder = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\14_KCS_done uploaded"
base_url = "https://console.cloud.google.com/storage/browser/chanintr-2d3d/production/{};tab=objects?inv=1&invt=Ab5Vew&prefix=&forceOnObjectsSortingFiltering=false"
mismatch_txt_path = r"C:\Users\tanapat\Desktop\mismatch_ids_ETH.txt"

# โหลด id จาก Excel
df = pd.read_excel(excel_path)
ids = df['id'].dropna().astype(str).tolist()

# ตั้งค่า selenium
chrome_options = Options()
chrome_options.add_argument("--start-maximized") 
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# --- ล็อกอิน Google ---
driver.get("https://accounts.google.com/v3/signin/accountchooser?continue=https%3A%2F%2Fwww.google.com%2F&ec=futura_exp_og_so_72776762_e&hl=en&ifkv=AdBytiMGIAttUZ1i34PUHaH6pAPdQX_zyXswMLfvRU4wKbXzwS2dw7_xaUs92mDAWPzAUA_TWbDLoA&passive=true&flowName=GlifWebSignIn&flowEntry=ServiceLogin&dsh=S1812811471%3A1755056943534734")

email_input = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//input[@type='email']"))
)
email_input.clear()
email_input.send_keys("tanapat@chanintr.com")

next_btn = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='identifierNext']"))
)
next_btn.click()

password_input = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//input[@type='password']"))
)
password_input.clear()
password_input.send_keys("Qwerty12345$$")

password_next = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='passwordNext']"))
)
password_next.click()

time.sleep(5)  # รอโหลดหน้า GCS

# --- เริ่มลูปตาม id ---
with open(mismatch_txt_path, "w") as txt_file:
    for id_value in ids:
        # หาโฟลเดอร์หลัก
        target_folders = [f for f in os.listdir(base_folder) if f.startswith(id_value) and os.path.isdir(os.path.join(base_folder, f))]
        if not target_folders:
            print(f"ไม่พบโฟลเดอร์หลักสำหรับ id {id_value}")
            txt_file.write(f"{id_value}\n")
            continue

        folder_path = os.path.join(base_folder, target_folders[0])
        folder_count = len([f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))])
        print(f"ID {id_value} มี {folder_count} subfolder ใน folder local")

        # เปิดเว็บ GCS ตาม id
        url = base_url.format(id_value)
        driver.get(url)
        time.sleep(5)  # รอโหลดหน้า

        # นับ row ใน table
        try:
            table_body_xpath = "/html/body/div[1]/div[3]/div[3]/div/pan-shell/pcc-shell/cfc-panel-container/div/div/cfc-panel/div/div/div[3]/cfc-panel-container/div/div/cfc-panel/div/div/cfc-panel-container/div/div/cfc-panel/div/div/cfc-panel-container/div/div/cfc-panel[2]/div/div/central-page-area/div/div/pcc-content-viewport/div/div/pangolin-home-wrapper/pangolin-home/cfc-router-outlet/div/ng-component/cfc-single-panel-layout/cfc-panel-container/div/div/cfc-panel/div/div/cfc-panel-body/cfc-virtual-viewport/div[1]/div/mat-tab-group/div/mat-tab-body[1]/div/storage-bucket-details-objects/cfc-left-panel-layout/cfc-panel-container/div/div/cfc-panel[2]/div/div/cfc-main-panel-content/cfc-panel-body/cfc-virtual-viewport/div[1]/div/storage-objects-table/storage-drop-target/div[2]/cfc-table/div[2]/cfc-table-columns-presenter-v2/div/div[3]/table/tbody"
            table_body = wait.until(EC.presence_of_element_located((By.XPATH, table_body_xpath)))
            rows = table_body.find_elements(By.TAG_NAME, "tr")
            row_count = len(rows)
            print(f"ID {id_value} มี {row_count} row ใน table GCS")

            # ถ้าไม่ตรงกัน ให้จดลง txt
            if row_count != folder_count:
                print(f"⚠️ จำนวน subfolder และ row ไม่ตรงกันสำหรับ ID {id_value}")
                txt_file.write(f"{id_value}\n")
            else:
                print(f"✅ จำนวน subfolder และ row ตรงกันสำหรับ ID {id_value}")

                # --- เข้า subfolder แต่ละอันและนับไฟล์ ---
                for sf in [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]:
                    sf_path = os.path.join(folder_path, sf)
                    file_count = len([f for f in os.listdir(sf_path) if os.path.isfile(os.path.join(sf_path, f))])
                    print(f"   📁 Subfolder '{sf}' มี {file_count} ไฟล์")

                    # --- กดลิงก์ใน table ที่ตรงกับชื่อ subfolder ---
                    try:
                        link_xpath = f"//table/tbody/tr/td[2]/div/div/a[contains(text(), '{sf}')]"
                        link_el = wait.until(EC.element_to_be_clickable((By.XPATH, link_xpath)))
                        link_el.click()
                        time.sleep(3)  # รอโหลดหน้า table

                        # นับจำนวน row ใน table ใหม่
                        sub_table_body_xpath = table_body_xpath
                        sub_table_body = wait.until(EC.presence_of_element_located((By.XPATH, sub_table_body_xpath)))
                        sub_rows = sub_table_body.find_elements(By.TAG_NAME, "tr")
                        sub_row_count = len(sub_rows)
                        print(f"     🔹 ใน table ของ '{sf}' มี {sub_row_count} row")

                        # --- เช็คไฟล์ vs row --- 
                        if sub_row_count != file_count:
                            print(f"⚠️ จำนวนไฟล์และ row ใน '{sf}' ไม่ตรงกัน")
                            txt_file.write(f"{id_value} - {sf} - folder files:{file_count}, table rows:{sub_row_count}\n")

                        driver.back()
                        time.sleep(3)

                    except Exception as e:
                        print(f"❌ ไม่สามารถเข้าลิงก์หรือ count rows ของ '{sf}': {e}")

        except Exception as e:
            print(f"❌ ไม่สามารถนับ row ได้สำหรับ ID {id_value}: {e}")

print("เสร็จสิ้นการประมวลผลทั้งหมด")
driver.quit()
