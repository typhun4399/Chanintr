from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd  # เพิ่ม import pandas สำหรับจัดการ Excel


def run_scraper():
    # ตั้งค่า Driver (Chrome)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless') # ปิด comment บรรทัดนี้ถ้าต้องการรันแบบไม่เปิดหน้าต่าง

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

    all_products = []  # ตัวแปรสำหรับเก็บข้อมูลทั้งหมดเพื่อลง Excel

    try:
        # 1. เปิด URL
        url = "https://www.bakerfurniture.com/design-story/designers-and-collections/kara-mann/"
        print(f"กำลังเข้าสู่เว็บไซต์: {url}")
        driver.get(url)

        # รอให้หน้าเว็บโหลด Element หลัก (รอสูงสุด 10 วินาที)
        wait = WebDriverWait(driver, 10)

        # --- ส่วนที่ 1: กดปุ่มขยายข้อมูลทุกปุ่มก่อน (Click All View More Buttons) ---
        print("-" * 30)
        try:
            print("กำลังค้นหาปุ่มที่มีข้อความ 'View ... More' ...")

            try:
                wait_btn = WebDriverWait(driver, 5)
                wait_btn.until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            "//button[contains(., 'View') and contains(., 'More')]",
                        )
                    )
                )
            except:
                pass

            expand_buttons = driver.find_elements(
                By.XPATH, "//button[contains(., 'View') and contains(., 'More')]"
            )

            if len(expand_buttons) > 0:
                print(
                    f"พบปุ่ม 'View ... More' ทั้งหมด {len(expand_buttons)} ปุ่ม กำลังทยอยกด..."
                )

                for i, btn in enumerate(expand_buttons, 1):
                    try:
                        driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'center'});", btn
                        )
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", btn)
                        print(f"  > กดปุ่มที่ {i} เรียบร้อย")
                        time.sleep(2)
                    except Exception as inner_e:
                        print(f"  (!) กดปุ่มที่ {i} ไม่สำเร็จ: {inner_e}")

                print(">>> กดครบทุกปุ่มแล้ว รอข้อมูลโหลดเพิ่ม 3 วินาที...")
                time.sleep(3)
            else:
                print(">>> ไม่พบปุ่มที่มีคำว่า 'View ... More' ในหน้านี้")

        except Exception as btn_error:
            print(f"(!) หมายเหตุ: {btn_error}")

        # --- ส่วนที่ 2: หา Container สินค้าและดึงข้อมูล (ชื่อ + รหัส) ---

        target_class_selector = ".dropBlocks"

        # ค้นหาโซนสินค้าหลัก
        parent_elements = driver.find_elements(By.CSS_SELECTOR, target_class_selector)
        parent_count = len(parent_elements)

        print("-" * 30)
        print(f"พบโซนสินค้า (dropBlocks) ทั้งหมด: {parent_count} จุด")
        print("-" * 30)

        # Selector สำหรับหา 'กล่องสินค้าแต่ละชิ้น' (li)
        # เพื่อที่เราจะดึงทั้ง ชื่อ และ รหัส ที่อยู่ภายในกล่องเดียวกันได้
        item_selector = "div.vr.vr_5x > ul > li"

        for index, parent in enumerate(parent_elements, start=1):
            print(f"กำลังดึงข้อมูลจากโซนที่ {index}...")

            # หาลิสต์สินค้า (li) ทั้งหมดในโซนนี้
            items = parent.find_elements(By.CSS_SELECTOR, item_selector)

            if len(items) == 0:
                print("  - ไม่พบสินค้าในโซนนี้")
                continue

            for item in items:
                try:
                    # 1. ดึงชื่อสินค้า (Name)
                    # Path: div > div.feature-label > a
                    try:
                        name_elem = item.find_element(
                            By.CSS_SELECTOR, "div > div.feature-label a"
                        )
                        name_text = name_elem.text.strip()
                    except:
                        name_text = "N/A"

                    # 2. ดึงรหัสสินค้า (Code)
                    # ตาม Path ที่คุณขอ: ... > li > div > div:nth-child(4) > a
                    try:
                        code_elem = item.find_element(
                            By.CSS_SELECTOR, "div > div:nth-child(4) > a"
                        )
                        code_text = code_elem.text.strip()
                    except:
                        code_text = "N/A"  # อาจจะไม่มีรหัส หรือโครงสร้างต่างออกไป

                    # ถ้ามีชื่อสินค้า ให้เก็บลง list
                    if name_text and name_text != "N/A":
                        print(f"  > สินค้า: {name_text} | รหัส: {code_text}")
                        all_products.append(
                            {"Product Name": name_text, "Product Code": code_text}
                        )

                except Exception as item_error:
                    continue  # ข้ามชิ้นที่มีปัญหา

        print("-" * 30)
        print(f"สรุป: ดึงข้อมูลได้ทั้งหมด {len(all_products)} รายการ")

        # --- ส่วนที่ 3: บันทึกลง Excel ---
        if all_products:
            print("กำลังบันทึกไฟล์ Excel...")
            df = pd.DataFrame(all_products)
            filename = "Kara Mann.xlsx"
            df.to_excel(filename, index=False)
            print(f"✅ บันทึกเสร็จสมบูรณ์: {filename}")
        else:
            print("ไม่มีข้อมูลสำหรับบันทึก")

    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")

    finally:
        # ปิด Browser
        print("จบการทำงาน กำลังปิด Browser...")
        driver.quit()


if __name__ == "__main__":
    run_scraper()
