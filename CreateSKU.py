# -*- coding: utf-8 -*-
import time
import logging
import pandas as pd
import getpass
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementNotInteractableException,
    NoSuchElementException,
)

# ---------------- CONFIGURATION ----------------
EXCEL_FILE_PATH = r"C:\Users\tanapat\Downloads\Create DEE.xlsx"
SKU_OUTPUT_FILE = r"C:\Users\tanapat\Downloads\sku_result.xlsx"

# Login & Base URLs
LOGIN_URL = "https://base.chanintr.com/login"
BASE_PRODUCT_URL = "https://base.chanintr.com/brand/420/products"

# ---------------- LOGGING SETUP ----------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - [%(levelname)s] - %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


class ChanintrBot:
    def __init__(self):
        self.driver = self._init_driver()
        # เพิ่มเวลา Wait มาตรฐานเป็น 20 วินาที
        self.wait = WebDriverWait(self.driver, 20)

    def _init_driver(self):
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-notifications")
        # chrome_options.add_argument("--headless") # เปิดบรรทัดนี้ถ้าไม่อยากเห็นหน้าต่าง Browser
        return webdriver.Chrome(options=chrome_options)

    def close(self):
        self.driver.quit()

    # ---------------- HELPERS ----------------
    def _fill(self, xpath, value):
        """กรอกข้อมูลลงในช่อง Input อย่างปลอดภัย"""
        if pd.isna(value) or str(value).strip() == "":
            return  # ข้ามถ้าไม่มีข้อมูล

        try:
            el = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            try:
                el.clear()
            except:
                self.driver.execute_script("arguments[0].value = '';", el)

            el.send_keys(str(value))
        except TimeoutException:
            logger.warning(f"⚠️ หาช่องกรอกข้อมูลไม่เจอ: {xpath}")

    def _click(self, xpath, timeout=20):
        """คลิกปุ่ม โดยรองรับทั้งการคลิกปกติและ JavaScript Click"""
        try:
            wait_custom = WebDriverWait(self.driver, timeout)
            el = wait_custom.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            try:
                el.click()
            except:
                self.driver.execute_script("arguments[0].click();", el)
        except TimeoutException:
            logger.warning(f"⚠️ คลิกปุ่มไม่ได้ (Timeout): {xpath}")
            raise  # ส่ง Error ออกไปเพื่อให้รู้ว่ากดไม่ได้

    def _select_dropdown_option(self, container_xpath, value):
        """เลือกตัวเลือกจาก Dropdown"""
        if pd.isna(value):
            return

        target_text = str(value).strip().lower()
        try:
            # รอให้รายการ Dropdown โหลด
            li_elements = self.wait.until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, f"{container_xpath}/ul/li")
                )
            )

            for li in li_elements:
                if target_text in li.text.strip().lower():
                    self.driver.execute_script("arguments[0].click();", li)
                    return  # เจอแล้วจบเลย

            # ถ้าไม่เจอ เลือกตัวแรก
            logger.warning(f"ไม่เจอตัวเลือก '{value}' เลือกรายการแรกแทน")
            self.driver.execute_script("arguments[0].click();", li_elements[0])

        except TimeoutException:
            logger.warning(f"⚠️ รายการ Dropdown ไม่ขึ้น: {container_xpath}")

    def save_result_to_excel(self, vendor_item, sku_text):
        """บันทึกผลลัพธ์ลง Excel"""
        try:
            new_data = {
                "Vendor Item Number": [vendor_item],
                "Generated SKU": [sku_text],
                "Timestamp": [time.strftime("%Y-%m-%d %H:%M:%S")],
            }
            df_new = pd.DataFrame(new_data)

            if os.path.exists(SKU_OUTPUT_FILE):
                df_existing = pd.read_excel(SKU_OUTPUT_FILE)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_combined.to_excel(SKU_OUTPUT_FILE, index=False)
            else:
                df_new.to_excel(SKU_OUTPUT_FILE, index=False)

            logger.info(f"💾 Saved SKU: {sku_text}")
        except Exception as e:
            logger.error(f"❌ Failed to save Excel: {e}")

    # ---------------- MAIN LOGIC ----------------
    def login_manual_fallback(self, email, password):
        """Login Google แบบมีระบบรอคนกดเองถ้าบอทกดไม่ได้"""
        logger.info("กำลังเข้าสู่หน้า Login...")
        self.driver.get("https://accounts.google.com/signin/v2/identifier")

        try:
            self._fill("//input[@type='email' or @id='identifierId']", email)
            self.driver.find_element(By.ID, "identifierNext").click()
            time.sleep(3)  # รอ Animation เปลี่ยนหน้า
            self._fill("//input[@type='password']", password)
            self.driver.find_element(By.ID, "passwordNext").click()
            time.sleep(5)
        except Exception:
            logger.warning("⚠ บอทกรอกรหัสไม่ได้ (อาจติด Captcha) กรุณากรอกเองใน Browser")

    def login_base(self):
        logger.info("เข้าสู่หน้า Base Chanintr...")
        self.driver.get(LOGIN_URL)
        try:
            # รอให้ปุ่ม Sign in Google ขึ้นแล้วกด
            btn = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., 'Sign in with Google')]")
                )
            )
            btn.click()

            # **สำคัญ** ให้เวลา User ในการยืนยันตัวตนหรือรอหน้าเว็บโหลดเสร็จ
            logger.info("⏳ รอการ Login... (ถ้าต้องยืนยันตัวตน 2FA ให้ทำใน Browser ได้เลย)")
            time.sleep(10)
            logger.info("Login สำเร็จ (สมมติ)")
        except TimeoutException:
            logger.error("หาปุ่ม Sign in with Google ไม่เจอ")

    def process_sku_creation(self, row):
        vendor_item = str(row["Vendor Item Number"]).strip()
        logger.info(f"🔹 กำลังทำรายการ: {vendor_item}")

        # 1. ค้นหา Product
        search_url = (
            f"{BASE_PRODUCT_URL}"
            f"?currentPage=1&searchText={vendor_item}"
            "&directionUser=DESC&sortBy=title&direction=ASC&isSearch=false"
        )
        self.driver.get(search_url)

        # 2. เลือก Product ที่ตรงกัน (Exact Match)
        try:
            # รอให้ List สินค้าขึ้น
            product_items = self.wait.until(
                EC.presence_of_all_elements_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/ul/li",
                    )
                )
            )

            product_clicked = False
            for li in product_items:
                try:
                    # ดึง text ของ Vendor Item Number ในการ์ดสินค้า
                    item_text = li.find_element(
                        By.XPATH, "./a/section/div[3]"
                    ).text.strip()
                    if item_text == vendor_item:
                        self.driver.execute_script(
                            "arguments[0].click();", li.find_element(By.XPATH, "./a")
                        )
                        product_clicked = True
                        break
                except:
                    continue

            if not product_clicked:
                logger.warning(f"❌ ไม่พบสินค้าที่ตรงกับ: {vendor_item}")
                return

        except TimeoutException:
            logger.warning(f"❌ ไม่พบรายการสินค้าใดๆ สำหรับ: {vendor_item}")
            return

        # 3. ไป Tab SKU/Price
        self._click("/html/body/div/div/section/section/div/ul/li[5]/a")

        # 4. กดปุ่ม Create
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div/div/div/a"
        )

        # 5. กดเลือก Variant/Option (Trigger Pop-up)
        self._click(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[1]/div/div[1]/div"
        )
        time.sleep(1)  # รอ Animation Pop-up นิดหน่อย

        # ==========================================
        # 🟢 LOGIC ตรวจสอบ Pop-up (ตามที่คุณต้องการ)
        # ==========================================
        target_option_xpath = (
            "/html/body/div/div/section/section/div/div/div[2]/div/div[2]/div[1]/li/div"
        )

        try:
            # ใช้ Short Wait (3 วิ) เพื่อเช็คว่ามีตัวเลือกไหม
            short_wait = WebDriverWait(self.driver, 3)
            short_wait.until(
                EC.presence_of_element_located((By.XPATH, target_option_xpath))
            )

            # --> กรณีเจอ (มีตัวเลือก)
            logger.info("✅ เจอตัวเลือก -> กดเลือก -> กดปุ่ม Confirm (Button 2)")
            self._click(target_option_xpath)
            time.sleep(0.5)
            # กดปุ่ม Confirm
            self._click(
                "/html/body/div/div/section/section/div/div/div[2]/div/div[3]/button[2]"
            )

        except TimeoutException:
            # --> กรณีไม่เจอ (Timeout)
            logger.info("⚠️ ไม่เจอตัวเลือก -> กดปุ่ม Cancel (Button 1)")
            # กดปุ่ม Cancel
            self._click(
                "/html/body/div/div/section/section/div/div/div[2]/div/div[3]/button[1]"
            )

        # ==========================================

        # รอให้ฟอร์มโหลดกลับมาเป็นปกติ
        time.sleep(1)

        # 6. กรอกข้อมูล Vendor (Dropdown)
        try:
            self._click(
                "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[3]/div/div[1]/span[1]"
            )
            self._click(f"//ul/li[contains(normalize-space(.), '{vendor_item}')]")
        except Exception:
            logger.warning("⚠ เลือก Vendor ใน Dropdown ไม่ได้ หรือมีค่าอยู่แล้ว")

        # 7. กรอก AP Number
        if pd.notna(row.get("AP Number")):
            try:
                # เปิด Dropdown AP
                self._click(
                    "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[4]/div/div[1]/span[1]"
                )
                # พิมพ์ค่าลงไป
                self._fill(
                    "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[4]/div/div[2]/div/div[1]/div/input",
                    row["AP Number"],
                )
                # เลือกจาก List ที่เด้งมา
                self._click(
                    f"//ul/li[contains(normalize-space(.), '{row['AP Number']}')]",
                    timeout=5,
                )
            except Exception:
                logger.warning("⚠ กรอก AP Number ไม่สำเร็จ")

        # 8. Purchasing Condition (Text Area)
        self._fill(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[10]/div/textarea",
            row.get("Purchasing Condition"),
        )

        # 9. Order Status (Dropdown)
        try:
            self._click(
                "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[11]/div/div[1]/span[1]"
            )
            self._select_dropdown_option(
                "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[11]/div/div[2]/div/div",
                row.get("SKU Status"),
            )
        except:
            pass

        # 10. Unit Price
        self._fill(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[3]/div[2]/div/div[4]/div/div/input",
            row.get("Unit Price"),
        )

        # 11. Description
        self._fill(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[4]/div[3]/div/textarea",
            row.get("Description For Vendor"),
        )

        # 12. กดปุ่ม SAVE
        logger.info("กดบันทึก...")
        self._click("/html/body/div/div/section/section/section[1]/div/button")

        # 13. ดึง SKU ID ที่ถูกสร้าง
        try:
            # รอให้ H1 ปรากฏ (ปกติหลังจาก Save จะเด้งไปหน้า View)
            sku_element = self.wait.until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[1]/div[1]/div[1]/h1",
                    )
                )
            )
            sku_text = sku_element.text.strip()

            if sku_text:
                self.save_result_to_excel(vendor_item, sku_text)
            else:
                logger.error("SKU Text ว่างเปล่า")

        except TimeoutException:
            logger.error("❌ หา SKU ID ไม่เจอหลังจากการบันทึก (อาจจะบันทึกไม่ผ่าน)")


# ---------------- RUN ----------------
if __name__ == "__main__":
    print("=== Chanintr SKU Bot ===")

    # รับค่า Login
    g_email = input("Google Email: ").strip()
    g_pass = getpass.getpass("Google Password: ").strip()

    if not os.path.exists(EXCEL_FILE_PATH):
        logger.error(f"ไม่พบไฟล์ Excel: {EXCEL_FILE_PATH}")
        exit()

    df = pd.read_excel(EXCEL_FILE_PATH)
    logger.info(f"โหลดข้อมูล {len(df)} แถว")

    bot = ChanintrBot()

    try:
        # Login
        bot.login_manual_fallback(g_email, g_pass)
        bot.login_base()

        # Loop ทำงาน
        for index, row in df.iterrows():
            try:
                bot.process_sku_creation(row)
                time.sleep(2)
            except Exception as e:
                logger.error(f"ข้ามรายการที่ {index+1} เนื่องจาก Error: {e}")
                continue

    except KeyboardInterrupt:
        logger.info("หยุดการทำงานโดยผู้ใช้")
    finally:
        logger.info("เสร็จสิ้น ปิดโปรแกรม...")
        bot.close()
