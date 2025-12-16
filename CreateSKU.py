# -*- coding: utf-8 -*-
import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ---------------- CONFIG ----------------
print("GOOGLE_EMAIL:")
GOOGLE_EMAIL = input().strip()

print("GOOGLE_PASSWORD:")
GOOGLE_PASSWORD = input().strip()

EXCEL_FILE = r"C:\Users\phunk\Downloads\create-consignment-template-20240716.xlsx"
SKU_OUTPUT_FILE = r"C:\Users\phunk\Downloads\sku_result.txt"

# ---------------- LOGGING ----------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# ---------------- CHROME OPTIONS ----------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)


# ---------------- HELPERS ----------------
def js_click(el):
    driver.execute_script("arguments[0].click();", el)


def fill_input(xpath, value):
    el = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    el.clear()
    el.send_keys(str(value))


def save_sku_to_file(sku_text):
    with open(SKU_OUTPUT_FILE, "a", encoding="utf-8") as f:
        f.write(sku_text + "\n")


try:
    # ---------- LOAD EXCEL ----------
    df = pd.read_excel(EXCEL_FILE)

    # ---------- GOOGLE LOGIN ----------
    driver.get("https://accounts.google.com/signin/v2/identifier")

    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='email']"))
    ).send_keys(GOOGLE_EMAIL)
    driver.find_element(By.ID, "identifierNext").click()

    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='password']"))
    ).send_keys(GOOGLE_PASSWORD)
    driver.find_element(By.ID, "passwordNext").click()
    time.sleep(3)

    # ---------- BASE LOGIN ----------
    driver.get("https://base.chanintr.com/login")

    wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Sign in with Google')]")
        )
    ).click()
    time.sleep(5)

    # ---------- LOOP PRODUCT ----------
    for _, row in df.iterrows():

        vendor = str(row["Vendor Item Number"]).strip()
        logging.info(f"Process Vendor Item Number: {vendor}")

        link_product = (
            "https://base.chanintr.com/brand/420/products"
            f"?currentPage=1&searchText={vendor}"
            "&directionUser=DESC&sortBy=title&direction=ASC&isSearch=false"
        )

        driver.get(link_product)
        time.sleep(2)

        # ---------- OPEN PRODUCT ----------
        js_click(
            wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/ul/li/a",
                    )
                )
            )
        )
        time.sleep(1)

        # ---------- SKU IMAGE ----------
        js_click(
            wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "/html/body/div/div/section/section/div/ul/li[5]/a")
                )
            )
        )

        js_click(
            wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div/div/div/a",
                    )
                )
            )
        )

        # ================= VENDOR =================
        js_click(
            wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[3]/div/div[1]/span[1]",
                    )
                )
            )
        )

        js_click(
            wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//ul/li[contains(normalize-space(.), '{vendor}')]")
                )
            )
        )

        # ================= AP =================
        js_click(
            wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[4]/div/div[1]/span[1]",
                    )
                )
            )
        )

        fill_input(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[4]/div/div[2]/div/div[1]/div/input",
            row["AP Number"],
        )

        js_click(
            wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        f"//ul/li[contains(normalize-space(.), '{row['AP Number']}')]",
                    )
                )
            )
        )

        # ================= PURCHASING CONDITION =================
        fill_input(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[10]/div/textarea",
            row["Purchasing Condition"],
        )

        # ================= ORDER STATUS =================
        js_click(
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[11]/div/div[1]/span[1]",
                    )
                )
            )
        )
        time.sleep(0.5)

        li_elements = wait.until(
            EC.presence_of_all_elements_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[11]/div/div[2]/div/div/ul/li",
                )
            )
        )

        target_status = str(row["SKU Status"]).strip().lower()
        selected = False

        for li in li_elements:
            if target_status in li.text.strip().lower():
                js_click(li)
                selected = True
                break

        if not selected:
            logging.warning("⚠ SKU Status not found → set first item active")
            js_click(li_elements[0])

        # ================= UNIT PRICE =================
        fill_input(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[3]/div[2]/div/div[4]/div/div/input",
            row["Unit Price"],
        )

        # ================= DESCRIPTION =================
        fill_input(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[6]/div/textarea",
            row["Description"],
        )

        # ================= SAVE =================
        js_click(
            wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[1]/div/button",
                    )
                )
            )
        )
        time.sleep(1)

        # ================= SAVE SKU TO NOTEPAD =================
        try:
            sku_text = wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[1]/div[1]/div[1]/h1",
                    )
                )
            ).text.strip()

            save_sku_to_file(sku_text)
            logging.info(f"💾 Saved SKU → {sku_text}")

        except Exception:
            logging.warning("⚠ Cannot read SKU text")

except TimeoutException:
    logging.error("❌ TimeoutException occurred", exc_info=True)

except Exception as e:
    logging.error(f"❌ Unexpected error: {e}", exc_info=True)

finally:
    logging.info("✅ Script finished")
    # driver.quit()
