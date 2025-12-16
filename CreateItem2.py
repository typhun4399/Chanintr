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

LINK_PRODUCT = "https://base.chanintr.com/brand/420/products?currentPage=1&searchText={Vendor Item Number จาก Excel}&directionUser=DESC&sortBy=title&direction=ASC&isSearch=true"

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
def fill_input(xpath, value):
    el = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    el.clear()
    el.send_keys(str(value))


def js_click(el):
    driver.execute_script("arguments[0].click();", el)


def select_dropdown_by_text(text: str):
    """
    เลือก <li> ใน dropdown โดยใช้ข้อความ
    ถ้าไม่เจอ → เลือกตัวแรกแทน
    """
    text = text.strip()

    try:
        js_click(
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        f"//ul/li[contains(normalize-space(.), '{text}')]",
                    )
                )
            )
        )
    except TimeoutException:
        logging.warning(f"Dropdown value '{text}' not found → select first item")
        js_click(
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        ".dropdown-list-container ul li:first-child",
                    )
                )
            )
        )


try:
    # ---------- STEP 0: Load Excel ----------
    df = pd.read_excel(EXCEL_FILE)

    # ---------- STEP 1: Google Login ----------
    driver.get("https://accounts.google.com/signin/v2/identifier")

    wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//input[@type='email' or @id='identifierId']")
        )
    ).send_keys(GOOGLE_EMAIL)

    driver.find_element(By.ID, "identifierNext").click()

    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='password']"))
    ).send_keys(GOOGLE_PASSWORD)

    driver.find_element(By.ID, "passwordNext").click()
    time.sleep(3)

    # ---------- STEP 2: Base Login ----------
    driver.get("https://base.chanintr.com/login")

    wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Sign in with Google')]")
        )
    ).click()

    time.sleep(5)

    for _, row in df.iterrows():

        # ---------- STEP 3: Open Product Page ----------
        driver.get(LINK_PRODUCT)
        time.sleep(3)

        # ---------- STEP 4: Create Product ----------
        logging.info(f"Creating product: {row['Product Title']}")

        # --- Create button ---
        js_click(
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/button",
                    )
                )
            )
        )

        # --- Product Title ---
        title_input = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[1]/div[2]/div/input",
                )
            )
        )
        title_input.clear()
        title_input.send_keys(str(row["Product Title"]))

        # ===== Category =====
        js_click(
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[4]/div[2]/div/div[1]",
                    )
                )
            )
        )

        cat_input = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[4]/div[2]/div/div[2]/div/div[1]/div/input",
                )
            )
        )
        cat_input.send_keys(str(row["Category"]))

        select_dropdown_by_text(str(row["Category"]))

        # ===== Not For Customer =====
        if str(row.get("Not For Customer", "")).upper() == "TRUE":
            js_click(
                wait.until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[5]/div[2]/div[2]/div",
                        )
                    )
                )
            )

        # ===== Order Status =====
        js_click(
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[6]/div[2]/div/div[1]",
                    )
                )
            )
        )

        order_input = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[6]/div[2]/div/div[2]/div/div[1]/div/input",
                )
            )
        )
        order_input.send_keys(str(row["Order Status"]))

        select_dropdown_by_text(str(row["Order Status"]))

        # --- Save ---
        js_click(
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[3]/button[2]",
                    )
                )
            )
        )

        time.sleep(2)

        # ---------- STEP 5: Rooms & Styles ----------
        logging.info("Updating Rooms & Styles")

        # 🔹 เปิด modal
        btn_open = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/button",
                )
            )
        )
        js_click(btn_open)
        time.sleep(1)

        # ===== Rooms =====

        rooms_dropdown = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div/div[2]/span",
                )
            )
        )
        js_click(rooms_dropdown)

        rooms_input = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div/div[3]/div/div/div/input",
                )
            )
        )
        rooms_input.send_keys(str(row["Rooms"]))

        li_rooms = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div/div[3]/div/ul/li[1]",
                )
            )
        )
        js_click(li_rooms)

        # ===== Styles =====

        styles_dropdown = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[2]/div[2]/div/div[2]/span",
                )
            )
        )
        js_click(styles_dropdown)

        styles_input = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[2]/div[2]/div/div[3]/div/div/div/input",
                )
            )
        )
        styles_input.send_keys(str(row["Styles"]))

        li_styles = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[2]/div[2]/div/div[3]/div/ul/li[1]",
                )
            )
        )
        js_click(li_styles)

        # 🔹 Save
        btn_save = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/span/button[2]",
                )
            )
        )
        js_click(btn_save)

        time.sleep(2)

        # ---------- STEP 6: Dimension Edit ----------
        if str(row.get("Size type", "")).upper() == "R":
            btn_edit = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[1]/div/button[2]",
                    )
                )
            )
            js_click(btn_edit)

            time.sleep(1)

            btn_round = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/ul/li[2]",
                    )
                )
            )
            js_click(btn_round)

            time.sleep(1)

            fill_input(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[1]/div/input",
                row["Width"],
            )

            fill_input(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[2]/div/input",
                row["Height"],
            )

            fill_input(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div[2]/div/input",
                row["Product Weight (kg)"],
            )

            if str(row.get("Auto Package Size", "")).upper() == "TRUE":
                # Auto Package
                btn_Auto = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[2]/div[1]/span/button",
                        )
                    )
                )
                js_click(btn_Auto)

                time.sleep(1)
                value = row.get("Volume (CBM)")

                if value is not None and str(value).strip() != "":

                    element = driver.find_element(
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[5]/div[2]/div/input",
                    )
                    driver.execute_script("arguments[0].value = '';", element)

                    fill_input(
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[5]/div[2]/div/input",
                        row["Volume (CBM)"],
                    )

            else:

                fill_input(
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[3]/div[2]/div/div[1]/div/input",
                    row["Package Size Width"],
                )

                fill_input(
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[3]/div[2]/div/div[2]/div/input",
                    row["Package Size Depth"],
                )

                fill_input(
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[3]/div[2]/div/div[3]/div/input",
                    row["Package Size Height"],
                )

                fill_input(
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[5]/div[2]/div/input",
                    row["Volume (CBM)"],
                )

            save_btn = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[1]/div/span/button[2]",
                    )
                )
            )
            js_click(save_btn)

            logging.info("✅ Dimension saved successfully")
            time.sleep(2)
        # S
        else:
            btn_edit = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[1]/div/button[2]",
                    )
                )
            )
            js_click(btn_edit)

            time.sleep(1)

            fill_input(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[1]/div/input",
                row["Width"],
            )

            fill_input(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[2]/div/input",
                row["Depth"],
            )

            fill_input(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[3]/div/input",
                row["Height"],
            )

            fill_input(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div[2]/div/input",
                row["Product Weight (kg)"],
            )

            if str(row.get("Auto Package Size", "")).upper() == "TRUE":
                # Auto Package
                btn_Auto = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[2]/div[1]/span/button",
                        )
                    )
                )
                js_click(btn_Auto)

                time.sleep(1)

                if value is not None and str(value).strip() != "":

                    element = driver.find_element(
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[5]/div[2]/div/input",
                    )
                    driver.execute_script("arguments[0].value = '';", element)

                    fill_input(
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[5]/div[2]/div/input",
                        row["Volume (CBM)"],
                    )

            else:

                fill_input(
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[3]/div[2]/div/div[1]/div/input",
                    row["Package Size Width"],
                )

                fill_input(
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[3]/div[2]/div/div[2]/div/input",
                    row["Package Size Depth"],
                )

                fill_input(
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[3]/div[2]/div/div[3]/div/input",
                    row["Package Size Height"],
                )

                fill_input(
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[5]/div[2]/div/input",
                    row["Volume (CBM)"],
                )

            save_btn = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[1]/div/span/button[2]",
                    )
                )
            )
            js_click(save_btn)

        logging.info("✅ Dimension saved successfully")
        time.sleep(2)

        # ---------- STEP 7: Vendor Edit ----------

        # ---------- 1. กดปุ่มเปิด Vendor Item Number ----------
        js_click(
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[3]/div/button",
                    )
                )
            )
        )

        # ---------- 2. พิมพ์ Vendor Item Number ----------
        vendor_input = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[3]/div/div[1]/div/div[1]/div[2]/div/input",
                )
            )
        )

        vendor_input.clear()
        vendor_input.send_keys(str(row["Vendor Item Number"]))

        # ---------- 3. กดปุ่ม Confirm / Save ----------
        js_click(
            wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[3]/div/span/button[2]",
                    )
                )
            )
        )

        time.sleep(1)


except TimeoutException:
    logging.error("TimeoutException occurred", exc_info=True)

except Exception as e:
    logging.error(f"Unexpected error: {e}", exc_info=True)

finally:
    logging.info("Script finished")
