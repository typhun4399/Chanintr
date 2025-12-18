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
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException

# ---------------- CONFIGURATION ----------------
EXCEL_FILE_PATH = (
    r"C:\Users\tanapat\Downloads\create-consignment-template-20240716(1).xlsx"
)
# Changed output file to Excel (.xlsx)
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
        self.wait = WebDriverWait(self.driver, 20)

    def _init_driver(self):
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
        return webdriver.Chrome(options=chrome_options)

    def close(self):
        self.driver.quit()

    # ---------------- HELPERS ----------------
    def _fill(self, xpath, value):
        """
        Fills an input field.
        - Tries standard clear/send_keys.
        - Falls back to JavaScript if element is not interactable (common in animations).
        """
        try:
            # Wait for element to be clickable
            el = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))

            # Attempt to clear and fill normally
            try:
                el.clear()
            except (ElementNotInteractableException, Exception):
                # Fallback: Clear using JavaScript if standard clear fails
                self.driver.execute_script("arguments[0].value = '';", el)

            # Only send keys if value is valid
            if pd.notna(value) and str(value).strip() != "":
                el.send_keys(str(value))

        except TimeoutException:
            logger.warning(f"Timeout trying to fill element: {xpath}")

    def _click(self, xpath):
        try:
            el = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            self.driver.execute_script("arguments[0].click();", el)
        except TimeoutException:
            logger.warning(f"Timeout trying to click element: {xpath}")

    def _select_dropdown_option(self, container_xpath, value):
        """
        Selects an option from a dropdown list (ul/li) that contains the text value.
        Defaults to the first option if the specific value isn't found.
        """
        if pd.isna(value):
            return

        target_text = str(value).strip().lower()

        try:
            # Wait for list items to appear
            li_elements = self.wait.until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, f"{container_xpath}/ul/li")
                )
            )

            selected = False
            for li in li_elements:
                if target_text in li.text.strip().lower():
                    self.driver.execute_script("arguments[0].click();", li)
                    selected = True
                    break

            if not selected and li_elements:
                logger.warning(
                    f"Option '{value}' not found, selecting first available option."
                )
                self.driver.execute_script("arguments[0].click();", li_elements[0])

        except TimeoutException:
            logger.warning(f"Timeout waiting for dropdown options at {container_xpath}")

    def save_result_to_excel(self, vendor_item, sku_text):
        """
        Saves the Vendor Item Number and Generated SKU to an Excel file.
        Appends to existing file or creates a new one.
        """
        try:
            new_data = {
                "Vendor Item Number": [vendor_item],
                "Generated SKU": [sku_text],
                "Timestamp": [time.strftime("%Y-%m-%d %H:%M:%S")],
            }
            df_new = pd.DataFrame(new_data)

            if os.path.exists(SKU_OUTPUT_FILE):
                # Read existing file, append new data, and save back
                df_existing = pd.read_excel(SKU_OUTPUT_FILE)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_combined.to_excel(SKU_OUTPUT_FILE, index=False)
            else:
                # Create new file
                df_new.to_excel(SKU_OUTPUT_FILE, index=False)

            logger.info(f"💾 Saved SKU to Excel: {sku_text}")

        except Exception as e:
            logger.error(f"Failed to write SKU to Excel: {e}")

    # ---------------- AUTH FLOW ----------------
    def login_google(self, email, password):
        logger.info("Starting Google Login...")
        self.driver.get("https://accounts.google.com/signin/v2/identifier")

        self._fill("//input[@type='email' or @id='identifierId']", email)
        self.driver.find_element(By.ID, "identifierNext").click()

        # Wait for animation/transition from email to password (Increased to 5s)
        time.sleep(5)

        self._fill("//input[@type='password']", password)
        self.driver.find_element(By.ID, "passwordNext").click()

        time.sleep(5)

    def login_base(self):
        logger.info("Logging into Base Chanintr...")
        self.driver.get(LOGIN_URL)

        try:
            btn = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., 'Sign in with Google')]")
                )
            )
            btn.click()
            time.sleep(5)
            logger.info("Login successful.")
        except TimeoutException:
            logger.error("Could not find 'Sign in with Google' button.")

    # ---------------- SKU LOGIC ----------------
    def process_sku_creation(self, row):
        vendor_item = str(row["Vendor Item Number"]).strip()
        logger.info(f"Processing Vendor Item: {vendor_item}")

        # 1. Search for Product
        search_url = (
            f"{BASE_PRODUCT_URL}"
            f"?currentPage=1&searchText={vendor_item}"
            "&directionUser=DESC&sortBy=title&direction=ASC&isSearch=false"
        )
        self.driver.get(search_url)
        time.sleep(2)

        # 2. Find Correct Product (Exact Match Loop)
        try:
            product_items = self.wait.until(
                EC.presence_of_all_elements_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[3]/div[1]/section/div/ul/li",
                    )
                )
            )

            product_clicked = False
            for idx, li in enumerate(product_items, start=1):
                try:
                    # Check the vendor text inside the list item
                    vendor_text = li.find_element(
                        By.XPATH, "./a/section/div[3]"
                    ).text.strip()

                    if vendor_text == vendor_item:
                        logger.info(f"✅ Matched Vendor Item at index {idx}")
                        self.driver.execute_script(
                            "arguments[0].click();", li.find_element(By.XPATH, "./a")
                        )
                        product_clicked = True
                        break
                except Exception:
                    continue

            if not product_clicked:
                logger.warning(f"❌ No matching product found for: {vendor_item}")
                return

        except TimeoutException:
            logger.warning(f"❌ No product list items found for: {vendor_item}")
            return

        time.sleep(2)

        # ==============================================================================

        # 3. Navigate to SKU/Price Tab (Index 5)
        self._click("/html/body/div/div/section/section/div/ul/li[5]/a")

        # 4. Click 'Create' (or similar button inside the tab)
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div/div/div/a"
        )

        # 5. Fill Vendor
        # Open Dropdown
        self._click(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[3]/div/div[1]/span[1]"
        )
        # Select Vendor
        try:
            self._click(f"//ul/li[contains(normalize-space(.), '{vendor_item}')]")
        except Exception:
            logger.warning(f"Vendor '{vendor_item}' not found in dropdown.")

        # 6. Fill AP Number
        # Open AP Dropdown
        self._click(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[4]/div/div[1]/span[1]"
        )

        ap_number = row["AP Number"]
        self._fill(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[4]/div/div[2]/div/div[1]/div/input",
            ap_number,
        )

        # Select AP Result
        try:
            self._click(f"//ul/li[contains(normalize-space(.), '{ap_number}')]")
        except Exception:
            logger.warning(f"AP Number '{ap_number}' not found in dropdown list.")

        # 7. Purchasing Condition
        self._fill(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[10]/div/textarea",
            row["Purchasing Condition"],
        )

        # 8. Order Status
        # Open Dropdown
        self._click(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[11]/div/div[1]/span[1]"
        )
        time.sleep(0.5)
        # Select Option
        self._select_dropdown_option(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[2]/div[11]/div/div[2]/div/div",
            row["SKU Status"],
        )

        # 9. Unit Price
        self._fill(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[3]/div[2]/div/div[4]/div/div/input",
            row["Unit Price"],
        )

        # 10. Description (Vendor)
        self._fill(
            "/html/body/div/div/section/section/section[2]/div[1]/div/section[6]/div/textarea",
            row["Description For Vendor"],
        )

        # ==============================================================================

        # 11. Save
        self._click("/html/body/div/div/section/section/section[1]/div/button")
        time.sleep(2)

        # 12. Extract and Save SKU ID
        try:
            sku_element = self.wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div/div/section/section/section[1]/div[1]/div[1]/h1",
                    )
                )
            )
            sku_text = sku_element.text.strip()
            if sku_text:
                # Changed to save to Excel
                self.save_result_to_excel(vendor_item, sku_text)
            else:
                logger.warning("SKU text found but empty.")
        except Exception:
            logger.warning("⚠ Could not locate SKU text element after saving.")


# ---------------- MAIN EXECUTION ----------------
if __name__ == "__main__":
    print("=== Chanintr SKU Automation Bot ===")

    # Securely get credentials
    google_email = input("Google Email: ").strip()
    google_password = getpass.getpass("Google Password: ").strip()

    try:
        # Load Excel
        logger.info(f"Loading Excel from: {EXCEL_FILE_PATH}")
        df = pd.read_excel(EXCEL_FILE_PATH)

        # Initialize Bot
        bot = ChanintrBot()

        # Perform Login
        bot.login_google(google_email, google_password)
        bot.login_base()

        # Iterate Rows
        for index, row in df.iterrows():
            try:
                logger.info(f"--- Starting Row {index + 1} ---")
                bot.process_sku_creation(row)
                logger.info(f"--- Finished Row {index + 1} ---")
            except Exception as e:
                logger.error(f"Failed processing row {index + 1}: {e}", exc_info=True)
                # Continue to next row
                continue

    except Exception as e:
        logger.critical(f"Critical Error: {e}", exc_info=True)
    finally:
        # Ask user before closing to see results
        input("\nPress Enter to close browser and exit...")
        if "bot" in locals():
            bot.close()
