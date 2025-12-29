# -*- coding: utf-8 -*-
import time
import logging
import pandas as pd
import getpass
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ---------------- CONFIGURATION ----------------
EXCEL_FILE_PATH = r"C:\Users\tanapat\Downloads\Create DEE.xlsx"
LINK_PRODUCT_CREATE = "https://base.chanintr.com/brand/420/products?currentPage=1&directionUser=DESC&sortBy=title&direction=ASC&isSearch=false"
LOGIN_URL = "https://base.chanintr.com/login"

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
        if pd.isna(value):
            return
        try:
            el = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            el.clear()
            el.send_keys(str(value))
        except TimeoutException:
            logger.warning(f"Timeout trying to fill element: {xpath}")

    def _click(self, xpath):
        try:
            el = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            self.driver.execute_script("arguments[0].click();", el)
        except TimeoutException:
            logger.warning(f"Timeout trying to click element: {xpath}")

    def _select_dropdown_by_text(self, text):
        """Clicks a list item containing the specific text, defaults to first item if not found."""
        if pd.isna(text):
            return
        text = str(text).strip()
        try:
            xpath = f"//ul/li[contains(normalize-space(.), '{text}')]"
            self._click(xpath)
        except Exception:
            logger.warning(f"Dropdown value '{text}' not found, selecting first item.")
            self._click(".dropdown-list-container ul li:first-child")

    # ---------------- AUTH FLOW ----------------
    def login_google(self, email, password):
        logger.info("Starting Google Login...")
        self.driver.get("https://accounts.google.com/signin/v2/identifier")

        self._fill("//input[@type='email' or @id='identifierId']", email)
        self.driver.find_element(By.ID, "identifierNext").click()

        self._fill("//input[@type='password']", password)
        self.driver.find_element(By.ID, "passwordNext").click()

        # Give Google a moment to process redirect
        time.sleep(5)

    def login_base(self):
        logger.info("Logging into Base Chanintr...")
        self.driver.get(LOGIN_URL)

        # Click "Sign in with Google"
        try:
            btn = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., 'Sign in with Google')]")
                )
            )
            btn.click()
            time.sleep(5)  # Wait for redirect
            logger.info("Login successful.")
        except TimeoutException:
            logger.error("Could not find 'Sign in with Google' button.")

    # ---------------- PRODUCT LOGIC ----------------
    def create_product(self, row):
        logger.info(f"Processing: {row['Product Title']}")
        self.driver.get(LINK_PRODUCT_CREATE)
        time.sleep(3)

        # 1. Click Create Button
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/button"
        )

        # 2. Basic Info
        self._fill(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[1]/div[2]/div/input",
            row["Product Title"],
        )

        # 3. Category
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[4]/div[2]/div/div[1]"
        )
        self._fill(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[4]/div[2]/div/div[2]/div/div[1]/div/input",
            row["Category"],
        )
        self._select_dropdown_by_text(row["Category"])

        # 4. Not For Customer
        if str(row.get("Not For Customer", "")).upper() == "TRUE":
            self._click(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[5]/div[2]/div[2]/div"
            )

        # 5. Order Status
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[6]/div[2]/div/div[1]"
        )
        self._fill(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[2]/div[6]/div[2]/div/div[2]/div/div[1]/div/input",
            row["Order Status"],
        )
        self._select_dropdown_by_text(row["Order Status"])

        # 6. Save Basic Info
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div/div[2]/div/div[3]/button[2]"
        )
        time.sleep(2)

    def update_rooms_and_styles(self, row):
        logger.info("Updating Rooms & Styles...")
        # Open Modal
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/button"
        )
        time.sleep(1)

        # Rooms
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div/div[2]/span"
        )
        self._fill(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div/div[3]/div/div/div/input",
            row["Rooms"],
        )
        # Select first result
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div/div[3]/div/ul/li[1]"
        )

        # Styles
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[2]/div[2]/div/div[2]/span"
        )
        self._fill(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[2]/div[2]/div/div[3]/div/div/div/input",
            row["Styles"],
        )
        # Select first result
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/div[2]/div[2]/div[2]/div/div[3]/div/ul/li[1]"
        )

        # Save Modal
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[1]/div[1]/div/span/button[2]"
        )
        time.sleep(2)

    def update_dimensions(self, row):
        logger.info("Updating Dimensions...")
        # Click Edit Button
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[1]/div/button[2]"
        )
        time.sleep(1)

        size_type = str(row.get("Size type", "")).upper()

        if size_type == "R":
            # Click "Round" shape button
            self._click(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/ul/li[2]"
            )
            time.sleep(1)

            # Fill Round Dimensions
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[1]/div/input",
                row["Width"],
            )
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[2]/div/input",
                row["Height"],
            )
            # Weight
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div[2]/div/input",
                row["Product Weight (kg)"],
            )

        else:
            # Standard/Rectangular logic (Default)
            time.sleep(1)
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[1]/div/input",
                row["Width"],
            )
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[2]/div/input",
                row["Depth"],
            )
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div[3]/div/input",
                row["Height"],
            )
            # Weight
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div[2]/div/input",
                row["Product Weight (kg)"],
            )

        # Handle Package Sizing (Common logic)
        self._handle_package_sizing(row)

        self._fill(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[4]/div[2]/div/input",
            row["Product Package Weight"],
        )

        # Click Save
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[1]/div/span/button[2]"
        )
        time.sleep(2)

    def _handle_package_sizing(self, row):
        """Helper to handle the lower section of the dimension form"""
        is_auto_package = str(row.get("Auto Package Size", "")).upper() == "TRUE"

        if is_auto_package:
            # Click Auto Button (Updated to find by text "Calculate Package Size")
            self._click("//button[contains(., 'Calculate Package Size')]")
            time.sleep(1)

            # If Volume is provided manually even when auto is on
            vol = row.get("Volume (CBM)")
            if pd.notna(vol) and str(vol).strip() != "":
                # Clear existing auto-calculated value using JS
                vol_input_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[5]/div[2]/div/input"
                try:
                    el = self.driver.find_element(By.XPATH, vol_input_xpath)
                    self.driver.execute_script("arguments[0].value = '';", el)
                    self._fill(vol_input_xpath, vol)
                except NoSuchElementException:
                    pass

        else:
            # Manual Package Size
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[3]/div[2]/div/div[1]/div/input",
                row["Package Size Width"],
            )
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[3]/div[2]/div/div[2]/div/input",
                row["Package Size Depth"],
            )
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[3]/div[2]/div/div[3]/div/input",
                row["Package Size Height"],
            )
            self._fill(
                "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[2]/div/div[2]/div[2]/div[4]/div[1]/div[5]/div[2]/div/input",
                row["Volume (CBM)"],
            )

    def update_vendor_info(self, row):
        logger.info("Updating Vendor Info...")
        # Open Vendor Edit
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[3]/div/button"
        )

        # Fill Number
        self._fill(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[3]/div/div[1]/div/div[1]/div[2]/div/input",
            row["Vendor Item Number"],
        )

        # Save
        self._click(
            "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div[3]/div/span/button[2]"
        )
        time.sleep(1)


# ---------------- MAIN EXECUTION ----------------
if __name__ == "__main__":
    print("=== Chanintr Product Automation Bot ===")

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
                bot.create_product(row)
                bot.update_rooms_and_styles(row)
                bot.update_dimensions(row)
                bot.update_vendor_info(row)
                logger.info(f"--- Finished Row {index + 1} ---")
            except Exception as e:
                logger.error(f"Failed processing row {index + 1}: {e}", exc_info=True)
                # Optional: Continue to next row even if one fails
                continue

    except Exception as e:
        logger.critical(f"Critical Error: {e}", exc_info=True)
    finally:
        # Ask user before closing to see results
        input("\nPress Enter to close browser and exit...")
        if "bot" in locals():
            bot.close()
