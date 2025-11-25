import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
import time
import logging

# --- CONFIG ---
excel_path = r"C:\Users\tanapat\Downloads\to check BAK & MGC discontinuation 090925.xlsx"
# This is the name of the column in your Excel file that contains the product codes.
code_column_name = "Feature Code" 
brand = "BAK"

GOOGLE_EMAIL = "tanapat@chanintr.com"
GOOGLE_PASSWORD = "Qwerty12345$$"

# --- Read Excel ---
try:
    df = pd.read_excel(excel_path)
    if code_column_name not in df.columns:
        logging.error(f"Error: Column '{code_column_name}' not found in the Excel file.")
        exit()
except FileNotFoundError:
    logging.error(f"Error: The file was not found at {excel_path}")
    exit()


# --- Set up Chrome ---
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20) 

# --- Set up logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Hardcoded Google Login ---
try:
    driver.get("https://accounts.google.com/signin/v2/identifier")
    
    email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
    email_input.clear()
    email_input.send_keys(GOOGLE_EMAIL)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='identifierNext']"))).click()
    
    password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
    password_input.clear()
    password_input.send_keys(GOOGLE_PASSWORD)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='passwordNext']"))).click()
    
    logging.info("Login successful... waiting 5 seconds to proceed.")
    time.sleep(5)
except TimeoutException:
    logging.error("Login failed! Could not find input fields or the request timed out.")
    driver.quit()
    exit()

# --- Open base.chanintr.com and click the login button ---
try:
    driver.get("https://base.chanintr.com/brand")
    button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/section/div[3]/button")))
    button.click()
    logging.info("Clicked the button on base.chanintr.com/login successfully.")
    time.sleep(3)

    # Click the button to select the account
    second_login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div/div/div[2]/div/div/div[1]/form/span/section/div/div/div/div/ul/li[1]")))
    second_login_button.click()
    logging.info("Account selection button clicked successfully.")
    time.sleep(5)

except TimeoutException:
    logging.error("Could not find the login button or the request timed out.")
    driver.quit()
    exit()

# --- Main automation block ---
try:
    # Find the brand search input field and type the brand name
    logging.info("Searching for the brand...")
    search_input_xpath = "/html/body/div/div/section/section/div[1]/section/div[2]/div[1]/div/input"
    search_input = wait.until(EC.element_to_be_clickable((By.XPATH, search_input_xpath)))
    search_input.clear()
    search_input.send_keys(brand)
    logging.info(f"Typed '{brand}' into the search field.")
    time.sleep(2) 

    # Click on the brand in the filtered list
    list_item_xpath = f"//li[.//span[text()='{brand}']]"
    brand_element_in_list = wait.until(EC.element_to_be_clickable((By.XPATH, list_item_xpath)))
    brand_element_in_list.click()
    logging.info(f"Successfully clicked on the list item for brand: {brand}")
    time.sleep(5) 

    # Click on the 'Feature' tab
    products_tab_xpath = "/html/body/div/div/section/section/div/ul/li[2]"
    products_tab = wait.until(EC.element_to_be_clickable((By.XPATH, products_tab_xpath)))
    products_tab.click()
    logging.info("Clicked on the second tab (Feature).")
    time.sleep(10)

    # Click on the toggle search
    discontinued_toggle_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[3]/div[2]/div[2]"
    discontinued_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, discontinued_toggle_xpath)))
    discontinued_toggle.click()
    logging.info("Clicked on the 'Discontinued' toggle button.")
    time.sleep(2)

    # --- Loop through each product code from the Excel file ---
    logging.info(f"Starting to process codes from the '{code_column_name}' column...")
    product_search_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[3]/div[3]/div[1]/div[1]/input"

    for index, code in enumerate(df[code_column_name]):
        try:
            code_str = str(code).strip()
            if not code_str or pd.isna(code):
                logging.warning(f"Row {index + 2}: Skipping empty or invalid code.")
                continue

            # Find the product search input field
            product_search_input = wait.until(EC.element_to_be_clickable((By.XPATH, product_search_xpath)))
            
            # Clear the field and type the new code
            product_search_input.clear()
            product_search_input.send_keys(code_str)
            logging.info(f"Row {index + 2}: Searched for code: {code_str}")
            time.sleep(3) # Wait for table to filter

            # --- NEW: Check if the product is already marked as discontinued ---
            try:
                status_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[3]/ul/li/div[6]/p/strong"
                # Use a short wait time for this check to avoid slowing down the loop
                status_element = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, status_xpath)))
                
                # If the element is found, check its text
                if "Discont" in status_element.text:
                    logging.info(f"Row {index + 2}: Product '{code_str}' is already discontinued. Skipping.")
                    continue # Skip to the next item in the loop
            except TimeoutException:
                # This is the expected outcome if the "Discontinued" label is NOT present.
                logging.info(f"Row {index + 2}: Product '{code_str}' is not discontinued. Proceeding with update.")
                pass # Continue with the normal script execution

            # STEP 1: Click the "Edit" button for the product
            edit_button_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[3]/ul/li/div[7]/button"
            edit_button = wait.until(EC.element_to_be_clickable((By.XPATH, edit_button_xpath)))
            edit_button.click()
            logging.info(f"Row {index + 2}: Clicked 'Edit' button.")
            time.sleep(2) # Wait for modal to open

            # STEP 2: Click the 'Discontinued' switch/toggle inside the modal
            discontinue_switch_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[1]/div/div/div[2]/div/div[2]/div[3]/div[2]/div/div"
            discontinue_switch = wait.until(EC.element_to_be_clickable((By.XPATH, discontinue_switch_xpath)))
            discontinue_switch.click()
            logging.info(f"Row {index + 2}: Toggled the 'Discontinued' switch.")
            time.sleep(1) # Short pause after clicking the switch

            # STEP 3: Click the 'Save' button to confirm the change
            save_button_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[1]/div/div/div[2]/div/div[3]/button[2]"
            save_button = wait.until(EC.element_to_be_clickable((By.XPATH, save_button_xpath)))
            save_button.click()
            logging.info(f"Row {index + 2}: Clicked 'Save'. Product '{code_str}' successfully updated.")
            time.sleep(1) # Wait for save to process and confirmation to appear
            
            # STEP 4: Click 'OK' or 'Close' on the confirmation pop-up
            confirmation_button_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[1]/div/div/div[2]/div/div[2]/button[2]"
            confirmation_ok_button = wait.until(EC.element_to_be_clickable((By.XPATH, confirmation_button_xpath)))
            confirmation_ok_button.click()
            logging.info(f"Row {index + 2}: Closed the confirmation dialog.")
            time.sleep(2) # Wait for the main page to be fully interactive again

        except (TimeoutException, NoSuchElementException, StaleElementReferenceException):
            logging.error(f"Row {index + 2}: Could not process code '{code_str}'. Product may not exist or an error occurred. Skipping.")
            try:
                product_search_input = wait.until(EC.element_to_be_clickable((By.XPATH, product_search_xpath)))
                product_search_input.clear()
            except Exception as clear_error:
                logging.error(f"Could not clear search box due to: {clear_error}. Refreshing page to recover.")
                driver.refresh()
                time.sleep(5) 
                products_tab = wait.until(EC.element_to_be_clickable((By.XPATH, products_tab_xpath)))
                products_tab.click()
                time.sleep(3)
                discontinued_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, discontinued_toggle_xpath)))
                discontinued_toggle.click()
                time.sleep(2)
            continue 

except TimeoutException as e:
    logging.error(f"A critical element was not found or the request timed out, stopping script: {e}")
    driver.quit()
    exit()


# --- Close Browser ---
logging.info("Script finished successfully. Closing the browser.")
driver.quit()