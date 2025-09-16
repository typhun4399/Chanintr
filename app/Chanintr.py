import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
import time
import logging
import threading

class TextHandler(logging.Handler):
    """
    คลาสนี้ใช้สำหรับส่ง log ไปยัง widget Text ของ Tkinter
    """
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.yview(tk.END)
        self.text_widget.after(0, append)

class AutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Web Automation Controller")
        self.geometry("800x750")

        # --- สร้าง Frame หลัก ---
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- ส่วนตั้งค่า ---
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.pack(fill=tk.X, expand=True, pady=5)
        settings_frame.columnconfigure(1, weight=1)

        # --- Credentials ---
        ttk.Label(settings_frame, text="Google Email:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.email_entry = ttk.Entry(settings_frame)
        self.email_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)

        ttk.Label(settings_frame, text="Google Password:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.password_entry = ttk.Entry(settings_frame, show="*")
        self.password_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)

        # --- Excel File ---
        ttk.Label(settings_frame, text="Excel File Path:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.file_path_entry = ttk.Entry(settings_frame)
        self.file_path_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
        self.browse_button = ttk.Button(settings_frame, text="Browse...", command=self.browse_file)
        self.browse_button.grid(row=2, column=2, sticky=tk.W, padx=5, pady=5)

        # --- Mode Selection ---
        ttk.Label(settings_frame, text="Operation Mode:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.mode_var = tk.StringVar()
        self.mode_combo = ttk.Combobox(settings_frame, textvariable=self.mode_var, values=["By Feature Code", "By Product ID"], state="readonly")
        self.mode_combo.grid(row=3, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        self.mode_combo.current(0)
        self.mode_combo.bind("<<ComboboxSelected>>", self.toggle_mode_fields)

        # --- Mode Specific Fields ---
        self.feature_code_frame = ttk.Frame(settings_frame)
        self.feature_code_frame.grid(row=4, column=0, columnspan=3, sticky=tk.EW)
        self.feature_code_frame.columnconfigure(1, weight=1)

        ttk.Label(self.feature_code_frame, text="Brand:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.brand_entry = ttk.Entry(self.feature_code_frame)
        self.brand_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        self.brand_entry.insert(0, "BAK")

        ttk.Label(self.feature_code_frame, text="Code Column Name:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.code_column_combo = ttk.Combobox(self.feature_code_frame, state="disabled")
        self.code_column_combo.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)

        self.product_id_frame = ttk.Frame(settings_frame)
        self.product_id_frame.grid(row=4, column=0, columnspan=3, sticky=tk.EW)
        self.product_id_frame.columnconfigure(1, weight=1)
        
        ttk.Label(self.product_id_frame, text="ID Column Name:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.id_column_combo = ttk.Combobox(self.product_id_frame, state="disabled")
        self.id_column_combo.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)

        self.toggle_mode_fields()

        # --- Control Button ---
        self.start_button = ttk.Button(main_frame, text="Start Automation", command=self.start_automation_thread)
        self.start_button.pack(pady=10, fill=tk.X)

        # --- Log Display ---
        log_frame = ttk.LabelFrame(main_frame, text="Logs", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', wrap=tk.WORD, height=20)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # --- ตั้งค่า logging ---
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(formatter)
        self.logger.addHandler(text_handler)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)
            self.scan_excel_columns(file_path)

    def scan_excel_columns(self, file_path):
        try:
            self.logger.info(f"Scanning columns from {file_path}...")
            df_cols = pd.read_excel(file_path, nrows=0).columns.tolist()
            
            self.code_column_combo.config(values=df_cols, state="readonly")
            self.id_column_combo.config(values=df_cols, state="readonly")

            # พยายามตั้งค่าเริ่มต้นถ้าเจอชื่อคอลัมน์ที่ตรงกัน
            if "Feature Code" in df_cols:
                self.code_column_combo.set("Feature Code")
            elif df_cols:
                self.code_column_combo.current(0)

            if "id" in df_cols:
                self.id_column_combo.set("id")
            elif df_cols:
                self.id_column_combo.current(0)
            
            self.logger.info("Excel columns loaded into dropdowns successfully.")

        except Exception as e:
            messagebox.showerror("Excel Read Error", f"Could not read columns from the selected file.\nPlease ensure it's a valid Excel file.\n\nError: {e}")
            self.file_path_entry.delete(0, tk.END)
            self.code_column_combo.config(values=[], state="disabled")
            self.id_column_combo.config(values=[], state="disabled")
            self.logger.error(f"Failed to read columns from Excel file: {e}")

    def toggle_mode_fields(self, event=None):
        if self.mode_var.get() == "By Feature Code":
            self.product_id_frame.grid_remove()
            self.feature_code_frame.grid()
        else:
            self.feature_code_frame.grid_remove()
            self.product_id_frame.grid()

    def start_automation_thread(self):
        self.start_button.config(state="disabled")
        thread = threading.Thread(target=self.run_automation)
        thread.daemon = True
        thread.start()
        
    def run_automation(self):
        try:
            # ดึงค่าทั้งหมดจาก UI
            email = self.email_entry.get()
            password = self.password_entry.get()
            excel_path = self.file_path_entry.get()
            mode = self.mode_var.get()

            # --- ตรวจสอบ Input พื้นฐาน ---
            if not all([email, password, excel_path]):
                messagebox.showerror("Input Error", "Please fill in Email, Password, and select an Excel file.")
                return

            # --- อ่านไฟล์ Excel ---
            try:
                df = pd.read_excel(excel_path)
                self.logger.info(f"Successfully read Excel file: {excel_path}")
            except FileNotFoundError:
                self.logger.error(f"Error: The file was not found at {excel_path}")
                return
            except Exception as e:
                self.logger.error(f"An error occurred while reading the Excel file: {e}")
                return

            # --- ตั้งค่า Selenium ---
            chrome_options = Options()
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--disable-notifications")
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 20)

            # --- Login ---
            if not self.perform_login(driver, wait, email, password):
                driver.quit()
                return

            # --- เลือกโหมดการทำงาน ---
            if mode == "By Feature Code":
                brand = self.brand_entry.get()
                code_column = self.code_column_combo.get()
                if not all([brand, code_column]):
                    messagebox.showerror("Input Error", "Please provide Brand and select a Code Column.")
                    driver.quit()
                    return
                if code_column not in df.columns:
                    self.logger.error(f"Error: Column '{code_column}' not found in the Excel file.")
                    driver.quit()
                    return
                self.run_feature_code_mode(driver, wait, df, brand, code_column)

            elif mode == "By Product ID":
                id_column = self.id_column_combo.get()
                if not id_column:
                     messagebox.showerror("Input Error", "Please select an ID Column.")
                     driver.quit()
                     return
                if id_column not in df.columns:
                    self.logger.error(f"Error: Column '{id_column}' not found in the Excel file.")
                    driver.quit()
                    return
                self.run_product_id_mode(driver, wait, df, id_column)

        except Exception as e:
            self.logger.error(f"An unexpected error occurred in the main process: {e}")
        finally:
            if 'driver' in locals() and driver:
                driver.quit()
                self.logger.info("Browser closed.")
            self.start_button.config(state="normal")
            self.logger.info("----- Automation Finished -----")


    def perform_login(self, driver, wait, email, password):
        try:
            # --- Google Login ---
            self.logger.info("Attempting to log in to Google...")
            driver.get("https://accounts.google.com/signin/v2/identifier")
            
            email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email']")))
            email_input.clear()
            email_input.send_keys(email)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='identifierNext']"))).click()
            
            password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
            password_input.clear()
            password_input.send_keys(password)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='passwordNext']"))).click()
            
            self.logger.info("Google login successful. Waiting 5 seconds...")
            time.sleep(5)

            # --- Base Login ---
            self.logger.info("Navigating to base.chanintr.com...")
            driver.get("https://base.chanintr.com/brand") # ไปที่หน้า brand โดยตรง
            button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/section/div[3]/button")))
            button.click()
            self.logger.info("Clicked the login button on base.chanintr.com.")
            time.sleep(3)

            second_login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div/div/div[2]/div/div/div[1]/form/span/section/div/div/div/div/ul/li[1]")))
            second_login_button.click()
            self.logger.info("Account selection button clicked successfully.")
            time.sleep(5)
            
            return True # Login สำเร็จ

        except TimeoutException:
            self.logger.error("Login failed! Could not find an element or the request timed out.")
            return False # Login ไม่สำเร็จ

    def run_feature_code_mode(self, driver, wait, df, brand, code_column):
        self.logger.info(f"--- Running in 'By Feature Code' mode for brand: {brand} ---")
        try:
            # Find the brand search input field and type the brand name
            self.logger.info("Searching for the brand...")
            search_input_xpath = "/html/body/div/div/section/section/div[1]/section/div[2]/div[1]/div/input"
            search_input = wait.until(EC.element_to_be_clickable((By.XPATH, search_input_xpath)))
            search_input.clear()
            search_input.send_keys(brand)
            self.logger.info(f"Typed '{brand}' into the search field.")
            time.sleep(2) 

            # Click on the brand in the filtered list
            list_item_xpath = f"//li[.//span[text()='{brand}']]"
            brand_element_in_list = wait.until(EC.element_to_be_clickable((By.XPATH, list_item_xpath)))
            brand_element_in_list.click()
            self.logger.info(f"Successfully clicked on the list item for brand: {brand}")
            time.sleep(5) 

            # Click on the 'Feature' tab
            products_tab_xpath = "/html/body/div/div/section/section/div/ul/li[2]"
            products_tab = wait.until(EC.element_to_be_clickable((By.XPATH, products_tab_xpath)))
            products_tab.click()
            self.logger.info("Clicked on the second tab (Feature).")
            time.sleep(10)

            # Click on the toggle search
            discontinued_toggle_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[3]/div[2]/div[2]"
            discontinued_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, discontinued_toggle_xpath)))
            discontinued_toggle.click()
            self.logger.info("Clicked on the 'Discontinued' toggle button.")
            time.sleep(2)
            
            # --- Loop through each product code from the Excel file ---
            self.logger.info(f"Starting to process codes from the '{code_column}' column...")
            product_search_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[3]/div[3]/div[1]/div[1]/input"

            for index, code in enumerate(df[code_column]):
                try:
                    code_str = str(code).strip()
                    if not code_str or pd.isna(code):
                        self.logger.warning(f"Row {index + 2}: Skipping empty or invalid code.")
                        continue

                    # Find the product search input field
                    product_search_input = wait.until(EC.element_to_be_clickable((By.XPATH, product_search_xpath)))
                    
                    product_search_input.clear()
                    product_search_input.send_keys(code_str)
                    self.logger.info(f"Row {index + 2}: Searched for code: {code_str}")
                    time.sleep(3) # Wait for table to filter
                    
                    # STEP 1: Click the "Edit" button for the product
                    edit_button_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[3]/ul/li/div[7]/button"
                    edit_button = wait.until(EC.element_to_be_clickable((By.XPATH, edit_button_xpath)))
                    edit_button.click()
                    self.logger.info(f"Row {index + 2}: Clicked 'Edit' button.")
                    time.sleep(2) # Wait for modal to open

                    # STEP 2: Click the 'Discontinued' switch/toggle inside the modal
                    discontinue_switch_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[1]/div/div/div[2]/div/div[2]/div[3]/div[2]/div/div"
                    discontinue_switch = wait.until(EC.element_to_be_clickable((By.XPATH, discontinue_switch_xpath)))
                    discontinue_switch.click()
                    self.logger.info(f"Row {index + 2}: Toggled the 'Discontinued' switch.")
                    time.sleep(1) # Short pause after clicking the switch

                    # STEP 3: Click the 'Save' button to confirm the change
                    save_button_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[1]/div/div/div[2]/div/div[3]/button[2]"
                    save_button = wait.until(EC.element_to_be_clickable((By.XPATH, save_button_xpath)))
                    save_button.click()
                    self.logger.info(f"Row {index + 2}: Clicked 'Save'. Product '{code_str}' successfully updated.")
                    time.sleep(1) # Wait for save to process and confirmation to appear
                    
                    # STEP 4: Click 'OK' or 'Close' on the confirmation pop-up
                    confirmation_button_xpath = "/html/body/div/div/section/section/section[3]/div[1]/section/div/div[2]/div/div[1]/div/div/div[2]/div/div[2]/button[2]"
                    confirmation_ok_button = wait.until(EC.element_to_be_clickable((By.XPATH, confirmation_button_xpath)))
                    confirmation_ok_button.click()
                    self.logger.info(f"Row {index + 2}: Closed the confirmation dialog.")
                    time.sleep(2) # Wait for the main page to be fully interactive again

                except (TimeoutException, NoSuchElementException, StaleElementReferenceException):
                    self.logger.error(f"Row {index + 2}: Could not process code '{code_str}'. Product may not exist or an error occurred. Skipping.")
                    try:
                        product_search_input = wait.until(EC.element_to_be_clickable((By.XPATH, product_search_xpath)))
                        product_search_input.clear()
                    except Exception as clear_error:
                        self.logger.error(f"Could not clear search box. Refreshing page to recover: {clear_error}.")
                        driver.refresh()
                        time.sleep(5)
                        # ต้องกดเข้ามาใหม่หลัง refresh
                        products_tab = wait.until(EC.element_to_be_clickable((By.XPATH, products_tab_xpath)))
                        products_tab.click()
                        time.sleep(3)
                        discontinued_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, discontinued_toggle_xpath)))
                        discontinued_toggle.click()
                        time.sleep(2)
                    continue

        except TimeoutException as e:
            self.logger.error(f"A critical element was not found, stopping script: {e}")
            return
    
    def run_product_id_mode(self, driver, wait, df, id_column):
        self.logger.info(f"--- Running in 'By Product ID' mode ---")
        for index, row in df.iterrows():
            product_id = row[id_column]
            if pd.isna(product_id):
                self.logger.warning(f"Row {index + 2}: Skipping empty product ID.")
                continue

            product_id = int(product_id)
            url = f"https://base.chanintr.com/brand/10/product/{product_id}/overview"
            self.logger.info(f"Row {index + 2}: Opening URL for product ID {product_id}")
            driver.get(url)

            try:
                # กดปุ่ม Edit Product
                product_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/section/section/section[2]/div/div[2]/button")))
                product_button.click()
                self.logger.info(f"Clicked 'Edit Product' button for ID {product_id}.")
                time.sleep(2)

                # กด Dropdown ของ Status
                status_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/section/section/div[1]/div/div[2]/div/div[2]/div[6]/div[2]/div/div[1]")))
                status_dropdown.click()
                self.logger.info(f"Clicked status dropdown for ID {product_id}.")
                time.sleep(2)

                # เลือก Discontinued
                discontinued_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/section/section/div[1]/div/div[2]/div/div[2]/div[6]/div[2]/div/div[2]/div/div[2]/ul/li[4]")))
                discontinued_button.click()
                self.logger.info(f"Selected 'Discontinued' for ID {product_id}.")
                time.sleep(2)

                # กดปุ่ม Save
                final_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/section/section/div[1]/div/div[2]/div/div[3]/button[2]")))
                final_button.click()
                self.logger.info(f"Clicked 'Save' button for ID {product_id}. Update complete.")

            except TimeoutException:
                self.logger.warning(f"Could not find an element for product ID {product_id}. It might already be discontinued or the page structure changed. Skipping.")

            time.sleep(3)

if __name__ == "__main__":
    app = AutomationApp()
    app.mainloop()

