# -*- coding: utf-8 -*-
import os
import time
import logging
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import shutil

# Configure logging with UTF-8 encoding
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('occhio_downloader.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# Set console to handle UTF-8 properly on Windows
import sys
if sys.platform.startswith('win'):
    try:
        # Try to set console to UTF-8
        import os
        os.system('chcp 65001 > nul')
    except:
        pass
logger = logging.getLogger(__name__)

class OcchioDownloader:
    def __init__(self, base_folder="downloads"):
        self.base_folder = Path(base_folder)
        self.temp_folder = self.base_folder / "temp_download"
        self.base_folder.mkdir(exist_ok=True)
        self.temp_folder.mkdir(exist_ok=True)
        
        self.exclude_subheaders = {"Energy Labels", "Documents of conformity"}
        self.valid_non_pdf_keywords = {"3DS", "DWG"}
        self.url = "https://www.occhio.com/service/downloads"
        
        self.driver = self._setup_driver()
    
    def _setup_driver(self):
        """Setup Chrome WebDriver with download preferences"""
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": str(self.temp_folder.absolute()),
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": True,
            "plugins.always_open_pdf_externally": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        
        return webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()), 
            options=options
        )
    
    def sanitize_filename(self, name):
        """Remove invalid characters from filename"""
        invalid_chars = r'\/:*?"<>|'
        return "".join(c for c in name if c not in invalid_chars)
    
    def wait_for_download(self, timeout=120, check_interval=0.5):
        """Wait for download to complete"""
        end_time = time.time() + timeout
        while time.time() < end_time:
            files = list(self.temp_folder.iterdir())
            if not files:
                time.sleep(check_interval)
                continue
            
            downloading = [f for f in files if f.suffix in {'.crdownload', '.tmp'}]
            if not downloading:
                return True
            time.sleep(check_interval)
        
        logger.warning(f"Download timeout after {timeout} seconds")
        return False
    
    def move_latest_file(self, dest_folder):
        """Move the latest downloaded file to destination folder and wait for completion"""
        files = [f for f in self.temp_folder.iterdir() 
                if f.is_file() and f.suffix not in {'.crdownload', '.tmp'}]
        
        if not files:
            logger.warning("No files found to move")
            return False
        
        latest_file = max(files, key=lambda f: f.stat().st_mtime)
        dest_folder = Path(dest_folder)
        dest_folder.mkdir(parents=True, exist_ok=True)
        
        dest_path = dest_folder / latest_file.name
        
        # Move file and wait for completion
        try:
            shutil.move(str(latest_file), str(dest_path))
            
            # Verify file was moved successfully
            if dest_path.exists() and not latest_file.exists():
                logger.info(f"[SUCCESS] Successfully moved {latest_file.name} to {dest_folder}")
                
                # Clean up temp folder to ensure it's empty for next download
                self._cleanup_temp_folder()
                return True
            else:
                logger.error(f"[ERROR] File move verification failed for {latest_file.name}")
                return False
                
        except Exception as e:
            logger.error(f"[ERROR] Error moving file {latest_file.name}: {e}")
            return False
    
    def is_valid_non_pdf(self, link_text):
        """Check if non-PDF file should be downloaded"""
        return any(keyword.lower() in link_text.lower() 
                  for keyword in self.valid_non_pdf_keywords)
    
    def _cleanup_temp_folder(self):
        """Clean up remaining files in temp folder"""
        try:
            for file in self.temp_folder.iterdir():
                if file.is_file():
                    file.unlink()
                    logger.debug(f"Cleaned up temp file: {file.name}")
        except Exception as e:
            logger.warning(f"Warning during temp cleanup: {e}")
    
    def wait_for_file_move_completion(self, timeout=30):
        """Wait for file move operation to complete"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            # Check if temp folder is empty (all files moved)
            temp_files = [f for f in self.temp_folder.iterdir() 
                         if f.is_file() and f.suffix not in {'.crdownload', '.tmp'}]
            if not temp_files:
                return True
            time.sleep(0.1)
        
        logger.warning(f"File move timeout after {timeout} seconds")
        return False
        """Check if non-PDF file should be downloaded"""
        return any(keyword.lower() in link_text.lower() 
                  for keyword in self.valid_non_pdf_keywords)
    
    def download_pdf(self, href, dest_folder):
        """Download PDF file by opening in new tab"""
        main_window = self.driver.current_window_handle
        try:
            logger.info(f"[PDF] Starting download: {href}")
            self.driver.execute_script("window.open(arguments[0]);", href)
            WebDriverWait(self.driver, 10).until(
                lambda d: len(d.window_handles) > 1
            )
            
            new_windows = [w for w in self.driver.window_handles if w != main_window]
            if new_windows:
                new_window = new_windows[0]
                self.driver.switch_to.window(new_window)
                
                if self.wait_for_download():
                    logger.info("[PDF] Download completed, moving file...")
                    if self.move_latest_file(dest_folder):
                        # Wait to ensure file move is completely finished
                        self.wait_for_file_move_completion()
                        logger.info(f"[SUCCESS] PDF download and move completed: {href}")
                    else:
                        logger.error(f"[ERROR] Failed to move PDF: {href}")
                else:
                    logger.error(f"[ERROR] Failed to download PDF: {href}")
                
                if new_window in self.driver.window_handles:
                    self.driver.close()
            else:
                logger.error(f"[ERROR] No new tab opened for: {href}")
                
        except Exception as e:
            logger.error(f"[ERROR] Error downloading PDF {href}: {e}")
        finally:
            self.driver.switch_to.window(main_window)
            # Extra wait to ensure everything is settled before next download
            time.sleep(2)
    
    def download_non_pdf(self, link, dest_folder):
        """Download non-PDF file by clicking link"""
        try:
            logger.info(f"[NON-PDF] Starting download: {link.text}")
            link.click()
            time.sleep(1)
            
            if self.wait_for_download():
                logger.info("[NON-PDF] Download completed, moving file...")
                if self.move_latest_file(dest_folder):
                    # Wait to ensure file move is completely finished
                    self.wait_for_file_move_completion()
                    logger.info(f"[SUCCESS] Non-PDF download and move completed: {link.text}")
                else:
                    logger.error(f"[ERROR] Failed to move non-PDF: {link.text}")
            else:
                logger.error(f"[ERROR] Failed to download non-PDF: {link.text}")
            
            # Extra wait to ensure everything is settled before next download
            time.sleep(2)
                
        except Exception as e:
            logger.error(f"[ERROR] Error downloading non-PDF {link.text}: {e}")
    
    def process_links(self, links, dest_folder):
        """Process a list of download links"""
        for link in links:
            try:
                link_text = link.text.strip()
                href = link.get_attribute("href")
                
                if not link_text or not href:
                    continue
                
                logger.info(f"Processing: {link_text} -> {href}")
                
                if href.lower().endswith(".pdf"):
                    self.download_pdf(href, dest_folder)
                elif self.is_valid_non_pdf(link_text):
                    self.download_non_pdf(link, dest_folder)
                else:
                    logger.info(f"Skipping: {link_text} (not PDF or valid non-PDF)")
                    
            except Exception as e:
                logger.error(f"Error processing link: {e}")
    
    def process_section(self, header_button, header_text):
        """Process a main section with potential subsections"""
        header_folder = self.base_folder / self.sanitize_filename(header_text)
        header_folder.mkdir(exist_ok=True)
        
        logger.info(f"\n=== Processing: {header_text} ===")
        
        # Click to expand section
        self.driver.execute_script("arguments[0].scrollIntoView(true);", header_button)
        self.driver.execute_script("arguments[0].click();", header_button)
        time.sleep(1)
        
        # Look for subsections
        sub_headers = self.driver.find_elements(
            By.XPATH, "//div[contains(@class,'LinkList js-LinkList')]//h"
        )
        
        if sub_headers:
            self._process_subsections(sub_headers, header_folder)
        else:
            self._process_direct_links(header_button, header_folder)
        
        # Collapse section
        self.driver.execute_script("arguments[0].click();", header_button)
        time.sleep(0.3)
    
    def _process_subsections(self, sub_headers, header_folder):
        """Process subsections within a main section"""
        for sub in sub_headers:
            try:
                sub_text = sub.text.strip()
                if not sub_text or sub_text in self.exclude_subheaders:
                    continue
                
                logger.info(f" - Processing subsection: {sub_text}")
                sub_folder = header_folder / self.sanitize_filename(sub_text)
                
                parent_div = sub.find_element(
                    By.XPATH, "./ancestor::div[contains(@class,'LinkList js-LinkList')]"
                )
                links = parent_div.find_elements(By.XPATH, ".//a")
                self.process_links(links, sub_folder)
                
            except Exception as e:
                logger.error(f"Error processing subsection {sub_text}: {e}")
    
    def _process_direct_links(self, header_button, header_folder):
        """Process direct links under a main section (no subsections)"""
        try:
            accordion_div = header_button.find_element(
                By.XPATH, "./following-sibling::div[contains(@class,'LinkList js-LinkList')]"
            )
            links = accordion_div.find_elements(By.XPATH, ".//a")
            self.process_links(links, header_folder)
            
        except NoSuchElementException:
            logger.warning("No direct links found in this section")
        except Exception as e:
            logger.error(f"Error processing direct links: {e}")
    
    def run(self):
        """Main execution method"""
        try:
            logger.info(f"Opening URL: {self.url}")
            self.driver.get(self.url)
            
            # Wait for page to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, '//h2/button'))
            )
            
            # Get all main section headers
            download_headers = self.driver.find_elements(By.XPATH, '//h2/button')
            logger.info(f"Found {len(download_headers)} main sections")
            
            # Process each section
            for i in range(len(download_headers)):
                # Re-find elements to avoid stale references
                all_headers = self.driver.find_elements(By.XPATH, '//h2/button')
                if i < len(all_headers):
                    header_button = all_headers[i]
                    header_text = header_button.text.strip()
                    self.process_section(header_button, header_text)
            
            logger.info("\n✅ Download process completed successfully!")
            
        except TimeoutException:
            logger.error("Page load timeout - check internet connection or URL")
        except Exception as e:
            logger.error(f"Unexpected error: {e}")
        finally:
            self.cleanup()
    
    def cleanup(self):
        """Clean up resources"""
        try:
            if self.driver.window_handles:
                logger.info(f"Page title: {self.driver.title}")
            
            # Clean up temp folder
            if self.temp_folder.exists():
                try:
                    shutil.rmtree(self.temp_folder)
                    logger.info("🧹 Cleaned up temporary folder")
                except Exception as e:
                    logger.warning(f"Warning cleaning temp folder: {e}")
                    # Try to clean files individually if rmtree fails
                    self._cleanup_temp_folder()
                
        except Exception as e:
            logger.warning(f"Error during cleanup: {e}")
        finally:
            self.driver.quit()


if __name__ == "__main__":
    downloader = OcchioDownloader()
    downloader.run()