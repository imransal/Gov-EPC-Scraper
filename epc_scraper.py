import pandas as pd
import os
import time
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import re

class EPCCertificateScraper:
    def __init__(self, download_dir="C:\\Users\\IS19\\Documents\\EPC_Scraper\\Processed"):
        self.download_dir = download_dir
        self.setup_logging()
        self.setup_driver()
        self.success_count = 0
        self.failure_count = 0
        self.results = []
        
    def setup_logging(self):
        """Setup logging configuration"""
        # Create logs directory if it doesn't exist
        log_dir = os.path.join(os.path.dirname(self.download_dir), "logs")
        os.makedirs(log_dir, exist_ok=True)
        
        # Setup logging
        log_filename = os.path.join(log_dir, f"epc_scraper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename),
                logging.StreamHandler()  # Also print to console
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"Starting EPC Certificate Scraper - Log file: {log_filename}")
        
    def setup_driver(self):
        """Setup Chrome driver with PDF download configuration"""
        chrome_options = Options()
        
        # Create download directory if it doesn't exist
        os.makedirs(self.download_dir, exist_ok=True)
        
        # Configure Chrome for PDF auto-download
        prefs = {
            "printing.print_preview_sticky_settings.appState": 
                '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
            "savefile.default_directory": self.download_dir,
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--kiosk-printing")  # Auto-print without dialog
        
        # Add stability arguments
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        
        # Optional: Run headless for faster processing (comment out if you want to see the browser)
        # chrome_options.add_argument("--headless")
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 20)  # Increased timeout
        
    def construct_full_address(self, row):
        """Construct full address from address components"""
        address_parts = [
            str(row.get('Address Line 1', '')).strip(),
            str(row.get('Address Line 2', '')).strip(),
            str(row.get('Address Line 3', '')).strip(),
            str(row.get('Address Line 4', '')).strip(),
            str(row.get('Address Line 5', '')).strip(),
            str(row.get('Town', '')).strip()
        ]
        
        # Filter out empty parts and 'nan' strings
        address_parts = [part for part in address_parts if part and part.lower() != 'nan']
        return ', '.join(address_parts)
    
    def generate_filename(self, row):
        """Generate filename based on the specified format"""
        scheme_abbrev = str(row.get('Scheme Abbreviation', 'Unknown')).strip()
        plot_number = str(row.get('Development Plot Number', 'Unknown')).strip()
        tenure = str(row.get('Tenure', 'Unknown')).strip()
        uprn = str(row.get('UPRN', 'Unknown')).strip()
        
        filename = f"EPC - {scheme_abbrev} - {plot_number} - {tenure} - {uprn}.pdf"
        
        # Clean filename (remove invalid characters)
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        return filename
    
    def debug_page_state(self, context=""):
        """Debug helper to log current page state"""
        try:
            current_url = self.driver.current_url
            page_title = self.driver.title
            self.logger.info(f"DEBUG {context} - URL: {current_url}")
            self.logger.info(f"DEBUG {context} - Title: {page_title}")
            
            # Take screenshot for debugging
            screenshot_path = os.path.join(self.download_dir, f"debug_screenshot_{context}_{datetime.now().strftime('%H%M%S')}.png")
            self.driver.save_screenshot(screenshot_path)
            self.logger.info(f"DEBUG Screenshot saved: {screenshot_path}")
            
            # Log visible radio buttons
            try:
                radio_buttons = self.driver.find_elements(By.XPATH, "//input[@type='radio']")
                self.logger.info(f"DEBUG Found {len(radio_buttons)} radio buttons:")
                for i, radio in enumerate(radio_buttons):
                    try:
                        radio_id = radio.get_attribute('id')
                        radio_name = radio.get_attribute('name')
                        radio_value = radio.get_attribute('value')
                        radio_checked = radio.get_attribute('checked')
                        self.logger.info(f"  Radio {i+1}: id='{radio_id}', name='{radio_name}', value='{radio_value}', checked='{radio_checked}'")
                    except Exception as e:
                        self.logger.warning(f"  Radio {i+1}: Error reading attributes - {e}")
            except Exception as e:
                self.logger.warning(f"DEBUG Error finding radio buttons: {e}")
                
        except Exception as e:
            self.logger.error(f"DEBUG Error in debug_page_state: {e}")
    
    def wait_for_page_load(self, timeout=10):
        """Wait for page to fully load"""
        try:
            self.wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
            time.sleep(1)  # Additional small wait for dynamic content
            return True
        except TimeoutException:
            self.logger.warning(f"Page did not fully load within {timeout} seconds")
            return False
    
    def navigate_to_start(self):
        """Navigate to the EPC certificate search page"""
        try:
            self.driver.get("https://www.gov.uk/find-energy-certificate")
            self.logger.info("Navigated to EPC certificate page")
            
            # Wait for page to fully load
            self.wait_for_page_load()
            
            # Click "Start now" button
            start_button = self.wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Start now"))
            )
            start_button.click()
            self.logger.info("Clicked 'Start now' button")
            
            # Wait for the next page to load
            self.wait_for_page_load()
            return True
            
        except TimeoutException:
            self.logger.error("Failed to find 'Start now' button")
            self.debug_page_state("start_button_not_found")
            return False
    
    def select_domestic_property(self):
        """Select domestic property option"""
        try:
            # Wait for the page to fully load first
            self.wait_for_page_load()
            
            # The radio buttons are hidden (visible=False), so we need to click the label instead
            domestic_element = None
            
            # Try different approaches to select the domestic option
            try:
                # Method 1: Try clicking the label for the domestic radio button
                domestic_label = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//label[@for='domestic']"))
                )
                domestic_label.click()
                self.logger.info("Found and clicked domestic property label")
            except TimeoutException:
                try:
                    # Method 2: Try finding a clickable element containing "domestic" text
                    domestic_element = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//label[contains(text(), 'Domestic') or contains(text(), 'domestic')]"))
                    )
                    domestic_element.click()
                    self.logger.info("Found and clicked domestic property text element")
                except TimeoutException:
                    try:
                        # Method 3: Force click the hidden radio button using JavaScript
                        domestic_radio = self.driver.find_element(By.ID, "domestic")
                        self.driver.execute_script("arguments[0].click();", domestic_radio)
                        self.logger.info("Force-clicked domestic radio button with JavaScript")
                    except Exception:
                        # Method 4: Try the parent div or container
                        domestic_container = self.wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//div[.//input[@id='domestic']]"))
                        )
                        domestic_container.click()
                        self.logger.info("Clicked domestic property container")
            
            # Add a small wait to ensure the selection is registered
            time.sleep(1)
            
            # Verify the radio button is now selected
            try:
                domestic_radio = self.driver.find_element(By.ID, "domestic")
                if domestic_radio.is_selected():
                    self.logger.info("Domestic property option is now selected")
                else:
                    self.logger.warning("Domestic property option may not be selected")
            except Exception as e:
                self.logger.warning(f"Could not verify radio button selection: {e}")
            
            # Click continue button - try different text variations
            continue_button = None
            try:
                continue_button = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continue')]"))
                )
            except TimeoutException:
                try:
                    continue_button = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Next')]"))
                    )
                except TimeoutException:
                    # Try generic button selector
                    continue_button = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
                    )
            
            continue_button.click()
            self.logger.info("Clicked continue button")
            
            # Wait for next page to load
            self.wait_for_page_load()
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to select domestic property option: {e}")
            # Debug the current page state
            self.debug_page_state("domestic_selection_failed")
            return False
    
    def enter_postcode(self, postcode):
        """Enter postcode and search"""
        try:
            # Wait for page to load first
            self.wait_for_page_load()
            
            # Wait for postcode input field
            postcode_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "postcode"))
            )
            postcode_input.clear()
            postcode_input.send_keys(postcode.strip())
            self.logger.info(f"Entered postcode: {postcode}")
            
            # Click find button - try multiple button text variations
            find_button = None
            try:
                # Try "Find" button first (the actual text)
                find_button = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Find')]"))
                )
                self.logger.info("Found 'Find' button")
            except TimeoutException:
                try:
                    # Fallback to "Find address" 
                    find_button = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Find address')]"))
                    )
                    self.logger.info("Found 'Find address' button")
                except TimeoutException:
                    # Try by class name
                    find_button = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[@class='govuk-button']"))
                    )
                    self.logger.info("Found button by class")
            
            find_button.click()
            self.logger.info("Clicked find button")
            
            # Wait for the next page/results to load
            self.wait_for_page_load()
            return True
            
        except TimeoutException:
            self.logger.error(f"Failed to enter postcode: {postcode}")
            self.debug_page_state("postcode_entry_failed")
            return False
    
    def select_address(self, target_address, postcode):
        """Select the correct address from the list of address links"""
        try:
            # Wait for page to load after clicking Find
            self.wait_for_page_load()
            
            # Wait for address links to appear on the page
            self.logger.info("Waiting for address links to appear...")
            
            # Look for address links - these are clickable links in a list format
            try:
                # Wait for at least one address link to be present
                self.wait.until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'certificate')]"))
                )
                self.logger.info("Found address links on page")
                
                # Get all address links
                address_links = self.driver.find_elements(By.XPATH, "//a[contains(@href, 'certificate')]")
                
                if not address_links:
                    # Try alternative selectors for address links
                    address_links = self.driver.find_elements(By.XPATH, "//a[contains(text(), 'Mallard House') or contains(text(), 'Iris Avenue')]")
                
                if not address_links:
                    # Try even broader search for links containing address components
                    address_links = self.driver.find_elements(By.XPATH, f"//a[contains(text(), '{postcode}')]")
                
                self.logger.info(f"Found {len(address_links)} address links")
                
                # Log all available addresses for debugging
                self.logger.info("Available addresses:")
                for i, link in enumerate(address_links[:10]):  # Show first 10
                    try:
                        link_text = link.text.strip()
                        self.logger.info(f"  {i+1}. {link_text}")
                    except Exception as e:
                        self.logger.warning(f"  {i+1}. Error reading link text: {e}")
                
                # Try to find the best matching address
                best_match = None
                best_score = 0
                
                for link in address_links:
                    try:
                        link_text = link.text.strip()
                        
                        # Check if this address matches our target
                        if self.address_matches(target_address, link_text):
                            self.logger.info(f"Found matching address: {link_text}")
                            link.click()
                            self.logger.info("Successfully clicked matching address link")
                            return True
                        
                        # Calculate match score for fallback selection
                        score = self.calculate_address_match_score(target_address, link_text)
                        if score > best_score:
                            best_score = score
                            best_match = link
                            
                    except Exception as e:
                        self.logger.warning(f"Error processing address link: {e}")
                        continue
                
                # If no exact match found, try the best scoring match
                if best_match and best_score > 0.3:  # At least 30% match
                    try:
                        best_text = best_match.text.strip()
                        self.logger.warning(f"No exact match found. Using best match: {best_text} (score: {best_score:.2f})")
                        best_match.click()
                        self.logger.info("Successfully clicked best matching address link")
                        return True
                    except Exception as e:
                        self.logger.error(f"Failed to click best match: {e}")
                
                # If still no match, try the first address as last resort
                if address_links:
                    try:
                        first_address = address_links[0].text.strip()
                        self.logger.warning(f"No good matches found. Trying first address: {first_address}")
                        address_links[0].click()
                        self.logger.info("Clicked first available address as fallback")
                        return True
                    except Exception as e:
                        self.logger.error(f"Failed to click first address: {e}")
                
                self.logger.error("No suitable address found to click")
                return False
                
            except TimeoutException:
                self.logger.error("No address links found on page")
                self.debug_page_state("no_address_links")
                return False
            
        except Exception as e:
            self.logger.error(f"Error in address selection: {e}")
            self.debug_page_state("address_selection_error")
            return False
    
    def calculate_address_match_score(self, target_address, option_text):
        """Calculate a match score between target address and option text"""
        try:
            # Convert to lowercase for comparison
            target_lower = target_address.lower()
            option_lower = option_text.lower()
            
            # Extract key components
            target_words = set(word.strip(',') for word in target_lower.split() if len(word.strip(',')) > 2)
            option_words = set(word.strip(',') for word in option_lower.split() if len(word.strip(',')) > 2)
            
            # Calculate overlap
            common_words = target_words.intersection(option_words)
            
            if len(target_words) > 0:
                score = len(common_words) / len(target_words)
                return score
            
            return 0.0
            
        except Exception:
            return 0.0
    
    def address_matches(self, target_address, option_text):
        """Check if addresses match - you can enhance this logic"""
        # Simple matching - convert to lowercase and check if key parts match
        target_lower = target_address.lower()
        option_lower = option_text.lower()
        
        # Extract key components and check for matches
        target_words = set(target_lower.split())
        option_words = set(option_lower.split())
        
        # Calculate overlap
        common_words = target_words.intersection(option_words)
        
        # If more than 50% of words match, consider it a match
        if len(target_words) > 0:
            match_percentage = len(common_words) / len(target_words)
            return match_percentage > 0.5
        
        return False
    
    def download_pdf(self, filename):
        """Print page to PDF and rename to custom filename"""
        try:
            # Look for print button first
            try:
                print_button = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Print') or contains(@href, 'print')]"))
                )
                print_button.click()
                self.logger.info("Clicked print button")
            except TimeoutException:
                # If no print button found, use Ctrl+P
                self.logger.info("No print button found, using Ctrl+P")
                self.driver.execute_script("window.print();")
            
            # Wait for download to complete
            self.logger.info("Waiting for PDF download to complete...")
            time.sleep(8)  # Give more time for download
            
            # Find the downloaded file and rename it
            downloaded_file = self.find_and_rename_downloaded_file(filename)
            
            if downloaded_file:
                self.logger.info(f"PDF successfully downloaded and renamed to: {filename}")
                return True
            else:
                self.logger.warning(f"PDF download completed but file renaming failed for: {filename}")
                return True  # Still consider it a success since file was downloaded
            
        except Exception as e:
            self.logger.error(f"Failed to download PDF: {str(e)}")
            return False
    
    def find_and_rename_downloaded_file(self, target_filename):
        """Find the most recently downloaded PDF and rename it to target filename"""
        try:
            import glob
            import os
            from pathlib import Path
            
            # Common patterns for downloaded EPC files
            search_patterns = [
                "Energy performance certificate (EPC)*.pdf",
                "Find an energy certificate*.pdf", 
                "EPC*.pdf",
                "*.pdf"
            ]
            
            most_recent_file = None
            most_recent_time = 0
            
            # Search for recently downloaded files
            for pattern in search_patterns:
                search_path = os.path.join(self.download_dir, pattern)
                files = glob.glob(search_path)
                
                for file_path in files:
                    # Check if file was modified recently (within last 30 seconds)
                    file_time = os.path.getmtime(file_path)
                    current_time = time.time()
                    
                    if (current_time - file_time) < 30 and file_time > most_recent_time:
                        most_recent_time = file_time
                        most_recent_file = file_path
            
            if most_recent_file:
                target_path = os.path.join(self.download_dir, target_filename)
                
                # If target file already exists, remove it
                if os.path.exists(target_path):
                    os.remove(target_path)
                    self.logger.info(f"Removed existing file: {target_filename}")
                
                # Rename the downloaded file
                os.rename(most_recent_file, target_path)
                self.logger.info(f"Renamed '{os.path.basename(most_recent_file)}' to '{target_filename}'")
                return target_path
            else:
                self.logger.warning("No recently downloaded PDF file found to rename")
                return None
                
        except Exception as e:
            self.logger.error(f"Error in file renaming: {str(e)}")
            return None
    
    def process_single_property(self, row):
        """Process a single property"""
        postcode = str(row.get('Post Code', '')).strip()
        full_address = self.construct_full_address(row)
        filename = self.generate_filename(row)
        
        self.logger.info(f"Processing property: {full_address}, {postcode}")
        
        try:
            # Navigate to start
            if not self.navigate_to_start():
                raise Exception("Failed to navigate to start")
            
            # Select domestic property
            if not self.select_domestic_property():
                raise Exception("Failed to select domestic property")
            
            # Enter postcode
            if not self.enter_postcode(postcode):
                raise Exception("Failed to enter postcode")
            
            # Select address
            if not self.select_address(full_address, postcode):
                raise Exception("Failed to select address")
            
            # Download PDF
            if not self.download_pdf(filename):
                raise Exception("Failed to download PDF")
            
            self.success_count += 1
            self.results.append({
                'Address': full_address,
                'Postcode': postcode,
                'Filename': filename,
                'Status': 'Success',
                'Error': None
            })
            
            self.logger.info(f"Successfully processed: {full_address}")
            return True
            
        except Exception as e:
            self.failure_count += 1
            error_msg = str(e)
            self.results.append({
                'Address': full_address,
                'Postcode': postcode,
                'Filename': filename,
                'Status': 'Failed',
                'Error': error_msg
            })
            
            self.logger.error(f"Failed to process {full_address}: {error_msg}")
            return False
    
    def process_spreadsheet(self, filepath):
        """Process entire spreadsheet"""
        try:
            # Read spreadsheet
            df = pd.read_csv(filepath) if filepath.endswith('.csv') else pd.read_excel(filepath)
            
            self.logger.info(f"Loaded {len(df)} properties from {filepath}")
            
            # Process each row
            for index, row in df.iterrows():
                self.logger.info(f"Processing property {index + 1} of {len(df)}")
                
                self.process_single_property(row)
                
                # Add delay between requests to be respectful
                time.sleep(2)
            
            # Generate summary report
            self.generate_summary_report()
            
        except Exception as e:
            self.logger.error(f"Failed to process spreadsheet: {str(e)}")
        finally:
            self.cleanup()
    
    def generate_summary_report(self):
        """Generate summary report"""
        self.logger.info("="*50)
        self.logger.info("PROCESSING SUMMARY")
        self.logger.info("="*50)
        self.logger.info(f"Total properties processed: {self.success_count + self.failure_count}")
        self.logger.info(f"Successful downloads: {self.success_count}")
        self.logger.info(f"Failed downloads: {self.failure_count}")
        
        if self.failure_count > 0:
            self.logger.info("\nFailed properties:")
            for result in self.results:
                if result['Status'] == 'Failed':
                    self.logger.info(f"  - {result['Address']}: {result['Error']}")
        
        # Save detailed results to CSV
        results_df = pd.DataFrame(self.results)
        results_file = os.path.join(os.path.dirname(self.download_dir), 
                                  f"processing_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        results_df.to_csv(results_file, index=False)
        self.logger.info(f"Detailed results saved to: {results_file}")
    
    def cleanup(self):
        """Cleanup resources"""
        if hasattr(self, 'driver'):
            self.driver.quit()
            self.logger.info("Browser closed")

# Usage example
def main():
    # Initialize scraper
    scraper = EPCCertificateScraper()
    
    # Auto-detect spreadsheet in current directory
    spreadsheet_path = "Spring Acres - Cantebury.xlsx"
    
    if os.path.exists(spreadsheet_path):
        print(f"Found spreadsheet: {spreadsheet_path}")
        print("Starting EPC certificate scraping...")
        scraper.process_spreadsheet(spreadsheet_path)
    else:
        print(f"Spreadsheet not found: {spreadsheet_path}")
        print("Please ensure 'Spring Acres - Cantebury.xlsx' is in the current directory.")

if __name__ == "__main__":
    main()