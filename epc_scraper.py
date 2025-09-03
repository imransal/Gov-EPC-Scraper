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
import sys
from pathlib import Path
import xlsxwriter
import argparse
import signal
import atexit

class EPCCertificateScraper:
    def __init__(self, download_dir="C:\\Users\\IS19\\Documents\\EPC_Scraper\\Processed"):
        self.download_dir = download_dir
        self.original_spreadsheet_data = None  # Store original data for report
        self.spreadsheet_filepath = None  # Store original file path
        self.setup_logging()
        self.setup_driver()
        self.success_count = 0
        self.failure_count = 0
        self.results = []
        self.cleanup_registered = False
        
        # Register cleanup function to run on exit
        if not self.cleanup_registered:
            atexit.register(self.emergency_cleanup)
            self.cleanup_registered = True
        
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
                exact_matches = []
                
                for link in address_links:
                    try:
                        link_text = link.text.strip()
                        
                        # Skip empty or irrelevant links
                        if not link_text or "get a new energy certificate" in link_text.lower():
                            continue
                        
                        # Calculate enhanced match score
                        score = self.calculate_enhanced_address_match_score(target_address, link_text)
                        
                        # Check for high confidence matches
                        if score >= 0.8:
                            exact_matches.append((link, link_text, score))
                        
                        if score > best_score:
                            best_score = score
                            best_match = (link, link_text, score)
                            
                    except Exception as e:
                        self.logger.warning(f"Error processing address link: {e}")
                        continue
                
                # Prefer high-confidence matches
                selected_match = None
                if exact_matches:
                    # Sort by score and take the highest
                    exact_matches.sort(key=lambda x: x[2], reverse=True)
                    selected_match = exact_matches[0]
                    self.logger.info(f"Found high-confidence match: {selected_match[1]} (score: {selected_match[2]:.3f})")
                elif best_match and best_score > 0.3:
                    selected_match = best_match
                    self.logger.info(f"Found best match: {selected_match[1]} (score: {selected_match[2]:.3f})")
                else:
                    self.logger.warning(f"No good match found. Best score: {best_score:.3f}")
                
                # Click the selected match
                if selected_match:
                    try:
                        selected_match[0].click()
                        self.logger.info("Successfully clicked matching address link")
                        return True
                    except Exception as e:
                        self.logger.error(f"Failed to click selected match: {e}")
                
                # Fallback to first address if no good matches
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
    
    def normalize_address_for_matching(self, address):
        """Normalize address text for better matching"""
        if not address:
            return ""
        
        # Convert to lowercase
        normalized = address.lower()
        
        # Remove common prefixes and suffixes
        normalized = re.sub(r'\b(flat|apartment|apt|unit)\s*', '', normalized)
        
        # Remove punctuation and extra spaces
        normalized = re.sub(r'[,\.\-\(\)]', ' ', normalized)
        normalized = re.sub(r'\s+', ' ', normalized).strip()
        
        # Remove common words that don't help with matching
        stop_words = ['the', 'and', 'of', 'in', 'at', 'to', 'for', 'with', 'by']
        words = normalized.split()
        words = [word for word in words if word not in stop_words]
        
        return ' '.join(words)

    def extract_property_number(self, address):
        """Extract property number from address"""
        # Look for patterns like "Flat 9", "9 HOUSE", "Unit 4", etc.
        patterns = [
            r'\b(?:flat|apartment|apt|unit)\s*(\d+)\b',  # "Flat 9"
            r'\b(\d+)\s+(?:flat|apartment|apt|unit)\b',  # "9 Flat"
            r'^(\d+)\s+\w+',  # "9 HOUSE" at start
            r'\b(\d+)\s*[,\s]',  # Any number followed by comma or space
        ]
        
        for pattern in patterns:
            match = re.search(pattern, address.lower())
            if match:
                return match.group(1)
        return None

    def extract_building_name(self, address):
        """Extract building name from address"""
        # Remove property numbers and common prefixes
        cleaned = re.sub(r'^\d+\s*', '', address)  # Remove leading numbers
        cleaned = re.sub(r'\b(?:flat|apartment|apt|unit)\s*\d+[,\s]*', '', cleaned, flags=re.IGNORECASE)
        
        # Get the first significant part (usually building name)
        parts = [part.strip() for part in cleaned.split(',') if part.strip()]
        if parts:
            # Return first part, but clean it up
            building = parts[0].strip()
            building = re.sub(r'[,\.]', '', building)
            return building.lower()
        return ""

    def calculate_enhanced_address_match_score(self, target_address, option_text):
        """Enhanced address matching with better logic"""
        try:
            target_lower = target_address.lower()
            option_lower = option_text.lower()
            
            # Extract components
            target_number = self.extract_property_number(target_address)
            option_number = self.extract_property_number(option_text)
            
            target_building = self.extract_building_name(target_address)
            option_building = self.extract_building_name(option_text)
            
            # Normalize both addresses
            target_normalized = self.normalize_address_for_matching(target_address)
            option_normalized = self.normalize_address_for_matching(option_text)
            
            score = 0.0
            
            # Property number matching (high weight)
            if target_number and option_number:
                if target_number == option_number:
                    score += 0.4  # Strong match for same number
                else:
                    score -= 0.3  # Penalty for different numbers
            elif target_number or option_number:
                score -= 0.1  # Small penalty if only one has a number
            
            # Building name matching (high weight)
            if target_building and option_building:
                if target_building in option_building or option_building in target_building:
                    score += 0.4
                else:
                    # Check for partial matches
                    target_words = set(target_building.split())
                    option_words = set(option_building.split())
                    common_building_words = target_words.intersection(option_words)
                    if common_building_words:
                        score += 0.2 * len(common_building_words) / max(len(target_words), len(option_words))
            
            # Overall word matching (medium weight)
            target_words = set(target_normalized.split())
            option_words = set(option_normalized.split())
            
            if len(target_words) > 0:
                common_words = target_words.intersection(option_words)
                word_score = len(common_words) / len(target_words)
                score += 0.2 * word_score
            
            # Substring matching for exact phrases (low weight)
            if target_building and target_building in option_lower:
                score += 0.1
            
            # Debug logging
            self.logger.debug(f"Matching '{target_address}' vs '{option_text}':")
            self.logger.debug(f"  Target number: {target_number}, Option number: {option_number}")
            self.logger.debug(f"  Target building: '{target_building}', Option building: '{option_building}'")
            self.logger.debug(f"  Final score: {score}")
            
            return max(0.0, min(1.0, score))  # Clamp between 0 and 1
            
        except Exception as e:
            self.logger.warning(f"Error in enhanced address matching: {e}")
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
    
    def generate_excel_report(self, intermediate=False, interrupted=False, error=False):
        """Generate comprehensive Excel report with results and original data."""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # Add status suffix to filename
            status_suffix = ""
            if intermediate:
                status_suffix = "_INTERMEDIATE"
            elif interrupted:
                status_suffix = "_INTERRUPTED"
            elif error:
                status_suffix = "_ERROR"
            
            report_filename = f"EPC_Processing_Report_{timestamp}{status_suffix}.xlsx"
            report_path = os.path.join(self.download_dir, report_filename)
            
            # Create workbook with NaN handling
            workbook = xlsxwriter.Workbook(report_path, {'nan_inf_to_errors': True})
            
            # Create formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#366092',
                'font_color': 'white',
                'border': 1
            })
            
            success_format = workbook.add_format({
                'bg_color': '#C6EFCE',
                'border': 1
            })
            
            failure_format = workbook.add_format({
                'bg_color': '#FFC7CE',
                'border': 1
            })
            
            # Create Results worksheet
            results_ws = workbook.add_worksheet('Processing Results')
            
            # Write results data
            if self.results:
                results_df = pd.DataFrame(self.results)
                
                # Write headers
                for col, header in enumerate(results_df.columns):
                    results_ws.write(0, col, header, header_format)
                
                # Write data with conditional formatting
                for row, record in enumerate(results_df.to_dict('records'), start=1):
                    for col, (key, value) in enumerate(record.items()):
                        # Handle NaN and None values
                        if pd.isna(value) or value is None:
                            value = ""
                        elif isinstance(value, float) and (value != value):  # Check for NaN
                            value = ""
                        
                        if key == 'Status':
                            if value == 'Success':
                                results_ws.write(row, col, value, success_format)
                            else:
                                results_ws.write(row, col, value, failure_format)
                        else:
                            results_ws.write(row, col, value)
                
                # Auto-adjust column widths
                for col in range(len(results_df.columns)):
                    results_ws.set_column(col, col, 20)
            
            # Create Original Data worksheet if available
            if self.original_spreadsheet_data is not None:
                original_ws = workbook.add_worksheet('Original Spreadsheet')
                
                # Write original data headers
                for col, header in enumerate(self.original_spreadsheet_data.columns):
                    original_ws.write(0, col, header, header_format)
                
                # Write original data
                for row, record in enumerate(self.original_spreadsheet_data.to_dict('records'), start=1):
                    for col, (key, value) in enumerate(record.items()):
                        # Handle NaN and None values
                        if pd.isna(value) or value is None:
                            value = ""
                        elif isinstance(value, float) and (value != value):  # Check for NaN
                            value = ""
                        
                        original_ws.write(row, col, value)
                
                # Auto-adjust column widths
                for col in range(len(self.original_spreadsheet_data.columns)):
                    original_ws.set_column(col, col, 15)
            
            # Create Summary worksheet
            summary_ws = workbook.add_worksheet('Summary')
            
            # Write summary information
            processing_status = "Completed"
            if intermediate:
                processing_status = "In Progress"
            elif interrupted:
                processing_status = "Interrupted by User"
            elif error:
                processing_status = "Stopped due to Error"
            
            summary_data = [
                ['Processing Summary', ''],
                ['Status', processing_status],
                ['Total Properties', self.success_count + self.failure_count],
                ['Successful Downloads', self.success_count],
                ['Failed Downloads', self.failure_count],
                ['Success Rate (%)', round((self.success_count / (self.success_count + self.failure_count)) * 100, 2) if (self.success_count + self.failure_count) > 0 else 0],
                ['Processing Date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                ['Original File', os.path.basename(self.spreadsheet_filepath) if self.spreadsheet_filepath else 'Unknown'],
                ['Download Directory', self.download_dir]
            ]
            
            for row, (label, value) in enumerate(summary_data):
                summary_ws.write(row, 0, label, header_format)
                summary_ws.write(row, 1, value)
            
            summary_ws.set_column(0, 0, 25)
            summary_ws.set_column(1, 1, 30)
            
            workbook.close()
            
            if not intermediate:
                print(f"📊 Excel report generated: {report_path}")
            self.logger.info(f"Excel report generated: {report_path}")
            return report_path
            
        except Exception as e:
            self.logger.error(f"Error generating Excel report: {e}")
            print(f"Error generating Excel report: {e}")
            return None

    def process_spreadsheet(self, file_path):
        """Read and process the Excel spreadsheet containing addresses."""
        try:
            print(f"Reading addresses from: {file_path}")
            self.logger.info(f"Processing spreadsheet: {file_path}")
            
            # Store original file path
            self.spreadsheet_filepath = file_path
            
            # Read the spreadsheet
            df = pd.read_excel(file_path)
            
            # Store original data for report
            self.original_spreadsheet_data = df.copy()
            
            # Check for different possible column formats
            postcode_col = None
            address_cols = []
            
            # Check for postcode column variations
            for col in df.columns:
                if col.lower() in ['postcode', 'post code', 'postal code']:
                    postcode_col = col
                    break
            
            # Check for address columns
            if 'Address' in df.columns:
                # Simple format: single Address column
                address_cols = ['Address']
            else:
                # Complex format: multiple address line columns
                for i in range(1, 6):  # Address Line 1-5
                    col_name = f'Address Line {i}'
                    if col_name in df.columns:
                        address_cols.append(col_name)
            
            if not postcode_col:
                print("Error: Spreadsheet must contain a postcode column ('Postcode', 'Post Code', etc.)")
                self.logger.error("No postcode column found in spreadsheet")
                return False
                
            if not address_cols:
                print("Error: Spreadsheet must contain address information ('Address' or 'Address Line X' columns)")
                self.logger.error("No address columns found in spreadsheet")
                return False
            
            total_addresses = len(df)
            print(f"Found {total_addresses} addresses to process")
            self.logger.info(f"Found {total_addresses} addresses to process")
            print(f"Using postcode column: '{postcode_col}'")
            print(f"Using address columns: {address_cols}")
            
            # Process each address
            for index, row in df.iterrows():
                # Construct full address from multiple columns
                address_parts = []
                for col in address_cols:
                    value = str(row[col]).strip()
                    if value and value.lower() != 'nan':
                        address_parts.append(value)
                
                # Add town if available and not already included
                if 'Town' in df.columns:
                    town = str(row['Town']).strip()
                    if town and town.lower() != 'nan':
                        address_parts.append(town)
                
                full_address = ', '.join(address_parts)
                postcode = str(row[postcode_col]).strip()
                
                print(f"\nProcessing address {index + 1}/{total_addresses}")
                print(f"Address: {full_address}")
                print(f"Postcode: {postcode}")
                
                result = self.download_epc_certificate(full_address, postcode, row)
                
                # Store result with original row data
                result_entry = {
                    'Original_Index': index,
                    'Address': full_address,
                    'Postcode': postcode,
                    'Status': 'Success' if result else 'Failed',
                    'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                
                # Add any additional columns from original data
                for col in df.columns:
                    if col not in address_cols and col != postcode_col:
                        result_entry[f'Original_{col}'] = row[col]
                
                self.results.append(result_entry)
                
                if result:
                    self.success_count += 1
                else:
                    self.failure_count += 1
                
                # Generate intermediate report every 10 properties or if we have failures
                if (index + 1) % 10 == 0 or not result:
                    self.generate_excel_report(intermediate=True)
                
                # Brief pause between requests
                time.sleep(2)
            
            self.generate_excel_report()
            return True
            
        except KeyboardInterrupt:
            print("\n⚠️  Processing interrupted by user")
            self.logger.info("Processing interrupted by user")
            self.generate_excel_report(interrupted=True)
            return False
        except Exception as e:
            print(f"Error processing spreadsheet: {e}")
            self.logger.error(f"Error processing spreadsheet: {e}")
            self.generate_excel_report(error=True)
            return False
    
    def download_epc_certificate(self, address, postcode, row_data=None):
        """Main method to download EPC certificate for a single address."""
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
            if not self.select_address(address, postcode):
                raise Exception("Failed to select address")
            
            # Generate filename using the original format
            filename = self.generate_epc_filename(row_data) if row_data is not None else self.generate_simple_filename(address, postcode)
            
            # Download PDF
            if not self.download_pdf(filename):
                raise Exception("Failed to download PDF")
            
            self.logger.info(f"Successfully processed: {address}, {postcode}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to process {address}, {postcode}: {e}")
            return False

    def generate_epc_filename(self, row_data):
        """Generate filename in the original EPC format: EPC - [Scheme] - [Plot] - [Tenure] - [UPRN].pdf"""
        try:
            # Extract components from row data
            scheme = str(row_data.get('Scheme Abbreviation', 'UNK')).strip()
            plot = str(row_data.get('Development Plot Number', 'UNK')).strip()
            tenure = str(row_data.get('Tenure', 'UNK')).strip()
            uprn = str(row_data.get('UPRN', 'UNK')).strip()
            
            # Clean components (remove 'nan' and empty values)
            scheme = scheme if scheme and scheme.lower() != 'nan' else 'UNK'
            plot = plot if plot and plot.lower() != 'nan' else 'UNK'
            tenure = tenure if tenure and tenure.lower() != 'nan' else 'UNK'
            uprn = uprn if uprn and uprn.lower() != 'nan' else 'UNK'
            
            filename = f"EPC - {scheme} - {plot} - {tenure} - {uprn}.pdf"
            self.logger.info(f"Generated EPC filename: {filename}")
            return filename
            
        except Exception as e:
            self.logger.warning(f"Error generating EPC filename, using fallback: {e}")
            return self.generate_simple_filename("Property", "Unknown")
    
    def generate_simple_filename(self, address, postcode):
        """Generate simple filename as fallback"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        clean_address = address.replace(' ', '_').replace(',', '').replace('/', '_')[:50]
        clean_postcode = postcode.replace(' ', '')
        return f"{clean_address}_{clean_postcode}_{timestamp}.pdf"

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
    
    def emergency_cleanup(self):
        """Emergency cleanup function that runs on unexpected exit"""
        try:
            if hasattr(self, 'results') and self.results:
                print("\n🚨 Emergency cleanup: Generating Excel report...")
                self.logger.info("Emergency cleanup: Generating Excel report")
                self.generate_excel_report(interrupted=True)
                print("✅ Excel report saved before exit")
        except Exception as e:
            print(f"❌ Error during emergency cleanup: {e}")
        
        try:
            if hasattr(self, 'driver'):
                self.driver.quit()
        except:
            pass

    def cleanup(self):
        """Cleanup resources and ensure final report is generated"""
        try:
            # Generate final report if we have any results
            if hasattr(self, 'results') and self.results:
                print("🔄 Generating final Excel report...")
                self.generate_excel_report()
        except Exception as e:
            self.logger.error(f"Error generating final report during cleanup: {e}")
        
        # Close browser
        if hasattr(self, 'driver'):
            try:
                self.driver.quit()
                self.logger.info("Browser closed")
            except:
                pass

# Usage example
def main():
    """Main function with command line argument support."""
    parser = argparse.ArgumentParser(description='EPC Certificate Scraper')
    parser.add_argument('--file', '-f', type=str, help='Path to spreadsheet file')
    parser.add_argument('--download-dir', '-d', type=str, 
                       default='C:\\Users\\IS19\\Documents\\EPC_Scraper\\Processed',
                       help='Download directory for PDFs')
    
    args = parser.parse_args()
    
    # Initialize scraper
    scraper = EPCCertificateScraper(download_dir=args.download_dir)
    
    def signal_handler(signum, frame):
        """Handle Ctrl+C gracefully"""
        print("\n🛑 Interrupt signal received. Saving progress...")
        scraper.logger.info("Interrupt signal received")
        if hasattr(scraper, 'results') and scraper.results:
            scraper.generate_excel_report(interrupted=True)
            print("✅ Progress saved to Excel report")
        scraper.cleanup()
        sys.exit(0)
    
    # Register signal handler for Ctrl+C
    signal.signal(signal.SIGINT, signal_handler)
    
    try:
        # Determine spreadsheet file
        spreadsheet_path = None
        
        if args.file:
            # Use specified file
            if os.path.exists(args.file):
                spreadsheet_path = args.file
            else:
                print(f"Error: Specified file '{args.file}' not found")
                return
        else:
            # Auto-detect spreadsheet files in order of preference
            current_dir = os.getcwd()
            potential_files = [
                "spreadsheet.xlsx",  # First priority
                "Spring Acres - Cantebury.xlsx"  # Fallback to existing file
            ]
            
            for filename in potential_files:
                full_path = os.path.join(current_dir, filename)
                if os.path.exists(full_path):
                    spreadsheet_path = full_path
                    print(f"Found spreadsheet: {filename}")
                    break
            
            # If no priority files found, look for any Excel files
            if not spreadsheet_path:
                excel_files = [f for f in os.listdir(current_dir) 
                             if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
                
                if excel_files:
                    spreadsheet_path = os.path.join(current_dir, excel_files[0])
                    print(f"Found Excel file: {excel_files[0]}")
                    
                    if len(excel_files) > 1:
                        print(f"Note: Found {len(excel_files)} Excel files. Using '{excel_files[0]}'")
                        print("Use --file argument to specify a different file")
        
        if spreadsheet_path:
            print(f"Processing spreadsheet: {spreadsheet_path}")
            print("Starting EPC certificate scraping...")
            success = scraper.process_spreadsheet(spreadsheet_path)
            
            if success:
                print("\nProcessing completed!")
                print(f"Successful downloads: {scraper.success_count}")
                print(f"Failed downloads: {scraper.failure_count}")
            else:
                print("Processing failed!")
        else:
            print("Error: No suitable spreadsheet file found!")
            print("Please ensure you have one of the following files:")
            print("  - spreadsheet.xlsx (preferred)")
            print("  - Any .xlsx or .xls file")
            print("Or use --file argument to specify a file path")
    
    except KeyboardInterrupt:
        print("\n🛑 Processing interrupted by user")
        if hasattr(scraper, 'results') and scraper.results:
            print("💾 Saving progress to Excel report...")
            scraper.generate_excel_report(interrupted=True)
            print("✅ Progress saved!")
        scraper.cleanup()
    except Exception as e:
        print(f"❌ An error occurred: {e}")
        scraper.logger.error(f"Main execution error: {e}")
        if hasattr(scraper, 'results') and scraper.results:
            print("💾 Saving progress to Excel report...")
            scraper.generate_excel_report(error=True)
    finally:
        try:
            scraper.cleanup()
        except:
            pass

if __name__ == "__main__":
    main()