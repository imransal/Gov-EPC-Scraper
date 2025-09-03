#!/usr/bin/env python3
"""
Enhanced test to find the clickable elements for radio buttons
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import time

def test_clickable_elements():
    """Test finding clickable elements for radio buttons"""
    print("Testing Clickable Elements for Radio Buttons...")
    print("=" * 60)
    
    # Setup Chrome
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    try:
        # Navigate to the page
        print("1. Getting to the property type page...")
        driver.get("https://www.gov.uk/find-energy-certificate")
        time.sleep(2)
        
        start_button = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Start now")))
        start_button.click()
        time.sleep(3)
        
        print(f"2. On page: {driver.current_url}")
        
        # Look for labels
        print("3. Looking for labels...")
        labels = driver.find_elements(By.TAG_NAME, "label")
        print(f"   Found {len(labels)} labels:")
        
        for i, label in enumerate(labels):
            try:
                label_text = label.text.strip()
                label_for = label.get_attribute('for')
                label_visible = label.is_displayed()
                label_clickable = label.is_enabled()
                print(f"   Label {i+1}: text='{label_text}', for='{label_for}', visible={label_visible}, clickable={label_clickable}")
            except Exception as e:
                print(f"   Label {i+1}: Error - {e}")
        
        # Look for divs that might contain the radio buttons
        print("4. Looking for divs containing radio buttons...")
        radio_divs = driver.find_elements(By.XPATH, "//div[.//input[@type='radio']]")
        print(f"   Found {len(radio_divs)} divs containing radio buttons:")
        
        for i, div in enumerate(radio_divs):
            try:
                div_text = div.text.strip()[:50]  # First 50 chars
                div_class = div.get_attribute('class')
                div_visible = div.is_displayed()
                print(f"   Div {i+1}: text='{div_text}...', class='{div_class}', visible={div_visible}")
            except Exception as e:
                print(f"   Div {i+1}: Error - {e}")
        
        # Try to find and click the domestic option
        print("5. Attempting to click domestic option...")
        
        # Method 1: Try label for domestic
        try:
            domestic_label = driver.find_element(By.XPATH, "//label[@for='domestic']")
            print(f"   Found label for domestic: text='{domestic_label.text}', visible={domestic_label.is_displayed()}")
            if domestic_label.is_displayed():
                domestic_label.click()
                print("   ✓ Successfully clicked domestic label!")
                
                # Check if radio is now selected
                domestic_radio = driver.find_element(By.ID, "domestic")
                print(f"   Radio button selected: {domestic_radio.is_selected()}")
            else:
                print("   Label not visible")
        except Exception as e:
            print(f"   Could not find/click label: {e}")
        
        # Method 2: Try JavaScript click
        if not driver.find_element(By.ID, "domestic").is_selected():
            print("6. Trying JavaScript click...")
            try:
                domestic_radio = driver.find_element(By.ID, "domestic")
                driver.execute_script("arguments[0].click();", domestic_radio)
                print("   ✓ JavaScript click executed!")
                print(f"   Radio button selected: {domestic_radio.is_selected()}")
            except Exception as e:
                print(f"   JavaScript click failed: {e}")
        
        print("\n7. Final status check...")
        domestic_radio = driver.find_element(By.ID, "domestic")
        print(f"   Domestic radio selected: {domestic_radio.is_selected()}")
        
        # Look for continue button
        print("8. Looking for continue button...")
        continue_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Continue')]")
        print(f"   Found {len(continue_buttons)} continue buttons")
        
        if continue_buttons:
            continue_button = continue_buttons[0]
            print(f"   Continue button visible: {continue_button.is_displayed()}")
            print(f"   Continue button enabled: {continue_button.is_enabled()}")
        
        print("\n9. Pausing for inspection...")
        time.sleep(10)
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("Closing browser...")
        driver.quit()

if __name__ == "__main__":
    test_clickable_elements()