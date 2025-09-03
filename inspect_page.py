#!/usr/bin/env python3
"""
Simple test to examine the radio button page
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import time

def test_radio_button_page():
    """Test accessing the radio button page"""
    print("Testing Radio Button Page Access...")
    print("=" * 50)
    
    # Setup Chrome
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    try:
        # Navigate to start
        print("1. Navigating to start page...")
        driver.get("https://www.gov.uk/find-energy-certificate")
        time.sleep(2)
        
        # Click start now
        print("2. Clicking 'Start now'...")
        start_button = wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Start now"))
        )
        start_button.click()
        time.sleep(3)
        
        # Check what page we're on
        print(f"3. Current URL: {driver.current_url}")
        print(f"4. Page title: {driver.title}")
        
        # Look for radio buttons
        print("5. Looking for radio buttons...")
        radio_buttons = driver.find_elements(By.XPATH, "//input[@type='radio']")
        print(f"   Found {len(radio_buttons)} radio buttons:")
        
        for i, radio in enumerate(radio_buttons):
            try:
                radio_id = radio.get_attribute('id')
                radio_name = radio.get_attribute('name') 
                radio_value = radio.get_attribute('value')
                radio_visible = radio.is_displayed()
                radio_enabled = radio.is_enabled()
                print(f"   Radio {i+1}: id='{radio_id}', name='{radio_name}', value='{radio_value}', visible={radio_visible}, enabled={radio_enabled}")
            except Exception as e:
                print(f"   Radio {i+1}: Error - {e}")
        
        # Look for the specific domestic radio button
        print("6. Looking for domestic radio button specifically...")
        try:
            domestic_radio = driver.find_element(By.ID, "domestic")
            print(f"   Found by ID: visible={domestic_radio.is_displayed()}, enabled={domestic_radio.is_enabled()}")
        except:
            print("   Not found by ID")
            
        try:
            domestic_radio = driver.find_element(By.XPATH, "//input[@value='domestic']")
            print(f"   Found by value: visible={domestic_radio.is_displayed()}, enabled={domestic_radio.is_enabled()}")
        except:
            print("   Not found by value")
        
        # Look for buttons
        print("7. Looking for buttons...")
        buttons = driver.find_elements(By.TAG_NAME, "button")
        print(f"   Found {len(buttons)} buttons:")
        for i, button in enumerate(buttons):
            try:
                button_text = button.text.strip()
                button_type = button.get_attribute('type')
                button_visible = button.is_displayed()
                print(f"   Button {i+1}: text='{button_text}', type='{button_type}', visible={button_visible}")
            except Exception as e:
                print(f"   Button {i+1}: Error - {e}")
        
        print("\n8. Pausing for 30 seconds so you can inspect the page...")
        print("   Check the browser window to see what's actually displayed.")
        time.sleep(30)
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("Closing browser...")
        driver.quit()

if __name__ == "__main__":
    test_radio_button_page()