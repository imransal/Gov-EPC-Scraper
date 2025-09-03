#!/usr/bin/env python3
"""
Test script to debug the address selection step specifically
"""

import pandas as pd
from epc_scraper import EPCCertificateScraper
from selenium.webdriver.common.by import By
import time

def test_address_selection():
    """Test the address selection step in detail"""
    print("Testing Address Selection Step...")
    print("=" * 50)
    
    # Initialize scraper
    scraper = EPCCertificateScraper()
    
    try:
        # Load sample property
        df = pd.read_excel("Spring Acres - Cantebury.xlsx")
        sample_row = df.iloc[0]  # First property
        
        postcode = str(sample_row.get('Post Code', '')).strip()
        full_address = scraper.construct_full_address(sample_row)
        
        print(f"Testing with:")
        print(f"  Address: {full_address}")
        print(f"  Postcode: {postcode}")
        print()
        
        # Navigate through the steps quickly
        print("Getting to postcode results page...")
        
        # Navigate to start
        scraper.navigate_to_start()
        scraper.select_domestic_property()
        scraper.enter_postcode(postcode)
        
        print("Now analyzing what appears after clicking Find...")
        time.sleep(5)  # Wait for page to load
        
        # Check current page state
        current_url = scraper.driver.current_url
        page_title = scraper.driver.title
        print(f"Current URL: {current_url}")
        print(f"Page title: {page_title}")
        
        # Look for different types of address selection elements
        print("\n1. Looking for SELECT dropdown...")
        selects = scraper.driver.find_elements(By.TAG_NAME, "select")
        print(f"   Found {len(selects)} select elements")
        
        for i, select in enumerate(selects):
            try:
                select_id = select.get_attribute('id')
                select_name = select.get_attribute('name')
                options_count = len(select.find_elements(By.TAG_NAME, "option"))
                print(f"   Select {i+1}: id='{select_id}', name='{select_name}', options={options_count}")
                
                if options_count > 0:
                    options = select.find_elements(By.TAG_NAME, "option")
                    print(f"     Options:")
                    for j, option in enumerate(options[:5]):  # Show first 5
                        print(f"       {j+1}. {option.text.strip()}")
                    if len(options) > 5:
                        print(f"       ... and {len(options)-5} more")
                        
            except Exception as e:
                print(f"   Select {i+1}: Error reading - {e}")
        
        print("\n2. Looking for radio buttons...")
        radios = scraper.driver.find_elements(By.XPATH, "//input[@type='radio']")
        print(f"   Found {len(radios)} radio buttons")
        
        for i, radio in enumerate(radios[:10]):  # Show first 10
            try:
                radio_name = radio.get_attribute('name')
                radio_value = radio.get_attribute('value')
                radio_id = radio.get_attribute('id')
                print(f"   Radio {i+1}: name='{radio_name}', value='{radio_value}', id='{radio_id}'")
            except Exception as e:
                print(f"   Radio {i+1}: Error - {e}")
        
        print("\n3. Looking for address links...")
        address_links = scraper.driver.find_elements(By.XPATH, "//a[contains(@href, 'address') or contains(text(), 'address')]")
        print(f"   Found {len(address_links)} address-related links")
        
        for i, link in enumerate(address_links[:5]):
            try:
                link_text = link.text.strip()[:50]
                link_href = link.get_attribute('href')
                print(f"   Link {i+1}: text='{link_text}...', href='{link_href}'")
            except Exception as e:
                print(f"   Link {i+1}: Error - {e}")
        
        print("\n4. Looking for any divs containing address info...")
        address_divs = scraper.driver.find_elements(By.XPATH, "//div[contains(text(), 'address') or contains(text(), 'Address')]")
        print(f"   Found {len(address_divs)} divs with address text")
        
        for i, div in enumerate(address_divs[:3]):
            try:
                div_text = div.text.strip()[:100]
                print(f"   Div {i+1}: {div_text}...")
            except Exception as e:
                print(f"   Div {i+1}: Error - {e}")
        
        print("\n5. Looking for error messages...")
        errors = scraper.driver.find_elements(By.XPATH, "//*[contains(@class, 'error') or contains(@class, 'govuk-error') or contains(text(), 'error') or contains(text(), 'Error')]")
        if errors:
            print(f"   Found {len(errors)} potential error elements:")
            for i, error in enumerate(errors):
                if error.is_displayed():
                    print(f"   Error {i+1}: {error.text.strip()}")
        else:
            print("   No error messages found")
        
        print("\n6. Looking for buttons...")
        buttons = scraper.driver.find_elements(By.TAG_NAME, "button")
        print(f"   Found {len(buttons)} buttons")
        for i, button in enumerate(buttons):
            try:
                button_text = button.text.strip()
                button_visible = button.is_displayed()
                if button_text and button_visible:
                    print(f"   Button {i+1}: '{button_text}' (visible)")
            except Exception as e:
                print(f"   Button {i+1}: Error - {e}")
        
        print(f"\n7. Waiting 30 seconds for you to inspect the page...")
        print("   Check the browser to see what's actually displayed.")
        time.sleep(30)
        
    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Closing browser...")
        scraper.cleanup()

if __name__ == "__main__":
    test_address_selection()