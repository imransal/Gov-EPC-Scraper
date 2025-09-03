#!/usr/bin/env python3
"""
Test script to verify postcode entry and find button clicking
"""

import pandas as pd
from epc_scraper import EPCCertificateScraper
from selenium.webdriver.common.by import By
import time

def test_postcode_entry():
    """Test the postcode entry and find button steps"""
    print("Testing Postcode Entry and Find Button...")
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
        
        # Navigate through the steps
        print("Step 1: Navigating to start...")
        if not scraper.navigate_to_start():
            print("✗ Failed at start navigation")
            return False
        
        print("Step 2: Selecting domestic property...")
        if not scraper.select_domestic_property():
            print("✗ Failed at domestic selection")
            return False
        
        print("Step 3: Entering postcode and clicking Find...")
        if not scraper.enter_postcode(postcode):
            print("✗ Failed at postcode entry")
            return False
        
        print("✓ Successfully entered postcode and clicked Find!")
        
        # Check what page we're on now
        current_url = scraper.driver.current_url
        page_title = scraper.driver.title
        print(f"Current URL: {current_url}")
        print(f"Page title: {page_title}")
        
        # Look for address dropdown or results
        print("Looking for address selection elements...")
        try:
            address_select = scraper.driver.find_element(By.ID, "address")
            print("✓ Found address dropdown!")
            
            # Get options count
            options = address_select.find_elements(By.TAG_NAME, "option")
            print(f"Found {len(options)} address options")
            
        except:
            print("No address dropdown found - checking for other elements...")
            
            # Look for any select elements
            selects = scraper.driver.find_elements(By.TAG_NAME, "select")
            print(f"Found {len(selects)} select elements")
            
            # Look for error messages
            errors = scraper.driver.find_elements(By.XPATH, "//*[contains(@class, 'error') or contains(@class, 'govuk-error')]")
            if errors:
                print("Found error messages:")
                for error in errors:
                    if error.is_displayed():
                        print(f"  - {error.text}")
        
        print("\nPausing for 15 seconds to inspect the page...")
        time.sleep(15)
        
        return True
        
    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        return False
    finally:
        print("Closing browser...")
        scraper.cleanup()

if __name__ == "__main__":
    test_postcode_entry()