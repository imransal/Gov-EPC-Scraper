#!/usr/bin/env python3
"""
Test script to verify the address link selection works
"""

import pandas as pd
from epc_scraper import EPCCertificateScraper
from selenium.webdriver.common.by import By
import time

def test_address_links():
    """Test clicking on address links from the list"""
    print("Testing Address Link Selection...")
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
        print("Navigating to address list page...")
        scraper.navigate_to_start()
        scraper.select_domestic_property()
        scraper.enter_postcode(postcode)
        
        # Now test the address selection
        print("Testing address selection...")
        if scraper.select_address(full_address, postcode):
            print("✓ Successfully selected and clicked address!")
            
            # Check what page we're on now
            time.sleep(3)
            current_url = scraper.driver.current_url
            page_title = scraper.driver.title
            print(f"Current URL: {current_url}")
            print(f"Page title: {page_title}")
            
            # Look for EPC certificate or next steps
            if "certificate" in current_url.lower():
                print("✓ Successfully navigated to certificate page!")
            else:
                print("? Landed on different page, checking content...")
                
            print("Pausing 10 seconds to inspect result...")
            time.sleep(10)
            
        else:
            print("✗ Failed to select address")
        
    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Closing browser...")
        scraper.cleanup()

if __name__ == "__main__":
    test_address_links()