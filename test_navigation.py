#!/usr/bin/env python3
"""
Quick test script to debug the domestic property selection issue
"""

import pandas as pd
from epc_scraper import EPCCertificateScraper
import time

def test_navigation_steps():
    """Test just the navigation steps that are failing"""
    print("Testing EPC Scraper Navigation Steps...")
    print("=" * 50)
    
    # Initialize scraper
    scraper = EPCCertificateScraper()
    
    try:
        # Load one sample property
        df = pd.read_excel("Spring Acres - Cantebury.xlsx")
        sample_row = df.iloc[0]  # First property
        
        postcode = str(sample_row.get('Post Code', '')).strip()
        full_address = scraper.construct_full_address(sample_row)
        
        print(f"Testing with:")
        print(f"  Address: {full_address}")
        print(f"  Postcode: {postcode}")
        print()
        
        # Step 1: Navigate to start
        print("Step 1: Navigating to start page...")
        if scraper.navigate_to_start():
            print("✓ Successfully navigated to start")
        else:
            print("✗ Failed to navigate to start")
            return False
        
        time.sleep(2)  # Give user time to see what's happening
        
        # Step 2: Select domestic property
        print("Step 2: Selecting domestic property...")
        if scraper.select_domestic_property():
            print("✓ Successfully selected domestic property")
        else:
            print("✗ Failed to select domestic property")
            return False
        
        time.sleep(2)
        
        # Step 3: Try entering postcode
        print("Step 3: Entering postcode...")
        if scraper.enter_postcode(postcode):
            print("✓ Successfully entered postcode")
        else:
            print("✗ Failed to enter postcode")
            return False
        
        print("\n✅ All navigation steps completed successfully!")
        print("The browser window will stay open for 10 seconds so you can see the result...")
        time.sleep(10)
        
        return True
        
    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        return False
    finally:
        # Keep browser open a bit longer for inspection
        print("Closing browser...")
        scraper.cleanup()

if __name__ == "__main__":
    test_navigation_steps()