#!/usr/bin/env python3
"""
Test the complete EPC scraper flow with just one property
"""

import pandas as pd
from epc_scraper import EPCCertificateScraper

def test_complete_flow():
    """Test the complete scraper flow with one property"""
    print("Testing Complete EPC Scraper Flow...")
    print("=" * 50)
    
    # Initialize scraper
    scraper = EPCCertificateScraper()
    
    try:
        # Load just the first property
        df = pd.read_excel("Spring Acres - Cantebury.xlsx")
        sample_row = df.iloc[0]  # First property only
        
        print("Testing with first property from spreadsheet...")
        
        # Process the single property
        success = scraper.process_single_property(sample_row)
        
        if success:
            print("🎉 SUCCESS! Complete flow worked for first property!")
            print("✅ Navigation")
            print("✅ Domestic property selection") 
            print("✅ Postcode entry")
            print("✅ Find button click")
            print("✅ Address selection")
            print("✅ PDF download attempt")
        else:
            print("❌ Flow failed at some step - check logs for details")
        
        # Show results
        if scraper.results:
            result = scraper.results[0]
            print(f"\nResult:")
            print(f"  Address: {result['Address']}")
            print(f"  Status: {result['Status']}")
            print(f"  Filename: {result['Filename']}")
            if result['Error']:
                print(f"  Error: {result['Error']}")
        
    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Cleaning up...")
        scraper.cleanup()

if __name__ == "__main__":
    test_complete_flow()