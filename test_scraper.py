#!/usr/bin/env python3
"""
Test script to verify the EPC scraper setup and functionality
"""

import pandas as pd
import os
from epc_scraper import EPCCertificateScraper

def test_setup():
    """Test the basic setup and environment"""
    print("Testing EPC Scraper Setup...")
    print("=" * 50)
    
    # Test 1: Check if required packages are available
    try:
        import pandas
        import selenium
        import openpyxl
        print("✓ All required packages are installed")
        print(f"  - Pandas: {pandas.__version__}")
        print(f"  - Selenium: {selenium.__version__}")
    except ImportError as e:
        print(f"✗ Missing package: {e}")
        return False
    
    # Test 2: Check if spreadsheet exists and can be read
    spreadsheet_path = "Spring Acres - Cantebury.xlsx"
    try:
        df = pd.read_excel(spreadsheet_path)
        print(f"✓ Spreadsheet loaded successfully: {len(df)} rows")
        print(f"  - Required columns present: {all(col in df.columns for col in ['UPRN', 'Post Code', 'Address Line 1'])}")
    except Exception as e:
        print(f"✗ Failed to load spreadsheet: {e}")
        return False
    
    # Test 3: Check if directories exist
    required_dirs = ['Processed', 'logs']
    for dir_name in required_dirs:
        if os.path.exists(dir_name):
            print(f"✓ Directory '{dir_name}' exists")
        else:
            print(f"✗ Directory '{dir_name}' missing")
    
    # Test 4: Test Chrome driver
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        
        driver = webdriver.Chrome(options=options)
        driver.get("https://www.google.com")
        print("✓ Chrome driver working")
        driver.quit()
    except Exception as e:
        print(f"✗ Chrome driver issue: {e}")
        return False
    
    print("\n" + "=" * 50)
    print("Setup verification completed successfully!")
    return True

def show_sample_data():
    """Show sample data from the spreadsheet"""
    print("\nSample Data Analysis:")
    print("=" * 50)
    
    df = pd.read_excel("Spring Acres - Cantebury.xlsx")
    
    # Show first 3 rows of key columns
    key_columns = ['UPRN', 'Scheme Abbreviation', 'Development Plot Number', 
                   'Address Line 1', 'Address Line 2', 'Town', 'Post Code', 'Tenure']
    
    print("First 3 properties:")
    sample_data = df[key_columns].head(3)
    
    for idx, row in sample_data.iterrows():
        print(f"\nProperty {idx + 1}:")
        for col in key_columns:
            if pd.notna(row[col]) and str(row[col]).strip():
                print(f"  {col}: {row[col]}")
    
    # Show address construction example
    print(f"\nAddress construction example for first property:")
    first_row = df.iloc[0]
    
    scraper = EPCCertificateScraper()
    full_address = scraper.construct_full_address(first_row)
    filename = scraper.generate_filename(first_row)
    
    print(f"  Full Address: {full_address}")
    print(f"  Generated Filename: {filename}")
    print(f"  Postcode: {first_row.get('Post Code', 'N/A')}")

if __name__ == "__main__":
    if test_setup():
        show_sample_data()
        print(f"\n🎉 Everything is set up correctly!")
        print(f"📁 Total properties to process: {len(pd.read_excel('Spring Acres - Cantebury.xlsx'))}")
        print(f"💡 You can now run the main scraper with: python epc_scraper.py")
    else:
        print(f"\n❌ Setup issues detected. Please fix the above errors before proceeding.")