#!/usr/bin/env python3
"""
Test the PDF download and renaming functionality
"""

import pandas as pd
import os
from epc_scraper import EPCCertificateScraper

def test_pdf_download_rename():
    """Test that PDF downloads and gets renamed correctly"""
    print("Testing PDF Download and Rename...")
    print("=" * 50)
    
    # Initialize scraper
    scraper = EPCCertificateScraper()
    
    try:
        # Clear any existing PDFs in the processed folder first
        processed_dir = scraper.download_dir
        print(f"Checking processed directory: {processed_dir}")
        
        existing_files = [f for f in os.listdir(processed_dir) if f.endswith('.pdf')]
        print(f"Found {len(existing_files)} existing PDF files")
        
        # Load sample property
        df = pd.read_excel("Spring Acres - Cantebury.xlsx")
        sample_row = df.iloc[0]  # First property
        
        expected_filename = scraper.generate_filename(sample_row)
        print(f"Expected filename: {expected_filename}")
        
        # Process the single property
        print("Running complete flow to test PDF download...")
        success = scraper.process_single_property(sample_row)
        
        if success:
            # Check if the file was created with correct name
            expected_path = os.path.join(processed_dir, expected_filename)
            
            if os.path.exists(expected_path):
                print(f"✅ SUCCESS! PDF saved with correct filename: {expected_filename}")
                
                # Check file size to make sure it's not empty
                file_size = os.path.getsize(expected_path)
                print(f"   File size: {file_size:,} bytes")
                
                if file_size > 1000:  # At least 1KB
                    print("✅ File appears to contain actual content")
                else:
                    print("⚠️  File seems very small - might be empty")
                    
            else:
                print(f"❌ Expected file not found: {expected_filename}")
                
                # List what files are actually in the directory
                current_files = [f for f in os.listdir(processed_dir) if f.endswith('.pdf')]
                print("Files currently in processed directory:")
                for file in current_files:
                    print(f"   - {file}")
        else:
            print("❌ Processing failed - check logs")
        
    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Cleaning up...")
        scraper.cleanup()

if __name__ == "__main__":
    test_pdf_download_rename()