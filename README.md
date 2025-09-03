# EPC Certificate Scraper

This tool automatically downloads EPC (Energy Performance Certificate) PDFs for properties listed in the "Spring Acres - Cantebury.xlsx" spreadsheet.

## 📁 Files Overview

- `epc_scraper.py` - Main scraping script
- `test_scraper.py` - Setup verification script
- `run_scraper.bat` - Windows batch file to run the scraper easily
- `Spring Acres - Cantebury.xlsx` - Input spreadsheet with 471 properties
- `Processed/` - Downloaded PDF certificates will be saved here
- `logs/` - Log files with detailed processing information

## 🚀 Quick Start

### Option 1: Double-click to run
Simply double-click `run_scraper.bat` to start the scraping process.

### Option 2: Command line
```cmd
.\.venv\Scripts\python.exe epc_scraper.py
```

### Option 3: Test setup first
```cmd
.\.venv\Scripts\python.exe test_scraper.py
```

## 📊 Spreadsheet Data

The spreadsheet contains 471 properties with the following key information:
- **UPRN**: Unique Property Reference Number
- **Scheme Abbreviation**: RSC (Riverside Square, Canterbury)
- **Development Plot Number**: Property number within development
- **Address Components**: Multiple address line fields
- **Post Code**: All properties are in Canterbury (CT1 1GE area)
- **Tenure**: SO (Shared Ownership)

## 🎯 How It Works

1. **Address Construction**: Combines address fields into searchable format
2. **Filename Generation**: Creates organized filenames like "EPC - RSC - 1.0 - SO - 56540000001.pdf"
3. **Web Scraping**: 
   - Navigates to gov.uk EPC search
   - Searches by postcode
   - Matches property address
   - Downloads PDF certificate
4. **Progress Tracking**: Logs all activities with success/failure status

## 📋 Example Output Filename

For the first property:
```
EPC - RSC - 1.0 - SO - 56540000001.pdf
```

Where:
- `RSC` = Scheme Abbreviation
- `1.0` = Development Plot Number  
- `SO` = Tenure (Shared Ownership)
- `56540000001` = UPRN

## 📈 Progress Monitoring

- **Console Output**: Real-time progress updates
- **Log Files**: Detailed logs saved to `logs/` directory with timestamp
- **Summary Report**: Final statistics and any failed downloads
- **Results CSV**: Detailed processing results with error information

## ⚙️ Configuration

The scraper is pre-configured for:
- Chrome browser automation
- 15-second timeout for web elements
- 2-second delay between property requests
- Automatic PDF download to `Processed/` folder
- Comprehensive logging

## 🛡️ Safety Features

- Respectful scraping with delays between requests
- Comprehensive error handling and logging
- Progress tracking and resume capability
- Clean filename sanitization
- Browser automation cleanup

## 📊 Expected Results

- **Total Properties**: 471
- **Estimated Time**: 30-60 minutes (depending on network and site response)
- **Success Rate**: Depends on address matching accuracy and site availability

## 🔧 Environment

- **Python**: 3.13.5 (virtual environment)
- **Key Packages**: pandas 2.3.2, selenium 4.35.0, openpyxl 3.1.5
- **Browser**: Chrome (automated via Selenium WebDriver)

## 📞 Support

Check the log files in the `logs/` directory for detailed information about any issues. The scraper provides comprehensive error reporting and progress tracking.