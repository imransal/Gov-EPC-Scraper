# EPC Certificate Scraper

This tool automatically downloads EPC (Energy Performance Certificate) PDFs for properties listed in Excel spreadsheets from the UK Government's energy certificate database.

## 📁 Project Structure

```
EPC_Scraper/
├── epc_scraper.py           # Main scraping script
├── run_scraper.bat          # Windows batch file to run the scraper
├── README.md                # This documentation
├── .gitignore              # Git ignore rules
├── logs/                   # Log files (created automatically)
├── Processed/              # Downloaded PDF certificates
└── .venv/                  # Python virtual environment
```

## 🚀 Quick Start

### Prerequisites
- Python 3.13+ 
- Chrome browser installed
- Excel file with property data

### Setup
1. Clone this repository
2. Install dependencies:
   ```bash
   .\.venv\Scripts\pip.exe install pandas openpyxl selenium
   ```

### Running the Scraper

**Option 1: Double-click to run**
```
Double-click run_scraper.bat
```

**Option 2: Command line**
```cmd
.\.venv\Scripts\python.exe epc_scraper.py
```

When prompted, enter the path to your Excel spreadsheet file.

## 📊 Input Data Format

The Excel spreadsheet should contain columns:
- **UPRN**: Unique Property Reference Number
- **Scheme Abbreviation**: Development abbreviation (e.g., "RSC")
- **Development Plot Number**: Property number within development
- **Address Line 1-5**: Property address components
- **Town**: Town/city
- **Post Code**: UK postcode
- **Tenure**: Property tenure type (e.g., "SO" for Shared Ownership)

## 🎯 How It Works

1. **Automated Navigation**: Opens gov.uk EPC search in Chrome
2. **Property Processing**: For each property in the spreadsheet:
   - Selects "Domestic property" option
   - Enters the postcode
   - Finds and selects the matching address
   - Downloads the EPC certificate as PDF
3. **Organized Output**: Saves PDFs with structured filenames
4. **Progress Tracking**: Logs all activities with success/failure status

## 📋 Output

### PDF Filenames
Files are saved with the format:
```
EPC - [Scheme] - [Plot] - [Tenure] - [UPRN].pdf
```

Example: `EPC - RSC - 1.0 - SO - 56540000001.pdf`

### Logging
- **Console Output**: Real-time progress updates
- **Log Files**: Detailed logs in `logs/` directory with timestamp
- **Summary Report**: Final statistics and failed downloads
- **Results CSV**: Processing results with error details

## ⚙️ Configuration

The scraper includes:
- 20-second timeout for web elements
- 2-second delay between property requests (respectful scraping)
- Automatic Chrome browser management
- Comprehensive error handling and logging
- Smart address matching with fallback options

## 🛡️ Features

- **Robust Address Matching**: Handles variations in address formats
- **Error Recovery**: Continues processing if individual properties fail
- **Progress Tracking**: Resume capability with detailed logging
- **File Management**: Automatic renaming and duplicate handling
- **Respectful Scraping**: Built-in delays and proper cleanup

## � Performance

- Processes ~30-60 properties per hour (depending on network speed)
- Handles large datasets (tested with 471+ properties)
- Memory efficient with automatic cleanup

## 🔧 Environment

- **Python**: 3.13.5 (virtual environment included)
- **Dependencies**: pandas, selenium, openpyxl
- **Browser**: Chrome (automated via Selenium WebDriver)
- **Platform**: Windows (PowerShell commands included)