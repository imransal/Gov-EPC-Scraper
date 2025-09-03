@echo off
echo Starting EPC Certificate Scraper...
echo.
.\.venv\Scripts\python.exe epc_scraper.py
echo.
echo Scraping completed. Check the logs folder for detailed results.
pause