@echo off
title EPC Certificate Scraper
echo ============================================
echo         EPC Certificate Scraper
echo ============================================
echo.
echo Starting the EPC scraper...
echo Make sure you have your Excel file ready.
echo.
.\.venv\Scripts\python.exe epc_scraper.py
echo.
echo ============================================
echo Scraping completed!
echo Check the Processed folder for your PDFs.
echo Check the logs folder for detailed results.
echo ============================================
pause