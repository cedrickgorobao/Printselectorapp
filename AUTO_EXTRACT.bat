@echo off
title XMAS Page Printer — Auto Extract ^& Launch
color 0A

echo.
echo  ==============================================
echo    XMAS Page Printer — Auto Extract ^& Launch
echo  ==============================================
echo.
echo  This will:
echo    1. Extract all 12 pages from SPECIAL_XMAS V2.pub
echo       as PNG images (first time only)
echo    2. Start a local web server
echo    3. Open the Page Selector app in your browser
echo.
echo  Keep this window open while using the app!
echo.
pause

python "%~dp0auto_extract.py"

if %ERRORLEVEL% NEQ 0 (
  echo.
  echo  Something went wrong. Check the error above.
  pause
)
