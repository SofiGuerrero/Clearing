@echo off
echo ============================================
echo   SAP Clearing Dashboard - Setup & Launch
echo ============================================
echo.
echo Installing dependencies...
pip install flask pandas openpyxl pywin32 --quiet
echo.
echo Starting dashboard...
echo Open your browser at: http://localhost:5001
echo Press Ctrl+C to stop the server.
echo.
python server.py
pause
