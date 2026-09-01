@echo off
cd /d "%~dp0"
py -3 --version >nul 2>nul && (start "" pyw -3 pdf_to_xlsx.py) || (start "" pythonw pdf_to_xlsx.py)
