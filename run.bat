@echo off
echo HR CV Outreach Automation
echo.
echo 1. Prepare tracker
echo 2. Email
echo 3. WhatsApp
echo 4. Both
set /p choice=Select 1-4:
if "%choice%"=="1" python hr_outreach.py --mode prepare
if "%choice%"=="2" python hr_outreach.py --mode email
if "%choice%"=="3" python hr_outreach.py --mode whatsapp
if "%choice%"=="4" python hr_outreach.py --mode both
pause
