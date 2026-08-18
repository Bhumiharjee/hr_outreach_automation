$EmailCount = 500
$WhatsAppCount = 500

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectDir
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

Write-Host "=== HR CV Outreach WhatsApp Auto-Send v5 ===" -ForegroundColor Cyan

if (!(Test-Path ".\.venv\Scripts\python.exe")) {
    python -m venv .venv
}
$Py = Join-Path $ProjectDir ".venv\Scripts\python.exe"

& $Py -m pip install -r requirements.txt
& $Py -m playwright install chromium

if (!(Test-Path ".\.env")) { throw ".env not found." }
if (!(Test-Path ".\hr details.xlsx")) { throw "hr details.xlsx not found." }
if (!(Test-Path ".\SAURAV KUMAR UPDATED CV.pdf")) { throw "SAURAV KUMAR UPDATED CV.pdf not found." }

& $Py hr_outreach.py --mode prepare

Write-Host "Test count: $EmailCount" -ForegroundColor Green
$run = Read-Host "Type RUN to start"
if ($run -ne "RUN") { exit }

& $Py hr_outreach.py --mode both --limit $EmailCount
Read-Host "Press Enter to exit"
