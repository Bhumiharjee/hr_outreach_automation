# HR Outreach Automation

An automated HR outreach tool that sends personalized emails and WhatsApp messages to HR professionals for job opportunities.

## Features

- **Automated Email Sending**: Send bulk emails via Gmail SMTP with customizable delays
- **WhatsApp Messaging**: Automated WhatsApp message delivery using Playwright browser automation
- **Contact Tracking**: Excel-based tracking of outreach attempts and responses
- **Dry Run Mode**: Test your configuration without sending actual messages
- **Configurable Delays**: Built-in delays to avoid being flagged as spam

## Prerequisites

- Python 3.8+
- Gmail account with app password enabled
- WhatsApp Web access (for WhatsApp automation)
- Playwright browser drivers

## Installation

1. Clone the repository:
```bash
git clone https://github.com/Bhumiharjee/hr_outreach_automation.git
cd hr_outreach_automation
```

2. Create a virtual environment:
```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Install Playwright browsers:
```bash
playwright install chromium
```

## Configuration

Create a `.env` file in the root directory with the following settings:

```env
# General Settings
DRY_RUN=false
CV_PATH=your_cv.pdf
TRACKER_PATH=outreach_tracker.xlsx
MAX_PER_RUN=20

# Sender Details
SENDER_NAME=Your Name
SENDER_EMAIL=your_email@gmail.com

# Email Configuration (Gmail SMTP)
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USERNAME=your_email@gmail.com
SMTP_PASSWORD=your_app_password
SMTP_USE_TLS=true

# Delays (in seconds) to avoid spam detection
EMAIL_DELAY_MIN=60
EMAIL_DELAY_MAX=120
WHATSAPP_DELAY_MIN=90
WHATSAPP_DELAY_MAX=180

# WhatsApp Settings
WHATSAPP_HEADLESS=false
WHATSAPP_TIMEOUT_MS=45000

# Message Template
EMAIL_SUBJECT=Application for Cloud / DevOps Engineer Opportunities
AUTO_SEND_WHATSAPP=true
```

### Important Notes

- **Gmail App Password**: Create an [app-specific password](https://myaccount.google.com/apppasswords) instead of using your regular Gmail password
- **WhatsApp**: Ensure you're logged into WhatsApp Web in the browser before running the script
- **DRY_RUN**: Set to `true` to test without sending actual messages

## Usage

### Windows (PowerShell)
```powershell
.\setup_and_run_hr_outreach.ps1
```

### Windows (Batch)
```bash
run.bat
```

### Linux/Mac
```bash
python hr_outreach.py
```

## File Structure

- `hr_outreach.py` - Main automation script
- `.env` - Configuration file
- `hr details.xlsx` - Contact list with HR details
- `outreach_tracker.xlsx` - Tracks sent emails and messages
- `requirements.txt` - Python dependencies
- `whatsapp_profile/` - WhatsApp browser profile data

## How It Works

1. **Reads Contacts**: Loads HR contact details from `hr details.xlsx`
2. **Personalizes Messages**: Creates customized emails and WhatsApp messages
3. **Sends Emails**: Uses Gmail SMTP to send emails with delays between each
4. **Sends WhatsApp Messages**: Automates WhatsApp Web using Playwright
5. **Tracks Progress**: Updates `outreach_tracker.xlsx` with results

## Delay Strategy

Built-in delays prevent your account from being flagged as spam:
- **Email Delays**: 1-2 minutes between emails
- **WhatsApp Delays**: 1.5-3 minutes between messages

Adjust `EMAIL_DELAY_MIN/MAX` and `WHATSAPP_DELAY_MIN/MAX` as needed.

## Troubleshooting

- **WhatsApp Connection Fails**: Ensure you're logged into WhatsApp Web and the session hasn't expired
- **Emails Not Sending**: Check Gmail app password and SMTP settings in `.env`
- **Script Crashes**: Check that all required files (CV, contact list) exist

## License

See LICENSE file for details.

## Disclaimer

Use this tool responsibly and in compliance with anti-spam laws and platform terms of service. Respect recipient preferences and always include unsubscribe options.
