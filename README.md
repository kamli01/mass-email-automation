# Certificate Email Sender

A Python script to automatically send certificate PNG files to participants via email using Gmail SMTP.

## Overview

This project provides two scripts for sending certificates to workshop participants:

- **`send_certificates_final.py`** - Production script using `Participants.xlsx`
- **`send_certificates_test.py`** - Testing script using `participants2.xlsx`

The scripts match participant names from an Excel file with certificate PNG files in the folder, then automatically send them via email with professional HTML-formatted email templates.

## Features

✅ Reads participant data from Excel files  
✅ Matches participant names with certificate PNG files  
✅ Sends certificates via Gmail SMTP with attachments  
✅ Professional HTML email templates  
✅ Fuzzy name matching for better accuracy  
✅ Dry-run mode to preview matches before sending  
✅ Detailed logging and error handling  

## Requirements

- Python 3.6+
- `openpyxl` - For reading Excel files

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/your-username/certificate-email-sender.git
   cd certificate-email-sender
   ```

2. **Install dependencies**
   ```bash
   pip install openpyxl
   ```

## Setup

### Gmail App Password

To use this script, you need to set up an **App Password** on your Google account:

1. Go to [Google Account Security](https://myaccount.google.com/security)
2. Enable **2-Step Verification** (if not already enabled)
3. Go to **App Passwords** and create a new app password
4. Copy the 16-character code

### Configuration

Edit the CONFIG section in your chosen script (`send_certificates_final.py` or `send_certificates_test.py`):

```python
SENDER_NAME     = "Your Organization Name"
SENDER_EMAIL    = "your-email@gmail.com"
SENDER_PASSWORD = "your-16-char-app-password"
XLSX_FILE       = "Participants.xlsx"          # Path to Excel file
CERTS_DIR       = "."                          # Folder containing certificates
DRY_RUN         = True                         # Preview mode (set False to send)
```

## Excel File Format

Your Excel file should have the following columns:
- **Name** - Participant's name (used to match certificates)
- **Email** - Participant's email address

Example:
| Name | Email |
|------|-------|
| John Doe | john@example.com |
| Jane Smith | jane@example.com |

## Certificate File Naming

Certificate PNG files should match participant names in the format:
```
John Doe.png
Jane Smith.png
```

The script uses fuzzy matching, so minor variations are tolerated.

## Usage

### Step 1: Preview Matches (Dry Run)

Always test first with `DRY_RUN = True` to verify name-to-certificate matching:

```bash
python send_certificates_final.py
```

This will show you:
- Matched certificates
- Email addresses
- Any unmatched participants

### Step 2: Send Certificates

Once you've verified the matches, set `DRY_RUN = False`:

```python
DRY_RUN = False
```

Then run the script:

```bash
python send_certificates_final.py
```

## Email Template

The script sends emails with:
- **Professional HTML formatting** with your organization's branding
- **Plain text fallback** for email clients that don't support HTML
- **Certificate attached** as PNG file
- **Custom subject line** (configurable)

## Troubleshooting

**Gmail authentication failed:**
- Verify you're using an **App Password**, not your regular Gmail password
- Ensure 2-Step Verification is enabled on your account
- Check that the 16-character password is copied correctly

**Certificates not found:**
- Ensure certificate PNG files are in the same folder as the script
- Check that file names match participant names
- The script uses fuzzy matching, but exact matches are preferred

**No matches found:**
- Verify participant names in Excel match certificate file names
- Check for extra spaces or special characters
- The fuzzy matching tolerance is ~80% similarity

## File Structure

```
├── send_certificates_final.py    # Production script
├── send_certificates_test.py     # Testing script
├── Participants.xlsx              # Production participant list
├── participants2.xlsx             # Test participant list
├── certificates/                  # Certificate PNG files
└── README.md                       # This file
```

## Security Notes

⚠️ **Important:** Never commit credentials to version control!

- Add `send_certificates_*.py` to `.gitignore` after removing passwords
- Use environment variables or a separate config file for sensitive data
- Regenerate App Passwords after committing them

## License

This project is provided as-is for educational and organizational use.

## Support

For issues or questions, please create an issue in the repository.

---

**Created for:** GDGoC TCOER Workshop  
**Purpose:** Automated certificate distribution for workshop participants
