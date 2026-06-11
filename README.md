# Mass Email Automation

Send certificates and emails to many people at once using Gmail.

---

## Quick Start (3 Steps)

### Step 1: Set Up Gmail

1. Open [Google Account Security](https://myaccount.google.com/security)
2. Turn on **2-Step Verification** if not already on
3. Go to **App passwords** → Select Mail → Select your device
4. Copy the 16-character password Google gives you

### Step 2: Prepare Your Files

**Create `Participants.xlsx`** in the same folder as the script:

| Name | Email |
|------|-------|
| John Doe | john@example.com |
| Jane Smith | jane@example.com |

**Add certificates** named like this:
- `john_doe_certificate.png`
- `jane_smith_certificate.png`

### Step 3: Update the Script

Open `send_certificates_final.py` and change:

```python
SENDER_NAME     = "Your Name"
SENDER_EMAIL    = "your-email@gmail.com"
SENDER_PASSWORD = "xxxx xxxx xxxx xxxx"  # Your 16-char password from Step 1
```

---

## Run It

**First Time (Test Mode):**
```bash
python send_certificates_final.py
```
This shows what will be sent without actually sending. Check if everything looks good.

**Second Time (Send for Real):**
1. Open the script
2. Change `DRY_RUN = False`
3. Run again: `python send_certificates_final.py`

Done! Emails are sent.

---

## Install Python Package

Before running, install openpyxl:
```bash
pip install openpyxl
```

---

## Problems?

| Problem | Fix |
|---------|-----|
| "Can't find name/email" | Check Excel has headers in row 1 |
| "Auth failed" | Use App Password, not regular Gmail password |
| "Certificates not matching" | Name files like: `FirstName_LastName_certificate.png` |
| "No module openpyxl" | Run: `pip install openpyxl` |

---

## Want to Change?

**Change subject line:**
```python
EMAIL_SUBJECT = "Your New Subject"
```

**Change sender name in email:**
```python
SENDER_NAME = "Your Organization"
```

**Change email folder:**
```python
CERTS_DIR = "path/to/certificates"  # Default is current folder
```

---

Done! Questions? Check the code comments or try dry-run mode first.
