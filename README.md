# Mass Email Automation

Send emails to many people at once using Gmail. Attach certificates, documents, or any files you need to send in bulk.

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

**Add attachments** (optional) named like this:
- `john_doe_certificate.png`
- `jane_smith_certificate.png`

Or use any other files you want to attach (PDFs, documents, etc.)

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
| "Attachments not matching" | Name files like: `FirstName_LastName_filename.ext` |
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

**Change attachments folder:**
```python
CERTS_DIR = "path/to/attachments"  # Default is current folder
```

---

## Use Cases

This tool works for more than just certificates:
- **Certificates** - Send personalized certificates to participants
- **Documents** - Distribute contracts, agreements, or PDFs
- **Invitations** - Send personalized event invites with attachments
- **Reports** - Email individual performance reports or statements
- **Confirmations** - Send receipts or confirmation letters
- **Any bulk email** - Attach any files and send to your list

---

Done! Questions? Check the code comments or try dry-run mode first.
