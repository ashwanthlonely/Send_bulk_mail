# Send_bulk_mail User Manual

## Overview
This project contains Python scripts to:
- Send bulk HTML emails from one or more sender accounts with per-account daily limits.
- Track delivery status back into an Excel file.
- Merge multiple Excel files into a single dataset.

## Requirements
- Python 3.x
- Packages: `pandas`, `tqdm`, `pyyaml`
- Network access to the configured SMTP/IMAP servers.

## Files and Purpose
- `go_mail.py`: bulk sender for a Job-Assured training program campaign.
- `e2mail2009.py`: bulk sender for EarEase internship campaign.
- `e2_thanks.py`: webinar registration/thank-you email sender.
- `e2_remainder.py`: webinar reminder email sender.
- `merge.py`: merge multiple Excel files in a folder into one file.
- `data.py`: combine Excel files and perform a merge by a specified column.
- `e2email_accounts.yaml`: sender accounts for EarEase scripts.
- `email_accounts.yaml`: sender accounts for `go_mail.py`.

## Setup
1) Install dependencies:
   - `pip install pandas tqdm pyyaml`
2) Confirm the SMTP/IMAP servers and ports in each script.
3) Populate account YAML files with valid credentials:
   - `e2email_accounts.yaml` for EarEase scripts.
   - `email_accounts.yaml` for `go_mail.py`.
4) Verify Excel file paths inside each script match your environment.

## Email Account Files (YAML)
Each account entry should include:
- `email`: sender address
- `password`: sender password
- `emails_sent`: counter (integer)
- `last_sent`: timestamp string `YYYY-MM-DD HH:MM:SS`

Example:
```
email_accounts:
  - email: sender@example.com
    password: yourpassword
    emails_sent: 0
    last_sent: '2025-01-01 00:00:00'
```

## How to Run (Bulk Email Scripts)
General steps for:
- `go_mail.py`
- `e2mail2009.py`
- `e2_thanks.py`
- `e2_remainder.py`

1) Update the Excel file path in the script.
2) Ensure the Excel sheet has the required columns.
3) Run:
   - `python <script_name>.py`

Behavior:
- Skips rows with `Status = Sent`.
- Sends HTML email and appends to the IMAP `Sent` mailbox.
- Updates `Status` to `Sent` or `Failed: <error>`.
- Rotates accounts when per-account daily limit is reached.

## Excel Input Columns
- `go_mail.py`: `Name`, `Email ID`
- `e2mail2009.py`: `Name`, `Email ID`
- `e2_thanks.py`: `Your name`, `Your email`
- `e2_remainder.py`: `Your name`, `Your email`

## Merge Utilities
### merge.py
- Merges all `.xlsx`/`.xls` files in a folder into one file.
- Update `folder_path` and `save_path` before running.

Run:
- `python merge.py`

### data.py
- Reads all `.xlsx` files in a folder, concatenates, then merges on a specified column.
- Update `dir_path` and `merge_column` before running.

Run:
- `python data.py`

## Troubleshooting
- **Authentication errors**: verify SMTP/IMAP credentials and provider access policies.
- **Nothing sent**: check `Status` column or required input columns.
- **File in use**: close Excel before running to allow writing.
- **Daily limit reached**: wait 24 hours or reset counters manually.

## Security Notes
- Do not commit account passwords to source control.
- Share credentials securely with trusted operators only.
- Ensure recipients have consent for bulk email.
