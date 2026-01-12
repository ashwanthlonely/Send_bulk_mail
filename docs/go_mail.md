# go_mail.py User Manual

## Purpose
Send personalized HTML emails in bulk from multiple sender accounts, track delivery status in the Excel file, and enforce per-account daily limits.

## Inputs
- Excel file: `C:\Users\ashwa\OneDrive\Desktop\Nonvoice_merged_data1.xlsx`
  - Required columns: `Name`, `Email ID`
  - Optional column: `Status` (created if missing)
- Account file: `email_accounts.yaml`
  - Key: `email_accounts` list with `email`, `password`, `emails_sent`, `last_sent`

## Configuration
- SMTP server: `smtpout.secureserver.net`, port `587`
- IMAP server: `imap.secureserver.net`, port `993`
- Daily quota per account: `500`
- Reset interval: `24 hours`
- Subject and HTML body are defined inline in the script.

## How to run
1) Update `email_accounts.yaml` with valid accounts and passwords.
2) Confirm the Excel file path and columns.
3) Run:
   - `python go_mail.py`

## What it does
- Loads recipient data from Excel.
- Skips rows where `Status` is `Sent`.
- Sends HTML email, appends the sent message to the IMAP `Sent` folder.
- Updates `Status` to `Sent` or `Failed: <error>`.
- Rotates to the next account after the per-account quota is reached.
- Persists per-account `emails_sent` and `last_sent` to `email_accounts.yaml`.

## Outputs
- Updates the same Excel file with `Status` values.
- Updates `email_accounts.yaml` with counters and timestamps.
- Prints progress and total sent count.

## Troubleshooting / Notes
- Ensure SMTP/IMAP credentials are correct and not blocked by the provider.
- Make sure the Excel file is not open in another app when writing.
- If you hit daily limits, wait 24 hours or reset counters manually.

## Safety
- Store credentials securely and avoid committing them to source control.
- Confirm recipient consent and comply with email regulations.
