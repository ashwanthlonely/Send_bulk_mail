# e2_thanks.py User Manual

## Purpose
Send a thank-you / registration confirmation email for the EarEase webinar campaign.

## Inputs
- Excel file: `D:\EarEase\Data\How to Get Started in Software Engineering to Land a High-Paying Offer (Responses).xlsx`
  - Required columns: `Your name`, `Your email`
  - Optional column: `Status` (created if missing)
- Account file: `e2email_accounts.yaml`
  - Key: `email_accounts` list with `email`, `password`, `emails_sent`, `last_sent`

## Configuration
- SMTP server: `smtpout.secureserver.net`, port `587`
- IMAP server: `imap.secureserver.net`, port `993`
- Daily quota per account: `450`
- Reset interval: `24 hours`
- Subject and HTML body are defined inline in the script.

## How to run
1) Update `e2email_accounts.yaml` with valid accounts and passwords.
2) Confirm the Excel file path and columns.
3) Run:
   - `python e2_thanks.py`

## What it does
- Loads recipient data from Excel.
- Skips rows where `Status` is `Sent`.
- Sends HTML email, appends the sent message to the IMAP `Sent` folder.
- Updates `Status` to `Sent` or `Failed: <error>`.
- Rotates to the next account after the quota is reached.

## Outputs
- Updates the same Excel file with `Status` values.
- Updates `e2email_accounts.yaml` with counters and timestamps.
- Prints progress.

## Troubleshooting / Notes
- Ensure SMTP/IMAP credentials are correct.
- Make sure the Excel file is not open while writing.
- Edit the `subject` and `body` strings to customize the template.

## Safety
- Store credentials securely and avoid committing them to source control.
- Confirm recipient consent and comply with email regulations.
