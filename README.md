# Send_bulk_mail

## Quick Start (EXE)
1) Copy `dist/cmg_mails.exe` to a folder of your choice.
2) Place your account YAML in the same folder (or browse to it in the GUI).
3) Prepare a CSV with recipient emails.
4) Run `cmg_mails.exe`.

## GUI Usage
- Use **Browse** to load your CSV.
- Select the email column (required) and name column (optional).
- Enter a subject and message body.
- You can use `{name}` and `{email}` placeholders in the body.
- Choose your accounts YAML and click **Load**.
- Click **Send Emails**.

## Samples
- `docs/sample_recipients.csv`
- `docs/sample_accounts.yaml`

## Notes
- The GUI adds a `Status` column to the output CSV.
- If you hit daily limits, wait 24 hours or reset counters in YAML.
- Do not commit real credentials to Git.
