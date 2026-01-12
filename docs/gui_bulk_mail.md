# gui_bulk_mail.py User Manual

## Purpose
Provide a desktop GUI to send bulk emails using a CSV file, map required fields, compose a message, and send using configured SMTP/IMAP accounts.

## Requirements
- Python 3.x
- Packages: `pandas`, `tqdm`, `pyyaml`
- A sender accounts YAML file (default: `email_accounts.yaml`)

## CSV Input
You can load any CSV and select which columns represent:
- Email address (required)
- Name (optional)

A `Status` column is added if it doesn't exist. Rows with `Status = Sent` are skipped.

## Message Personalization
Use placeholders in the body:
- `{name}`: value from the selected name column
- `{email}`: value from the selected email column

If the name column is not selected, `{name}` will be empty.

## Account File (YAML)
The YAML must contain an `email_accounts` list. Each entry should include:
- `email`
- `password`
- `emails_sent`
- `last_sent` (format: `YYYY-MM-DD HH:MM:SS`)

Example:
```
email_accounts:
  - email: sender@example.com
    password: yourpassword
    emails_sent: 0
    last_sent: '2025-01-01 00:00:00'
```

## How to Run
1) Activate your venv:
   - `..venv\Scripts\Activate.ps1`
2) Run the GUI:
   - `python gui_bulk_mail.py`

## GUI Fields
- **CSV file**: choose the recipient CSV.
- **Output CSV**: where to save updates (defaults to input file).
- **Email column**: select the email field from your CSV.
- **Name column**: optional, for `{name}` in the template.
- **Subject**: email subject line.
- **Message body**: use `{name}` and `{email}` placeholders.
- **Send as HTML**: send HTML or plain text.
- **Append to IMAP Sent**: store a copy in Sent folder.
- **Auto-rotate accounts**: switch to next account once the daily limit is reached.
- **SMTP/IMAP server + port**: defaults to GoDaddy servers.
- **Accounts YAML**: file containing sender accounts.
- **Daily limit / Reset days**: quota per account and refresh interval.

## Output
- Updates CSV with `Status` column values:
  - `Sent`
  - `Failed: <error>`
- Updates `emails_sent` and `last_sent` in the YAML.

## Troubleshooting
- **No module named pandas**: run `pip install -r requirements.txt` in the venv.
- **Auth errors**: verify SMTP/IMAP credentials and provider security settings.
- **CSV not saving**: close the CSV if it's open in Excel.
- **Nothing sent**: check `Status` column or empty email values.

## Security Notes
- Do not commit YAML credentials to Git.
- Share credentials securely and only with trusted users.
