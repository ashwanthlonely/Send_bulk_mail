import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tqdm import tqdm
import imaplib
import yaml
import time
from datetime import datetime, timedelta

# Load email accounts from YAML file
with open('e2email_accounts.yaml', 'r') as f:
    email_accounts_config = yaml.safe_load(f)
email_accounts = email_accounts_config['email_accounts']

# Load your DataFrame and add a "Status" column if it doesn't exist
excel_path = r"D:\EarEase\Data\How to Get Started in Software Engineering to Land a High-Paying Offer (Responses).xlsx"
df = pd.read_excel(excel_path)
if 'Status' not in df.columns:
    df['Status'] = ''  # Create a new column for Status if it doesn't exist

# Set up the SMTP and IMAP server details
smtp_server = 'smtpout.secureserver.net'
smtp_port = 587
imap_server = 'imap.secureserver.net'
imap_port = 993

# Email sending limit per account
email_limit_per_account = 450
email_refresh_interval = timedelta(days=1)  # 24 hours

# Function to update email count and timestamp in the YAML file
def update_email_count(account_index, emails_sent):
    email_accounts_config['email_accounts'][account_index]['emails_sent'] = emails_sent
    email_accounts_config['email_accounts'][account_index]['last_sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open('e2email_accounts.yaml', 'w') as f:
        yaml.safe_dump(email_accounts_config, f)

# Function to check if the limit should be reset (24 hours passed)
def check_reset_limit(account_index):
    # Initialize fields if they don't exist
    if 'emails_sent' not in email_accounts[account_index]:
        email_accounts[account_index]['emails_sent'] = 0
    if 'last_sent' not in email_accounts[account_index]:
        email_accounts[account_index]['last_sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    last_sent_str = email_accounts[account_index]['last_sent']
    last_sent_time = datetime.strptime(last_sent_str, '%Y-%m-%d %H:%M:%S')
    
    # Check if the 24-hour limit reset is applicable
    if datetime.now() - last_sent_time >= email_refresh_interval:
        email_accounts[account_index]['emails_sent'] = 0  # Reset the email count
        with open('e2email_accounts.yaml', 'w') as f:
            yaml.safe_dump(email_accounts_config, f)
        print(f"Email count reset for account {email_accounts[account_index]['email']} due to 24-hour refresh.")

# Function to connect to SMTP server
def connect_smtp(email, password):
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email, password)
    return server

# Function to connect to IMAP server
def connect_imap(email, password):
    mail = imaplib.IMAP4_SSL(imap_server, imap_port)
    mail.login(email, password)
    return mail

# Initialize account index and retrieve the email count and last sent time
account_index = 0
check_reset_limit(account_index)  # Reset the limit if 24 hours have passed
emails_sent_from_current_account = email_accounts[account_index]['emails_sent']

# Calculate the number of emails left to send
emails_left = df[df['Status'] != 'Sent'].shape[0]
print(f"Emails left to send: {emails_left}")

# Connect to the first email account's SMTP and IMAP
server = connect_smtp(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
mail = connect_imap(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

# Initialize progress bar
total_emails = len(df)
pbar = tqdm(total=total_emails, desc='Sending emails', unit='email')

# Count of total emails sent
total_emails_sent = 0

# Retry mechanism for failed emails
max_retries = 3

# Iterate through each row in the DataFrame and send an email
for index, row in df.iterrows():
    # Check if the email was already sent (by checking the "Status" column)
    if row['Status'] == 'Sent':
        pbar.update(1)
        continue  # Skip this email since it was already sent

    name = row['Your name']
    email = row['Your email']
    cc_email = ''

    # Create the email content
    subject = """Reminder: EarEase Tech Webinar Tomorrow - Excited to Meet You!""" # Subject of the email       

    body = f"""<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <p><strong>Dear {name},</strong></p>
    <p>
        This is a friendly reminder that our much-awaited webinar, 
        <strong>"How to Get Started in Software Engineering to Land a High-Paying Offer,"</strong> is happening tomorrow!
    </p>
    <p>
        <strong>📅 Date:</strong> 9th Jan 2025 <br>
        <strong>🕒 Time:</strong> 11:00 AM to 12:15 PM <br>
        <strong>📍 Platform:</strong> <a href="https://cuboulder.zoom.us/j/92587596298" target="_blank">https://cuboulder.zoom.us/j/92587596298</a>
    </p>
    <p>
        We’re thrilled to have you join us as <strong>Mr. Ashwanth Karibindi</strong> (CEO, EarEase Tech) and <strong>Mr. Deeptanshu Sankhwar</strong> (Project Lead) share 
        <strong>expert insights</strong> and <strong>actionable steps</strong> to launch your career in software engineering.
    </p>
    <p>
        Get ready to explore <strong>exciting opportunities</strong> and learn from <strong>seasoned professionals</strong>. We can’t wait to meet you virtually!
    </p>
    <p><strong>See you tomorrow!</strong></p>
    <p>
        <strong>Best Regards,</strong><br>
        <strong>Jasmine</strong><br>
        <strong>HR Manager</strong><br>
        <strong>EarEase Tech</strong><br>
        📞 <strong>+91-8309556828</strong><br>
        📧 <a href="mailto:hr@eareasetech.com"><strong>hr@eareasetech.com</strong></a><br>
        🌐 <a href="https://www.eareasetech.com/" target="_blank"><strong>https://www.eareasetech.com/</strong></a>
    </p>
</body>

"""
# To enroll, please complete the form below:</p>
# <p><a href="https://forms.gle/bBFBVkUVhQmnFg8Q6"><b>Enrollment Form</b></a></p>
# <p><b>Note: For any further clarifications, please reach out to us at hr@eareasetech.com</b>.</p>
    message = MIMEMultipart()
    message.attach(MIMEText(body, 'html'))

    message['From'] = email_accounts[account_index]['email']
    message['To'] = email
    message['Cc'] = cc_email
    message['Subject'] = subject

    retries = 0
    while retries < max_retries:
        try:
            # Send the email
            server.sendmail(email_accounts[account_index]['email'], [email], message.as_string())

            # Append the sent email to the 'Sent' mailbox
            mail.append('Sent', None, None, message.as_bytes())

            # Mark email as sent in the DataFrame
            df.at[index, 'Status'] = 'Sent'
            total_emails_sent += 1
            emails_sent_from_current_account += 1

            # Update email count and timestamp in the YAML file
            update_email_count(account_index, emails_sent_from_current_account)
            break
        except Exception as e:
            retries += 1
            if retries >= max_retries:
                # Mark the email as failed if retries are exhausted
                df.at[index, 'Status'] = f'Failed: {str(e)}'
                print(f"Failed to send email to {email}. Error: {str(e)}")

    # Update progress bar
    pbar.update(1)

    # Check if the current account has reached its limit
    if emails_sent_from_current_account >= email_limit_per_account:
        # Logout and switch to the next account
        server.quit()
        mail.logout()

        account_index += 1
        if account_index >= len(email_accounts):
            print("All email accounts have reached the limit.")
            break

        # Reset the email count if 24 hours have passed for the new account
        check_reset_limit(account_index)

        # Get the updated count for the new account
        emails_sent_from_current_account = email_accounts[account_index]['emails_sent']

        # Connect to the next email account
        server = connect_smtp(email_accounts[account_index]['email'], email_accounts[account_index]['password'])
        mail = connect_imap(email_accounts[account_index]['email'], email_accounts[account_index]['password'])

# Close the progress bar
pbar.close()

# Print total emails sent after the process is completed
print(f"Total emails sent: {total_emails_sent}")

# Save the updated Excel file
df.to_excel(excel_path, index=False)

# Quit the SMTP server and IMAP logout
server.quit()
mail.logout()