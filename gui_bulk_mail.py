import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import smtplib
import imaplib
import yaml
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta

DEFAULT_SMTP_SERVER = 'smtpout.secureserver.net'
DEFAULT_SMTP_PORT = 587
DEFAULT_IMAP_SERVER = 'imap.secureserver.net'
DEFAULT_IMAP_PORT = 993
DEFAULT_EMAIL_LIMIT = 500
DEFAULT_REFRESH_DAYS = 1
DEFAULT_YAML = 'email_accounts.yaml'
ALT_YAML = 'e2email_accounts.yaml'


def load_yaml_accounts(path):
    with open(path, 'r') as f:
        cfg = yaml.safe_load(f) or {}
    accounts = cfg.get('email_accounts', [])
    return cfg, accounts


def update_yaml(path, cfg):
    with open(path, 'w') as f:
        yaml.safe_dump(cfg, f)


def reset_email_count_if_needed(account, refresh_interval):
    if 'emails_sent' not in account:
        account['emails_sent'] = 0
    if 'last_sent' not in account:
        account['last_sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    last_sent_time = datetime.strptime(account['last_sent'], '%Y-%m-%d %H:%M:%S')
    if datetime.now() - last_sent_time >= refresh_interval:
        account['emails_sent'] = 0
        return True
    return False


def get_next_available_account(accounts, limit, refresh_interval):
    for index, account in enumerate(accounts):
        if reset_email_count_if_needed(account, refresh_interval) or account['emails_sent'] < limit:
            return index
    return None


def connect_smtp(server, port, email, password):
    smtp = smtplib.SMTP(server, port)
    smtp.starttls()
    smtp.login(email, password)
    return smtp


def connect_imap(server, port, email, password):
    mail = imaplib.IMAP4_SSL(server, port)
    mail.login(email, password)
    return mail


class BulkMailApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Bulk Mail Sender')
        self.df = None
        self.csv_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.yaml_path = tk.StringVar(value=self._default_yaml_path())
        self.smtp_server = tk.StringVar(value=DEFAULT_SMTP_SERVER)
        self.smtp_port = tk.StringVar(value=str(DEFAULT_SMTP_PORT))
        self.imap_server = tk.StringVar(value=DEFAULT_IMAP_SERVER)
        self.imap_port = tk.StringVar(value=str(DEFAULT_IMAP_PORT))
        self.limit_per_account = tk.StringVar(value=str(DEFAULT_EMAIL_LIMIT))
        self.refresh_days = tk.StringVar(value=str(DEFAULT_REFRESH_DAYS))
        self.subject = tk.StringVar()
        self.use_html = tk.BooleanVar(value=True)
        self.append_sent = tk.BooleanVar(value=True)
        self.auto_rotate = tk.BooleanVar(value=True)
        self.email_col = tk.StringVar()
        self.name_col = tk.StringVar()
        self.account_choice = tk.StringVar()
        self.accounts = []
        self.account_cfg = {}

        self._build_ui()
        self._load_accounts_silent()

    def _default_yaml_path(self):
        if os.path.exists(DEFAULT_YAML):
            return DEFAULT_YAML
        if os.path.exists(ALT_YAML):
            return ALT_YAML
        return DEFAULT_YAML

    def _build_ui(self):
        pad = {'padx': 6, 'pady': 4}
        frm = ttk.Frame(self.root)
        frm.pack(fill='both', expand=True)

        file_row = ttk.Frame(frm)
        file_row.pack(fill='x', **pad)
        ttk.Label(file_row, text='CSV file').pack(side='left')
        ttk.Entry(file_row, textvariable=self.csv_path, width=60).pack(side='left', padx=6)
        ttk.Button(file_row, text='Browse', command=self.browse_csv).pack(side='left')

        out_row = ttk.Frame(frm)
        out_row.pack(fill='x', **pad)
        ttk.Label(out_row, text='Output CSV').pack(side='left')
        ttk.Entry(out_row, textvariable=self.output_path, width=60).pack(side='left', padx=6)
        ttk.Button(out_row, text='Browse', command=self.browse_output).pack(side='left')

        map_row = ttk.Frame(frm)
        map_row.pack(fill='x', **pad)
        ttk.Label(map_row, text='Email column').pack(side='left')
        self.email_menu = ttk.OptionMenu(map_row, self.email_col, '')
        self.email_menu.pack(side='left', padx=6)
        ttk.Label(map_row, text='Name column (optional)').pack(side='left', padx=6)
        self.name_menu = ttk.OptionMenu(map_row, self.name_col, '')
        self.name_menu.pack(side='left')

        subject_row = ttk.Frame(frm)
        subject_row.pack(fill='x', **pad)
        ttk.Label(subject_row, text='Subject').pack(side='left')
        ttk.Entry(subject_row, textvariable=self.subject, width=80).pack(side='left', padx=6)

        ttk.Label(frm, text='Message body (use {name} and {email} placeholders)').pack(anchor='w', **pad)
        self.body_text = tk.Text(frm, height=12)
        self.body_text.pack(fill='both', expand=True, padx=6)

        options_row = ttk.Frame(frm)
        options_row.pack(fill='x', **pad)
        ttk.Checkbutton(options_row, text='Send as HTML', variable=self.use_html).pack(side='left')
        ttk.Checkbutton(options_row, text='Append to IMAP Sent', variable=self.append_sent).pack(side='left', padx=10)
        ttk.Checkbutton(options_row, text='Auto-rotate accounts', variable=self.auto_rotate).pack(side='left', padx=10)

        smtp_row = ttk.Frame(frm)
        smtp_row.pack(fill='x', **pad)
        ttk.Label(smtp_row, text='SMTP server').pack(side='left')
        ttk.Entry(smtp_row, textvariable=self.smtp_server, width=28).pack(side='left', padx=6)
        ttk.Label(smtp_row, text='Port').pack(side='left')
        ttk.Entry(smtp_row, textvariable=self.smtp_port, width=6).pack(side='left', padx=6)
        ttk.Label(smtp_row, text='IMAP server').pack(side='left', padx=6)
        ttk.Entry(smtp_row, textvariable=self.imap_server, width=28).pack(side='left')
        ttk.Label(smtp_row, text='Port').pack(side='left', padx=6)
        ttk.Entry(smtp_row, textvariable=self.imap_port, width=6).pack(side='left')

        acct_row = ttk.Frame(frm)
        acct_row.pack(fill='x', **pad)
        ttk.Label(acct_row, text='Accounts YAML').pack(side='left')
        ttk.Entry(acct_row, textvariable=self.yaml_path, width=40).pack(side='left', padx=6)
        ttk.Button(acct_row, text='Browse', command=self.browse_yaml).pack(side='left')
        ttk.Button(acct_row, text='Load', command=self.load_accounts).pack(side='left', padx=6)
        ttk.Label(acct_row, text='Account').pack(side='left', padx=6)
        self.account_menu = ttk.OptionMenu(acct_row, self.account_choice, '')
        self.account_menu.pack(side='left')

        limit_row = ttk.Frame(frm)
        limit_row.pack(fill='x', **pad)
        ttk.Label(limit_row, text='Daily limit').pack(side='left')
        ttk.Entry(limit_row, textvariable=self.limit_per_account, width=8).pack(side='left', padx=6)
        ttk.Label(limit_row, text='Reset days').pack(side='left')
        ttk.Entry(limit_row, textvariable=self.refresh_days, width=6).pack(side='left', padx=6)

        action_row = ttk.Frame(frm)
        action_row.pack(fill='x', **pad)
        ttk.Button(action_row, text='Send Emails', command=self.send_emails).pack(side='left')
        self.status_label = ttk.Label(action_row, text='Idle')
        self.status_label.pack(side='left', padx=10)
        self.counts_label = ttk.Label(action_row, text='Pending: 0 | Sent: 0 | Failed: 0')
        self.counts_label.pack(side='left', padx=10)

    def browse_csv(self):
        path = filedialog.askopenfilename(filetypes=[('CSV files', '*.csv'), ('All files', '*.*')])
        if path:
            self.csv_path.set(path)
            if not self.output_path.get():
                self.output_path.set(path)
            self.load_csv(path)

    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files', '*.csv')])
        if path:
            self.output_path.set(path)

    def browse_yaml(self):
        path = filedialog.askopenfilename(
            filetypes=[('YAML files', '*.yaml;*.yml'), ('All files', '*.*')]
        )
        if path:
            self.yaml_path.set(path)
            self.load_accounts()

    def load_csv(self, path):
        try:
            self.df = pd.read_csv(path)
        except Exception as exc:
            messagebox.showerror('Error', f'Failed to read CSV: {exc}')
            return

        cols = list(self.df.columns)
        self._set_option_menu(self.email_menu, self.email_col, cols)
        self._set_option_menu(self.name_menu, self.name_col, [''] + cols)
        if 'Email' in cols:
            self.email_col.set('Email')
        elif 'Email ID' in cols:
            self.email_col.set('Email ID')
        self.status_label.config(text=f'Loaded {len(self.df)} rows')

    def _set_option_menu(self, menu, var, options):
        menu['menu'].delete(0, 'end')
        for opt in options:
            menu['menu'].add_command(label=opt, command=tk._setit(var, opt))
        if options:
            var.set(options[0])

    def _load_accounts_silent(self):
        try:
            self.account_cfg, self.accounts = load_yaml_accounts(self.yaml_path.get())
            self._refresh_account_menu()
        except Exception:
            pass

    def load_accounts(self):
        if not self.yaml_path.get():
            messagebox.showerror('Error', 'Please select an accounts YAML file.')
            return
        if not os.path.exists(self.yaml_path.get()):
            messagebox.showerror('Error', f'Accounts file not found: {self.yaml_path.get()}')
            return
        try:
            self.account_cfg, self.accounts = load_yaml_accounts(self.yaml_path.get())
        except Exception as exc:
            messagebox.showerror('Error', f'Failed to load accounts: {exc}')
            return
        self._refresh_account_menu()

    def _refresh_account_menu(self):
        labels = [a.get('email', 'unknown') for a in self.accounts]
        self._set_option_menu(self.account_menu, self.account_choice, labels)

    def send_emails(self):
        if self.df is None:
            messagebox.showerror('Error', 'Please load a CSV file.')
            return
        if not self.email_col.get():
            messagebox.showerror('Error', 'Please select the email column.')
            return
        if not self.accounts:
            messagebox.showerror('Error', 'No accounts loaded.')
            return

        try:
            limit = int(self.limit_per_account.get())
            refresh_days = int(self.refresh_days.get())
        except ValueError:
            messagebox.showerror('Error', 'Daily limit and reset days must be numbers.')
            return

        refresh_interval = timedelta(days=refresh_days)
        output_path = self.output_path.get() or self.csv_path.get()
        subject = self.subject.get().strip()
        body = self.body_text.get('1.0', 'end').strip()
        if not subject:
            messagebox.showerror('Error', 'Please enter a subject.')
            return
        if not body:
            messagebox.showerror('Error', 'Please enter a message body.')
            return

        if 'Status' not in self.df.columns:
            self.df['Status'] = ''

        total = len(self.df)
        sent_count = 0
        failed_count = 0
        pending_total = int((self.df['Status'] != 'Sent').sum())
        self.counts_label.config(text=f'Pending: {pending_total} | Sent: 0 | Failed: 0')

        account_index = None
        smtp = None
        imap = None

        try:
            if self.auto_rotate.get():
                account_index = get_next_available_account(self.accounts, limit, refresh_interval)
            else:
                if not self.account_choice.get():
                    messagebox.showerror('Error', 'Please select an account or enable auto-rotate.')
                    return
                account_index = next((i for i, a in enumerate(self.accounts)
                                      if a.get('email') == self.account_choice.get()), None)

            if account_index is None:
                messagebox.showinfo('Info', 'No available accounts with remaining quota.')
                return

            for idx, row in self.df.iterrows():
                if row.get('Status') == 'Sent':
                    continue

                name = row.get(self.name_col.get(), '') if self.name_col.get() else ''
                email = row.get(self.email_col.get(), '')
                if not email:
                    self.df.at[idx, 'Status'] = 'Failed: missing email'
                    failed_count += 1
                    self.counts_label.config(
                        text=f'Pending: {pending_total} | Sent: {sent_count} | Failed: {failed_count}'
                    )
                    continue

                if smtp is None:
                    acct = self.accounts[account_index]
                    smtp = connect_smtp(self.smtp_server.get(), int(self.smtp_port.get()),
                                        acct['email'], acct['password'])
                    if self.append_sent.get():
                        imap = connect_imap(self.imap_server.get(), int(self.imap_port.get()),
                                            acct['email'], acct['password'])

                body_text = body
                try:
                    body_text = body.format(name=name, email=email)
                except Exception:
                    body_text = body

                message = MIMEMultipart()
                subtype = 'html' if self.use_html.get() else 'plain'
                message.attach(MIMEText(body_text, subtype))
                message['From'] = self.accounts[account_index]['email']
                message['To'] = email
                message['Subject'] = subject

                try:
                    smtp.sendmail(self.accounts[account_index]['email'], [email], message.as_string())
                    if imap is not None:
                        imap.append('Sent', None, None, message.as_bytes())
                    self.df.at[idx, 'Status'] = 'Sent'
                    sent_count += 1
                    pending_total -= 1

                    self.accounts[account_index]['emails_sent'] = self.accounts[account_index].get('emails_sent', 0) + 1
                    self.accounts[account_index]['last_sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.account_cfg['email_accounts'] = self.accounts
                    update_yaml(self.yaml_path.get(), self.account_cfg)
                except Exception as exc:
                    self.df.at[idx, 'Status'] = f'Failed: {exc}'
                    failed_count += 1
                    pending_total -= 1

                self.status_label.config(text=f'Sent {sent_count}/{total}')
                self.counts_label.config(
                    text=f'Pending: {pending_total} | Sent: {sent_count} | Failed: {failed_count}'
                )
                self.root.update_idletasks()

                if self.accounts[account_index]['emails_sent'] >= limit:
                    if smtp:
                        smtp.quit()
                    if imap:
                        imap.logout()
                    smtp = None
                    imap = None

                    if self.auto_rotate.get():
                        account_index = get_next_available_account(self.accounts, limit, refresh_interval)
                        if account_index is None:
                            break
                    else:
                        break

            self.df.to_csv(output_path, index=False)
            messagebox.showinfo(
                'Done',
                f'Pending: {pending_total}, Sent: {sent_count}, Failed: {failed_count}. '
                f'Output saved to {output_path}'
            )
        finally:
            if smtp:
                smtp.quit()
            if imap:
                imap.logout()


if __name__ == '__main__':
    root = tk.Tk()
    app = BulkMailApp(root)
    root.mainloop()
