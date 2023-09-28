import configparser
import email
import imaplib
import logging
import os
import smtplib
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
import random

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Read configuration from the "config.txt" file
config = configparser.ConfigParser()
config.read('config.txt')

# Extract values from the configuration
smtp_server = config.get('SMTP', 'server')
smtp_port = int(config.get('SMTP', 'port'))
smtp_username = config.get('SMTP', 'username')
smtp_password = config.get('SMTP', 'app_password')
email_subject = config.get('Email', 'subject')
email_body = config.get('Email', 'body')
attachment_path = config.get('Email', 'attachment_path')
attachment_type = config.get('Email', 'attachment_type')
start_row = int(config.get('GoogleSheet', 'start_row'))
email_column = int(config.get('GoogleSheet', 'email_column'))
undelivered_column = int(config.get('GoogleSheet', 'undelivered_column'))
send_text = config.get('GoogleSheet', 'send_text')
undelivered_text = config.get('GoogleSheet', 'undelivered_text')
sheet_name = config.get('GoogleSheet', 'sheet_name')
log_file_path = 'log.txt'


# Function to clean up old log entries (older than 7 days)
def clean_up_old_logs(log_file_path, days_to_keep=7):
    try:
        cutoff_date = datetime.now() - timedelta(days=days_to_keep)
        if os.path.exists(log_file_path):
            with open(log_file_path, 'r') as log_file:
                lines = log_file.readlines()
            recent_lines = [line for line in lines if parse_log_entry_date(line) >= cutoff_date]
            with open(log_file_path, 'w') as log_file:
                log_file.writelines(recent_lines)
    except Exception as e:
        print(f"Error cleaning up old logs: {str(e)}")


# Function to parse the date from a log entry
def parse_log_entry_date(log_entry):
    try:
        date_str = log_entry.split(' - ', 1)[0]
        return datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
    except ValueError:
        return datetime.min


# Clean up old log entries before configuring logging
clean_up_old_logs(log_file_path)

# Set up logging
logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)


def log_and_print(message, level=logging.INFO):
    logging.log(level, message)
    print(message)


# Define the scope and credentials file
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('YOUR_CREDENTIALS.json', scope)

# Authenticate with Google Sheets API
gc = gspread.authorize(credentials)

# Open a specific Google Sheet by title
try:
    spreadsheet = gc.open(sheet_name)
except gspread.exceptions.SpreadsheetNotFound:
    log_and_print(f"Error: Google Sheet '{sheet_name}' not found.", level=logging.ERROR)
    exit()

# Select a specific worksheet by index
worksheet = spreadsheet.get_worksheet(0)


# Function to mark email as undelivered in the Google Sheet
def mark_email_as_undelivered(address):
    found_cell = worksheet.find(address)
    worksheet.update_cell(found_cell.row, undelivered_column, undelivered_text)
    log_and_print(f"Marked email as undelivered for: {address}")


# Connect to the mailbox using IMAP to check for undelivered emails
try:
    imap_server = imaplib.IMAP4_SSL('imap.gmail.com')
    imap_server.login(smtp_username, smtp_password)
    imap_server.select('INBOX')
    seven_days_ago = (datetime.now() - timedelta(days=7)).strftime('%d-%b-%Y')
    search_criteria = f'(UNSEEN FROM "mailer-daemon@googlemail.com" SINCE {seven_days_ago})'
    _, email_ids = imap_server.search(None, search_criteria)

    for email_id in email_ids[0].split():
        _, email_data = imap_server.fetch(email_id, '(RFC822)')
        raw_email = email_data[0][1]
        parsed_email = email.message_from_bytes(raw_email)

        email_body = None
        if parsed_email.is_multipart():
            for part in parsed_email.walk():
                if part.get_content_type() == "text/plain":
                    email_body = part.get_payload(decode=True).decode()
                    break

        if email_body and "Your message wasn't delivered" in email_body:
            start_index = email_body.find("to ") + 3
            end_index = email_body.find(" because")
            wrong_email_address = email_body[start_index:end_index]
            mark_email_as_undelivered(wrong_email_address)
            log_and_print(f"Marked email as undelivered for: {wrong_email_address}")

    imap_server.logout()
    log_and_print("Undelivered emails marked in the Google Sheet.")
except Exception as e:
    log_and_print(f"Error checking mailbox: {str(e)}", level=logging.ERROR)

# Get all values from the specified email_column starting from the specified start_row
column_e_values = worksheet.col_values(email_column)
undelivered_values = worksheet.col_values(undelivered_column)

# Create a list of email addresses that have an empty value in the undelivered_column
filtered_values = []
for i, value in enumerate(column_e_values, start=1):
    if i >= start_row and value != '' and undelivered_values[i - 1] == '':
        filtered_values.append(value)

# Create and send emails with attachment
for row_number, email_address in enumerate(filtered_values, start=start_row):
    msg = MIMEMultipart()
    msg['From'] = smtp_username
    msg['To'] = email_address
    msg['Subject'] = email_subject

    # Add email body
    msg.attach(MIMEText(email_body, 'plain'))

    # Add attachment
    with open(attachment_path, 'rb') as attachment_file:
        attachment = MIMEApplication(attachment_file.read(), _subtype=attachment_type)
        attachment.add_header("Content-Disposition", "attachment", filename=os.path.basename(attachment_path))
        msg.attach(attachment)

    try:
        # Generate a random delay between 5 and 12 seconds
        delay_seconds = random.uniform(5, 12)
        time.sleep(delay_seconds)  # Sleep for the random delay

        # Connect to the SMTP server and send the email
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(smtp_username, email_address, msg.as_string())
        server.quit()
        log_and_print(f"Email sent to {email_address} (Row {row_number})")

        # Update the cell with the send_text when email is sent
        cell = worksheet.find(email_address)
        worksheet.update_cell(cell.row, undelivered_column, send_text)
    except Exception as e:
        log_and_print(f"Failed to send email to {email_address} (Row {row_number}): {str(e)}", level=logging.ERROR)
