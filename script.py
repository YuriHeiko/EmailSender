import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
import imaplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import logging
import os
from datetime import datetime, timedelta
import configparser  # Import the configparser module

# Read configuration from the "config.txt" file
config = configparser.ConfigParser()
config.read('config.txt')

# Extract values from the configuration
smtp_server = config.get('SMTP', 'server')
smtp_port = int(config.get('SMTP', 'port'))
smtp_username = config.get('SMTP', 'username')
smtp_password = config.get('SMTP', 'password')
email_subject = config.get('Email', 'subject')
email_body = config.get('Email', 'body')
start_row = int(config.get('GoogleSheet', 'start_row'))  # Extract start_row as an integer
email_column = int(config.get('GoogleSheet', 'email_column'))  # Extract email_column as an integer
undelivered_column = int(config.get('GoogleSheet', 'undelivered_column'))  # Extract undelivered_column as an integer
sheet_name = config.get('GoogleSheet', 'sheet_name')  # Extract sheet_name

# Define a custom log file path
log_file_path = 'log.txt'

# Function to clean up old log entries (older than 7 days)
def clean_up_old_logs(log_file_path, days_to_keep=7):
    try:
        # Calculate the date 7 days ago from today
        cutoff_date = datetime.now() - timedelta(days=days_to_keep)

        # Read the existing log file
        with open(log_file_path, 'r') as log_file:
            lines = log_file.readlines()

        # Filter log entries to keep only those within the last 7 days
        recent_lines = [line for line in lines if parse_log_entry_date(line) >= cutoff_date]

        # Write the filtered log entries back to the log file
        with open(log_file_path, 'w') as log_file:
            log_file.writelines(recent_lines)
    except Exception as e:
        print(f"Error cleaning up old logs: {str(e)}")

# Function to parse the date from a log entry
def parse_log_entry_date(log_entry):
    try:
        # Assuming the date is at the beginning of the log entry in the format 'YYYY-MM-DD'
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

# Define the scope and credentials file (replace 'YOUR_CREDENTIALS.json' with your actual JSON file)
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('YOUR_CREDENTIALS.json', scope)

# Authenticate with Google Sheets API
gc = gspread.authorize(credentials)

# Open a specific Google Sheet by title (using the sheet_name from the config file)
spreadsheet = gc.open(sheet_name)

# Select a specific worksheet by index (0-based) or by title
worksheet = spreadsheet.get_worksheet(0)

# Function to mark email as undelivered in the Google Sheet
def mark_email_as_undelivered(email_address):
    cell = worksheet.find(email_address)
    worksheet.update_cell(cell.row, undelivered_column, 'undelivered')

# Get all values from the specified email_column starting from the specified start_row
column_e_values = worksheet.col_values(email_column, value_render_option='UNFORMATTED_VALUE', include_tailing_empty=True)
filtered_values = [value for i, value in enumerate(column_e_values, start=start_row) if i >= start_row and value != '']

# Connect to the mailbox using IMAP to check for undelivered emails
try:
    imap_server = imaplib.IMAP4_SSL('imap.gmail.com')
    imap_server.login(smtp_username, smtp_password)
    imap_server.select('INBOX')

    # Calculate the date 7 days ago from today
    seven_days_ago = (datetime.now() - timedelta(days=7)).strftime('%d-%b-%Y')

    # Search for undelivered emails with the specified subject within the last 7 days
    search_criteria = f'(UNSEEN SUBJECT "{email_subject}" SINCE {seven_days_ago})'
    _, email_ids = imap_server.search(None, search_criteria)

    for email_id in email_ids[0].split():
        _, email_data = imap_server.fetch(email_id, '(RFC822)')
        raw_email = email_data[0][1]
        parsed_email = email.message_from_bytes(raw_email)

        recipient_email = parsed_email['To']
        mark_email_as_undelivered(recipient_email)
        log_and_print(f"Marked email as undelivered for: {recipient_email}")
        log_and_print(f"Email details:\n{parsed_email}", level=logging.DEBUG)

    imap_server.logout()
    log_and_print("Undelivered emails marked in the Google Sheet.")
except Exception as e:
    log_and_print(f"Error checking mailbox: {str(e)}", level=logging.ERROR)

# Create and send emails
for row_number, email_address in enumerate(filtered_values, start=start_row):
    msg = MIMEMultipart()
    msg['From'] = smtp_username
    msg['To'] = email_address
    msg['Subject'] = email_subject

    # Add email body
    msg.attach(MIMEText(email_body, 'plain'))

    try:
        # Connect to the SMTP server and send the email
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(smtp_username, email_address, msg.as_string())
        server.quit()
        log_and_print(f"Email sent to {email_address} (Row {row_number})")
    except Exception as e:
        log_and_print(f"Failed to send email to {email_address} (Row {row_number}): {str(e)}", level=logging.ERROR)
