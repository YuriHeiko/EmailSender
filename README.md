# User Manual: Google Sheets Email Automation Script

## Table of Contents
1. [Introduction](#1-introduction)
2. [Setup and Prerequisites](#2-setup-and-prerequisites)
3. [Configuration](#3-configuration)
   - 3.1. [Setting up Google Sheets Credentials](#31-setting-up-google-sheets-credentials)
   - 3.2. [Creating the Configuration File](#32-creating-the-configuration-file)
4. [Running the Script](#4-running-the-script)
   - 4.1. [Using Bash (Linux/Mac)](#41-using-bash-linuxmac)
   - 4.2. [Using Batch Script (Windows)](#42-using-batch-script-windows)
5. [Log Management](#5-log-management)
6. [Troubleshooting](#6-troubleshooting)
7. [Advanced Customization](#7-advanced-customization)

---

## 1. Introduction

The Google Sheets Email Automation Script is a Python script designed to automate the process of sending emails to a list of recipients stored in a Google Sheet. It also checks for undelivered emails, logs activities, and manages log retention.

**Key features of the script:**
- Send emails using your Gmail account.
- Mark undelivered emails in the Google Sheet.
- Maintain log files with a retention period (default: last 7 days).
- Customize email content and configuration.

## 2. Setup and Prerequisites

Before using the script, ensure you have the following prerequisites in place:

- Python 3.x installed on your system.
- Required Python libraries installed (use `pip install -r requirements.txt` to install them).
- Google Sheets API credentials (JSON file) for authentication.
- A Gmail account for sending emails.
- Access to the Google Sheet you want to work with.

## 3. Configuration

### 3.1. Setting up Google Sheets Credentials

To set up Google Sheets API credentials, follow these steps:

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).

2. Create a new project or select an existing one.

3. Enable the Google Sheets API for your project.

4. Create **service account** credentials for the API with the following details:
   - **API key:** Choose "Other non-UI (e.g. cron job, daemon)" as the application type.
   - Download the credentials as a JSON file.
   - share the sheet with the service account email

5. Place the downloaded JSON file in the script's directory and rename it to `YOUR_CREDENTIALS.json`.

### 3.2. Creating the Configuration File

Create a configuration file named `config.txt` in the script's directory with the following sections and settings:

```plaintext
[SMTP]
server: smtp.gmail.com
port: 587
username: your@gmail.com
app_password: your_password

[Email]
subject: Your email subject
body: This is the first line of the email body.
      This is the second line.
      You can continue on additional lines as well.
attachment_path: path to the attachment
attachment_type: type of the attachment
      
[GoogleSheet]
start_row: 518
email_column: 5 # Column E
undelivered_column: 7 # Column G
sheet_name: Your Sheet Name  # Add this line to specify the sheet name
send_text: sent
undelivered_text: undelivered
```
[SMTP] section: Configure your SMTP server details for sending emails.

[Email] section: Set your email subject, body and attachment path.
You might need to generate an App password to grant access to your mailbox

Generate an App Password:
Once two-factor authentication is enabled, you can generate an app-specific password for your script. Here's how:

a. Go to the "Security" section of your Google Account settings.

b. Under the "2-Step Verification" section, click on "App passwords."

c. Enter app name

g. Click the "Create" button.

h. You'll be presented with a 16-character app password. Copy this password.


[GoogleSheet] section: Specify the starting row, email column, undelivered column, and the name of the Google Sheet.


## 4. Running the Script

To run the script, follow these instructions:

1. Make sure you have activated your virtual environment (if using one) and installed the required Python packages from `requirements.txt`.

2. Open a terminal or command prompt in the script's directory.

3. Run the Python script using the following command:

   On Linux/Mac:
   ```bash
   ./run_script.sh
   ```
   On Windows:
   ```batch
   run_script.bat
   ```
## 5. Log Management

The script maintains log files in the `log.txt` file. Log entries older than the specified retention period (default: 7 days) are automatically pruned.

### Manual Log Cleanup

You can manually remove log entries by opening the `log.txt` file and deleting the entries you no longer need.

## 6. Troubleshooting

If you encounter any issues or errors while running the script, refer to the following troubleshooting steps:

1. **Check Credentials:** Ensure that the `YOUR_CREDENTIALS.json` file exists and contains valid Google Sheets API credentials.

2. **SMTP Configuration:** Verify the SMTP server settings in the `config.txt` file. Make sure they are correct and allow sending emails through your Gmail account.

3. **Logging:** Check the `log.txt` file for error messages and log details. It provides information about script activities and potential issues.

## 7. Advanced Customization

- You can customize the log retention period by modifying the `clean_up_old_logs` function in the script.
- Adjust the log levels and format in the script's logging configuration for more detailed or concise logging.
- Customize the email content in the `config.txt` file to match your specific email requirements.

For more advanced customizations, you can modify the script according to your needs or consult Python documentation for related libraries (e.g., logging, smtplib, imaplib).

---

By following this user manual, you can effectively use the Google Sheets Email Automation Script to automate email communication with recipients stored in a Google Sheet while managing logs and log retention.
