# User Manual: Google Sheets Email Automation Script

## Table of Contents
1. [Introduction](#1-introduction)
2. [Setup and Prerequisites](#2-setup-and-prerequisites)
3. [Configuration](#3-configuration)
   - 3.1. [Setting up Google Sheets Credentials](#31-setting-up-google-sheets-credentials)
   - 3.2. [Creating the Configuration File](#32-creating-the-configuration-file)
4. [Running the Script](#4-running-the-script)
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

4. Create credentials for the API with the following details:
   - **API key:** Choose "Other non-UI (e.g. cron job, daemon)" as the application type.
   - Download the credentials as a JSON file.

5. Place the downloaded JSON file in the script's directory and rename it to `YOUR_CREDENTIALS.json`.

### 3.2. Creating the Configuration File

Create a configuration file named `config.txt` in the script's directory with the following sections and settings:

```plaintext
[SMTP]
server: smtp.gmail.com
port: 587
username: your@gmail.com
password: your_password

[Email]
subject: Your email subject
body: Your email body

[GoogleSheet]
start_row: 518
email_column: 5
undelivered_column: 7
sheet_name: Your Sheet Name  # Add this line to specify the sheet name
