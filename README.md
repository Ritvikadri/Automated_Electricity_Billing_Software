AUTOMATED ELECTRICITY BILLING SOFTWARE
=====================================
Automated Electricity Billing Software is a Python project for managing electricity bills in Excel workbooks. It can send automated WhatsApp payment reminders and receipts, track payments from bank notification emails, and report outstanding balances.



IMPORTANT — BEFORE USING
-------------------------------------
The scripts contain hard-coded paths and personal phone/UPI settings that must be changed for your computer and account. Payment_Tracker.py also contains an email app password in its source, which you have to create from the google app passwords and enter there. Then configure your own credentials securely before publishing or running the tracker. Never publish passwords, tenant records, or real phone numbers.



WHAT EACH SCRIPT DOES
---------------------
Whatsapp_Reminders.py
  Finds bills marked Unpaid, matches them to tenants by flat number, and sends a WhatsApp reminder with the amount, due date, and a UPI payment link.

Unpaid_Reminders.py
  Also sends reminders for unpaid bills, with a different message format. Run only one reminder script for a billing cycle to avoid duplicate messages.

Payment_Tracker.py
  Checks unread emails from a configured bank sender and marks the first unpaid bill with an exactly matching amount as Paid. It checks once per minute until stopped. It        matches by amount alone, so it can update the wrong flat when amounts repeat. So never keep the amount the same for any of the tenant. Verify every update manually for clarity.

Payment_Confirmation_Message.py
  Finds bills marked Paid whose receipt_sent value is not Yes, sends a receipt, and marks the row Yes.

Final_Bill_Report.py
  Sends a summary to the configured administrator number with unpaid flats and the outstanding total, or an all-paid message.



REQUIREMENTS
------------
- Windows 10 or 11.
- Python 3.10 or later, available as python in Command Prompt or PowerShell.
- Excel or another application that can edit legacy .xls workbooks.
- Google Chrome and a signed-in WhatsApp Web session for the WhatsApp scripts.
- Gmail with IMAP enabled and a Google App Password, if using Payment_Tracker.py.



DOWNLOAD AND INSTALL
--------------------
Open PowerShell in the extracted project folder by using cd <file path>, where the Automated_Programs, Database, and Data_Log folders are located.

Create a virtual environment if required.

Install the required libraries from powershell:

    python -m pip install --upgrade pip
    python -m pip install pandas pyautogui xlrd==2.0.2 xlutils xlwt
                         OR
    pip install -r requirements.txt

The modules time, datetime, os, random, imaplib, email, and re come with Python. xlrd reads .xls files, while xlutils and xlwt support updating and saving them. Renaming an .xlsx file to .xls does not convert its format.



PROJECT FOLDER AND DATA SETUP
-----------------------------
Keep the folders together in this structure:

    project-folder/
    |-- Automated_Programs/
    |   |-- Final_Bill_Report.py
    |   |-- Payment_Confirmation_Message.py
    |   |-- Payment_Tracker.py
    |   |-- Unpaid_Reminders.py
    |   `-- Whatsapp_Reminders.py
    |-- Database/
    |   |-- tenant_list.xls
    |   `-- monthly_bills.xls
    `-- Data_Log/
        |-- Message_Log.txt
        `-- Payments_Log.txt

The scripts currently refer to the original author's Windows folder using absolute paths. They will not work on another computer until you update those paths in the scripts to point to your extracted project folder. Keep the folder names above. If the log files are missing, create empty files with those names.

Use the first worksheet in each workbook and include a header row.

Database/tenant_list.xls

    Column A: tenant_name or name — tenant's name
    Column B: phone_number — WhatsApp number, including country code, such as +91...
    Column C: flat_number — flat number that matches monthly_bills.xls

Database/monthly_bills.xls

    Column A: flat_no — flat number
    Column B: amount — numeric bill amount
    Column C: due_date — due date
    Column D: status — Paid or Unpaid
    Column E: receipt_sent — set to Yes after a receipt is sent; leave blank or use No if a paid bill needs a receipt

Flat numbers must match in both workbooks. Enter amounts as numbers, without a currency symbol, and use Paid or Unpaid for bill status. Back up the workbooks before running scripts that update them.



CONFIGURE YOUR DETAILS
----------------------
Before running scripts, review the code and replace the original author's settings:

- In Final_Bill_Report.py, set ADMIN_PHONE to your administrator's WhatsApp number with country code.
- In Whatsapp_Reminders.py and Unpaid_Reminders.py, set MY_UPI_ID and MY_NAME to the intended payment recipient. Check the payment link and message text carefully.
- In Payment_Tracker.py, configure your Gmail address and the bank notification sender. Put a regular Gmail password in the script created by you. Enable IMAP, and revoke any credential that has been exposed.
- Review tenant names, phone numbers, amounts, due dates, message text, and recipients. Try the scripts with dummy data and your own number first.



RUN A SCRIPT
------------
1. Close tenant_list.xls and monthly_bills.xls in Excel or any other program. Other programs can also lock the files.
2. For a WhatsApp task, open Chrome and sign in to https://web.whatsapp.com. Wait until the page has fully loaded, and keep Chrome visible and maximized.
3. In PowerShell, change into the scripts folder:

       cd <project-folder>\Automated_Programs

4. Run the script you need. For example:

       python Whatsapp_Reminders.py

   Replace the filename with the script you want to run. Payment_Tracker.py does not use WhatsApp; it keeps running until you stop it with Ctrl+C.

5. When a WhatsApp script prompts you, press Enter, then immediately click the WhatsApp Web tab. The script waits 5 seconds before sending keystrokes. WhatsApp must already be open and signed in; the script does not launch it within those five seconds.
6. Keep the computer awake and avoid using the mouse or keyboard until the script finishes. The scripts search for phone numbers, type messages, and press Enter to send. Tenant messaging scripts pause 8–10 seconds between messages.

Run only one WhatsApp automation at a time. The scripts type into whichever window has focus, so switching windows can send text to the wrong place. If that happens, stop with Ctrl+C, then check WhatsApp and the workbooks before restarting.



LOGS AND WORKBOOK CHANGES
-------------------------
- Data_Log/Message_Log.txt stores reminder, receipt, and administrator-summary entries.
- Data_Log/Payments_Log.txt stores payment updates made by Payment_Tracker.py.
- Payment_Confirmation_Message.py updates column E (receipt_sent) to Yes.
- Payment_Tracker.py changes the first unpaid row with an exact amount match to Paid.

A terminal message saying a WhatsApp message was sent only means the script issued the keyboard actions. It cannot confirm delivery or reading. Check WhatsApp directly. Verify payment updates against bank records, since the tracker may select the wrong flat when multiple unpaid bills have the same amount.



TROUBLESHOOTING
---------------
Problem: “File missing” or FileNotFoundError
Check: Update the hard-coded paths. Check folder and file names, extensions, and that both workbooks are in Database.

Problem: PermissionError or workbook cannot save
Check: Close both .xls workbooks in Excel or LibreOffice. Close other applications that may have the files open, including File Explorer preview panes. Keep the workbooks closed during the run.

Problem: ModuleNotFoundError or ImportError
Check: Activate the virtual environment, then reinstall the required packages using that environment's python -m pip.

Problem: Unsupported or invalid Excel format
Check: Confirm the files are actual .xls workbooks, not .xlsx files renamed to .xls. Confirm xlrd==2.0.2, xlutils, and xlwt installed successfully.

Problem: WhatsApp search cannot find a tenant
Check: Make sure WhatsApp Web is signed in and the number in column B includes the country code and belongs to an active WhatsApp account.

Problem: Message goes to the wrong window or is incomplete
Check: Keep WhatsApp visible and focused during the 5-second countdown. Avoid other mouse or keyboard use, and turn Caps Lock off.

Problem: No reminders are sent
Check: Confirm column D says Unpaid and that flat numbers match between the workbooks.

Problem: No payment receipt is sent
Check: Confirm the bill status is Paid, column E is not Yes, the tenant flat matches, and the bills workbook has at least five columns.

Problem: Payment tracker cannot log in
Check: Enable IMAP, check the Gmail App Password and bank sender address, and store credentials securely. Never post credentials in a public issue or commit.

Problem: Payment is assigned to the wrong flat
Check: The tracker matches by amount alone and picks the first unpaid row. Stop it and verify/correct the ledger using the bank transaction.

PRIVACY AND SAFE USE
--------------------
These scripts handle tenant contact details, bills, payment notifications, and WhatsApp messages. Keep real tenant data, logs, and credentials private. Review the recipient and UPI details before sending, obtain any required consent before contacting tenants, and independently verify payment records.
