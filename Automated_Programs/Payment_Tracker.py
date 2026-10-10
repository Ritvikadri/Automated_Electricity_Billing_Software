import imaplib
import email
import re
import time
import os
import datetime
import xlrd
from xlutils.copy import copy

# --- GMAIL & BANK SECURITY CONFIGURATION ---# --- GMAIL & BANK SECURITY CONFIGURATION ---
EMAIL_USER = "<ENTER YOUR GMAIL>"        # Your Gmail address
EMAIL_PASS = "ENTER THE APP PASSWORD CREATED IN GOOGLE MANAGEMENT SECTION OR CREATE ONE IT CONTAINS 16 ALPHABETS"       # Your 16-character Google App Password
IMAP_SERVER = "imap.gmail.com"             # Google IMAP server
BANK_EMAIL = "<OFFICIAL BANKS EMAIL, EX:cbsalerts.sbi@alerts.sbi.bank.in>"   # SBI official credit alert emailprint(f"⚠️ Warning: Received ₹{received_amount}, but no matching unpaid row layout exists inside Excel.")
# -------------------------------------------

def scan_emails_and_update_bills_file():
    bills_file = r"C:\<..........>\Automated_Electricity_Billing_Software\Database\monthly_bills.xls"
    payment_log_file = r"C:\<........>\Automated_Electricity_Billing_Software\Data_Log\Payments_Log.txt"
    
    if not os.path.exists(bills_file):
        print(f"Error: Cannot find 'monthly_bills.xls' at {bills_file}")
        return

    try:
        # 1. Establish secure connection to your Gmail inbox
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")

        # 2. Query only UNREAD (UNSEEN) alerts coming from SBI
        status, response = mail.search(None, f'(UNSEEN FROM "{BANK_EMAIL}")')
        
        if response and response[0]:
            email_ids = response[0].split()
        else:
            email_ids = []

        if not email_ids:
            mail.logout()
            return

        print(f"[{time.strftime('%X')}] Found {len(email_ids)} new unread SBI transaction alert(s). Processing...")

        for e_id in email_ids:
            status, msg_data = mail.fetch(e_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    raw_email_bytes = response_part[1]  # Fix to target the byte contents inside the tuple pair
                    msg = email.message_from_bytes(raw_email_bytes)
                    
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                payload = part.get_payload(decode=True)
                                if payload:
                                    body = payload.decode(errors='ignore')
                    else:
                        payload = msg.get_payload(decode=True)
                        if payload:
                            body = payload.decode(errors='ignore')

                    # 3. Read the credited money amount from text
                    amount_match = re.search(r'(?:Rs\.?|INR|₹)\s*([\d,]+\.?\d*)', body, re.IGNORECASE)
                    
                    if amount_match:
                        received_amount = float(amount_match.group(1).replace(',', ''))
                        print(f"Parsed credit transaction verification for: ₹{received_amount}")

                        # 4. Open and make a writable copy of the old .xls sheet
                        rb = xlrd.open_workbook(bills_file, formatting_info=True)
                        sheet_read = rb.sheet_by_index(0) 
                        
                        wb = copy(rb)
                        sheet_write = wb.get_sheet(0)
                        
                        match_found = False
                        
                        for row in range(1, sheet_read.nrows):
                            excel_amount = float(sheet_read.cell_value(row, 1))
                            excel_status = str(sheet_read.cell_value(row, 3)).strip()
                            
                            if excel_amount == received_amount and excel_status.lower() == "unpaid":
                                sheet_write.write(row, 3, "Paid")
                                wb.save(bills_file) 
                                
                                flat_no = sheet_read.cell_value(row, 0)
                                if isinstance(flat_no, float):
                                    flat_no = int(flat_no)
                                    
                                print(f"EXCEL AUTOMATION SUCCESS: Flat {flat_no} auto-marked as PAID inside monthly_bills.xls.")
                                
                                # RECEIVED PAYMENTS RECORD WRITER
                                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                with open(payment_log_file, "a", encoding="utf-8") as file:
                                    file.write(f"==================================================\n")
                                    file.write(f"PAYMENT RECEIVED ON : {current_time}\n")
                                    file.write(f"FROM FLAT NUMBER    : {flat_no}\n")
                                    file.write(f"AMOUNT COLLECTED    : ₹{received_amount}\n")
                                    file.write(f"STATUS IN SYSTEM    : SUCCESSFUL - EXCEL UPDATED\n")
                                    file.write(f"==================================================\n\n")
                                print(f"Receipt safely logged into received_payments_log.txt")
                                
                                mail.store(e_id, '+FLAGS', '\\Seen')
                                match_found = True
                                break
                        
                        if not match_found:
                            print(f"Warning: Received ₹{received_amount}, but no matching unpaid row exists inside your bills file.")
                    else:
                        print("Error: Script failed to locate currency parameters inside this email text.")

        mail.logout()

    except Exception as e:
        # FIXED: This will now print the EXACT error message string so we can target the exact fix line
        print(f"Spreadsheet Background Sync Engine Error: [{str(e)}]")

if __name__ == "__main__":
    print("=========================================")
    print("  EXCEL EMAIL TRANSACTION ENGINE ACTIVE  ")
    print("=========================================")
    while True:
        scan_emails_and_update_bills_file()
        print("Sleeping for 1 minute before scanning inbox again...\n")
        time.sleep(60)
