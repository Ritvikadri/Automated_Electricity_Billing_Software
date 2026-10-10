import time
import os
import random
import datetime
import xlrd
import pyautogui
from xlutils.copy import copy

def run_new_receipts_blast():
    #ENTER CORRECT FILE PATH
    tenants_file = r"C:\<.........>\Automated_Electricity_Billing_Software\Database\tenant_list.xls"
    bills_file = r"C:\<............>\Automated_Electricity_Billing_Software\Database\monthly_bills.xls"
    log_path = r"C:\<...........>\Automated_Electricity_Billing_Software\Data_Log\Message_Log.txt"

    if not os.path.exists(tenants_file) or not os.path.exists(bills_file):
        print("Error: Make sure both tenant_list.xls and monthly_bills.xls exist in your folder.")
        return

    print("====================================================")
    print("      HIGH-SPEED SMART RECEIPT DISPATCH ENGINE      ")
    print("====================================================")
    print("Reading and cross-referencing spreadsheet records...")

    try:
        rb_bills = xlrd.open_workbook(bills_file, formatting_info=True)
        sheet_bills = rb_bills.sheet_by_index(0)
        
        wb_bills = copy(rb_bills)
        sheet_bills_write = wb_bills.get_sheet(0)

        rb_tenants = xlrd.open_workbook(tenants_file)
        sheet_tenants = rb_tenants.sheet_by_index(0)
    except PermissionError:
        print("\nCRITICAL ERROR: Please close 'monthly_bills.xls' inside Excel or OpenOffice before running this script!")
        return

    if sheet_bills.ncols < 5:
        print("Error: Please open monthly_bills.xls and type 'receipt_sent' as the header in Column E first!")
        return

    targets_to_send = []

    for row in range(1, sheet_bills.nrows):
        raw_flat = sheet_bills.cell_value(row, 0)
        flat_no = str(int(raw_flat)).strip() if isinstance(raw_flat, float) else str(raw_flat).strip()
        
        amount = sheet_bills.cell_value(row, 1)
        status = str(sheet_bills.cell_value(row, 3)).strip().lower()
        receipt_sent = str(sheet_bills.cell_value(row, 4)).strip().lower()

        if status == "paid" and receipt_sent != "yes":
            tenant_name = "Tenant"
            phone_number = None

            for t_row in range(1, sheet_tenants.nrows):
                t_flat_raw = sheet_tenants.cell_value(t_row, 2)
                t_flat = str(int(t_flat_raw)).strip() if isinstance(t_flat_raw, float) else str(t_flat_raw).strip()
                
                if t_flat == flat_no:
                    tenant_name = str(sheet_tenants.cell_value(t_row, 0)).strip()
                    phone_number = str(sheet_tenants.cell_value(t_row, 1)).strip()
                    break

            if phone_number:
                targets_to_send.append({
                    "row_index": row,
                    "flat_no": flat_no,
                    "amount": amount,
                    "name": tenant_name,
                    "phone": phone_number
                })

    if not targets_to_send:
        print("Clean Slate! No new paid tenants require a receipt message today.")
        return

    print(f"Identified {len(targets_to_send)} newly paid flat(s) waiting for validation notes.")
    print("\nSTEPS FOR RUNNING:")
    print("1. Ensure Google Chrome with WhatsApp Web is open and fully maximized.")
    input("2. Press [ENTER] here and IMMEDIATELY click onto your open WhatsApp tab... ")

    print("\nStarting in 5 seconds... Click inside your open WhatsApp tab now!")
    time.sleep(5)

    # Automated processing loop execution inside a single open tab
    for index, tenant in enumerate(targets_to_send):
        clean_phone = tenant["phone"].replace("+", "").replace(" ", "").replace("-", "")

        # Kept the exact structure and formatting symbols intact
        message_lines = [
            f"Dear {tenant['name']},",
            "",
            "✅ *Payment Received Successfully!*",
            "",
            f"Thank you for clearing your electricity bill for *Flat {tenant['flat_no']}*.",
            f" - Amount Settled: *₹{tenant['amount']}*",
            " - Status: *PAID*",
            "",
            "Your updated digital receipt has been logged. Have a wonderful day!"
        ]

        print(f"[{index + 1}/{len(targets_to_send)}] Dispatching receipt to Flat {tenant['flat_no']} ({tenant['name']})...")

        # Focus WhatsApp search (Ctrl + F)
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        time.sleep(0.5)

        #Types phone number at 0.1 seconds per character
        pyautogui.write(clean_phone, interval=0.1)
        pyautogui.press('enter')
        time.sleep(4)

        #Types out the message line-by-line with human speed intervals
        for line_idx, line in enumerate(message_lines):
            if line:  # Type the content if the line is not empty
                pyautogui.write(line, interval=0.01)
            
            # Press Shift+Enter for all lines except the very last one
            if line_idx < len(message_lines) - 1:
                pyautogui.hotkey('shift', 'enter')
                time.sleep(0.1)

        time.sleep(1)
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl', 'f')

        # Cache the 'Yes' change in memory (Doesn't write to disk until loop finishes)
        sheet_bills_write.write(tenant["row_index"], 4, "Yes")

        # Log entry to notepad file archive
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as file:
            file.write(f"TIMESTAMP: {current_time} | SMART RECEIPT SENT | Flat: {tenant['flat_no']} | Tenant: {tenant['name']} | Amount: ₹{tenant['amount']}\n")

        # Anti-ban human pacing cooling step
        cooldown = random.randint(8, 10)
        print(f"Pacing Defense: Sleeping for {cooldown} seconds...")
        time.sleep(cooldown)

    # THE FINAL SINGLE-TAB DRIVE WRITER (Executes when script finishes to avoid locking loops)
    try:
        wb_bills.save(bills_file)
        print("\n====================================================")
        print("High-Speed Smart Blast Complete! Excel File Saved.")
        print("====================================================\n")
    except PermissionError:
        print("\nError: Message sent but script could not save the changes to Excel because monthly_bills.xls was opened mid-run!")

if __name__ == "__main__":
    run_new_receipts_blast()
