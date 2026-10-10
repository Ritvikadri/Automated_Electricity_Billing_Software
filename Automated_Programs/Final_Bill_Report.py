import pandas as pd
import time
import datetime
import os
import pyautogui

def run_unpaid_summary_blast():
    # Folder paths setup
    bills_file = r"C:\<..........>\Automated_Electricity_Billing_Software\Database\monthly_bills.xls"
    log_path = r"C:<...........>\Automated_Electricity_Billing_Software\Data_Log\Message_Log.txt"

    #CONFIGURATION BOX: Put your own mobile number here (with country code, no spaces)
    ADMIN_PHONE = "+91<OWNERS PHONE NUMBER>" 
    
    if not os.path.exists(bills_file):
        print("Error: 'monthly_bills.xls' missing from the directory folder.")
        return

    print("====================================================")
    print("      UNPAID TENANTS SUMMARY REPORT GENERATOR       ")
    print("====================================================")
    print("Reading spreadsheet records...")
    
    # Read files completely ignoring header naming mismatch bugs
    df_bills = pd.read_excel(bills_file, header=None)

    # Clean out top header rows if present
    if "flat" in str(df_bills.iloc[0, 0]).lower(): 
        df_bills = df_bills.iloc[1:]

    # Parse and structure the bills details cleanly: 0=Flat, 1=Amount, 2=Due Date, 3=Status
    df_bills_clean = pd.DataFrame()
    df_bills_clean['flat_no'] = df_bills.iloc[:, 0].astype(float).astype(int).astype(str)
    df_bills_clean['bill_amount'] = df_bills.iloc[:, 1].astype(float)
    df_bills_clean['status'] = df_bills.iloc[:, 3].astype(str).str.strip()

    # Isolate strictly outstanding accounts that are still UNPAID
    df_unpaid = df_bills_clean[df_bills_clean['status'].str.lower() == 'unpaid']

    current_date = datetime.datetime.now().strftime("%d-%m-%Y")
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Setup the message structure and compact log tracking flags
    if df_unpaid.empty:
        print("Clean Ledger! Generating all-paid confirmation message...")
        message = (
            f"*ELECTRICITY BILL REPORT* \n"
            f"Date: {current_date}\n\n"
            f"*Excellent News! All tenants have successfully paid their electricity bills for this month.*\n\n"
            f"*Total Outstanding Balance: Rs. 0.00*\n\n"
            f"This is an automated system generation log report."
        )
        # Compact single-line trace string for an all-paid event
        log_entry = f"TIMESTAMP: {current_time} | ELEC REPORT SENT   | Status: ALL PAID | Outstanding: Rs. 0.00\n"
    else:
        # Compile the regular delinquent metrics calculations
        total_unpaid_amount = df_unpaid['bill_amount'].sum()
        total_unpaid_flats = len(df_unpaid)

        # Build the list of delinquent flats line by line using "Rs." text formatting
        unpaid_flats_list_text = ""
        for index, row in df_unpaid.iterrows():
            unpaid_flats_list_text += f" ▪️ Flat {row['flat_no']}: Rs. {row['bill_amount']}\n"

        message = (
            f"*ELECTRICITY BILL REPORT*\n"
            f"Date: {current_date}\n\n"
            f"The following {total_unpaid_flats} flats have not paid their electricity bills yet:\n\n"
            f"{unpaid_flats_list_text}\n"
            f"*Total Outstanding Balance: Rs. {total_unpaid_amount:.2f}*\n\n"
            f"This is an automated system generation log report."
        )
        # Compact single-line trace string for an outstanding dues event
        log_entry = f"TIMESTAMP: {current_time} | ELEC REPORT SENT   | Status: UNPAID DUES | Outstanding: Rs. {total_unpaid_amount:.2f} ({total_unpaid_flats} flats)\n"

    print("\n STEPS FOR DISPATCH:")
    print("1. Ensure your Chrome browser with WhatsApp Web is open and fully maximized.")
    print("2. Make sure CAPS LOCK IS TURNED OFF on your physical keyboard!")
    input("3. Press [ENTER] here and IMMEDIATELY click onto your open WhatsApp tab... ")

    print("\nStarting in 5 seconds... Click inside your open WhatsApp tab now!")
    time.sleep(5)

    clean_phone = ADMIN_PHONE.replace("+", "").replace(" ", "").replace("-", "")

    # WhatsApp Search Focus Execution
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    time.sleep(0.5)
    
    # Types the administrator mobile number with exactly 0.1 seconds gap between digits
    pyautogui.write(clean_phone, interval=0.1)
    time.sleep(5)
    pyautogui.press('enter')
    time.sleep(2)

    # High-speed message line writer engine (types at 0.01 speed using Shift+Enter)
    lines = message.split('\n')
    for i, line in enumerate(lines):
        if line:  
            pyautogui.write(line, interval=0.01)
        if i < len(lines) - 1:
            pyautogui.hotkey('shift', 'enter')
    
    time.sleep(1)
    pyautogui.press('enter') # Send the complete message bubble
    pyautogui.hotkey('ctrl', 'f')
    print(f"Electricity report successfully typed and sent to {ADMIN_PHONE}!")

    # UNIFIED COMPACT LOG RECORD WRITER (Appends directly to your clean tracker lines)
    with open(log_path, "a", encoding="utf-8") as file:
        file.write(log_entry)
    print("Report summary trace safely logged in a single line inside Message_Log.txt!")

if __name__ == "__main__":
    run_unpaid_summary_blast()
