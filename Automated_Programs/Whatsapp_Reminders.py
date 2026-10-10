import pandas as pd
import time
import random
import datetime
import os
import pyautogui

def run_logged_excel_blast():
    #ENTER CORRECT FILE PATH
    tenants_file = r"C:\<............>\Automated_Electricity_Billing_Software\Database\tenant_list.xls"
    bills_file = r"C:\<..............>\Automated_Electricity_Billing_Software\Database\monthly_bills.xls"
    log_path = r"C:\<..............>\Automated_Electricity_Billing_Software\Data_Log\Message_Log.txt"

    if not os.path.exists(tenants_file) or not os.path.exists(bills_file):
        print("Error: Excel source files missing from the directory folder.")
        return

    print("Reading text profiles from Excel rows...")
    df_tenants = pd.read_excel(tenants_file, header=None)
    df_bills = pd.read_excel(bills_file, header=None)

    if "name" in str(df_tenants.iloc[0, 0]).lower(): df_tenants = df_tenants.iloc[1:]
    if "flat" in str(df_bills.iloc[0, 0]).lower(): df_bills = df_bills.iloc[1:]

    df_tenants_clean = pd.DataFrame()
    df_tenants_clean['tenant_name'] = df_tenants.iloc[:, 0].astype(str).str.strip()
    df_tenants_clean['phone_number'] = df_tenants.iloc[:, 1].astype(str).str.strip()
    df_tenants_clean['flat_no'] = df_tenants.iloc[:, 2].astype(float).astype(int).astype(str)

    df_bills_clean = pd.DataFrame()
    df_bills_clean['flat_no'] = df_bills.iloc[:, 0].astype(float).astype(int).astype(str)
    df_bills_clean['bill_amount'] = df_bills.iloc[:, 1].astype(float)
    df_bills_clean['due_date'] = df_bills.iloc[:, 2].astype(str).str.strip()
    df_bills_clean['status'] = df_bills.iloc[:, 3].astype(str).str.strip()

    df_unpaid = df_bills_clean[df_bills_clean['status'].str.lower() == 'unpaid']
    df_master = pd.merge(df_unpaid, df_tenants_clean, on='flat_no')

    if df_master.empty:
        print("All records are cleared in your bills file sheet ledger!")
        return

    print(f"Found {len(df_master)} unpaid records. Initializing layout...")
    input("Press [ENTER] to start loop tracking... ")
    print("\nStarting in 5 seconds... Click inside your active WhatsApp Web tab now!")
    time.sleep(5)

    #ENTER CORRECT AND VALID UPI ID AND NAME
    MY_UPI_ID = "<......>@oksbi"  
    MY_NAME = "<.......>"

    for index, row in df_master.iterrows():
        name = str(row['tenant_name'])
        phone = str(row['phone_number'])
        flat = str(row['flat_no'])
        amount = float(row['bill_amount'])
        due_date = str(row['due_date']).split()[0]

        clean_phone = phone.replace("+", "").replace(" ", "").replace("-", "")

        message = (
            f"Dear {name},\n\n"
            f"This is an automated reminder for your electricity bill for *Flat {flat}*.\n"
            f" - Amount Due: *₹{amount}*\n"
            f" - Due Date: *{due_date}*\n\n"
            f"Please click the link below to open GPay/UPI, and pay the bill:\n"
            f"upi://pay?pa={MY_UPI_ID}&pn={MY_NAME}&am={amount}&cu=INR&tn=Flat{flat}Electricity\n\n"
            f"Thank you!"
        )

        print(f"[{index + 1}/{len(df_master)}] Broadcasting to Flat {flat}...")

        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        time.sleep(0.5)
        pyautogui.write(clean_phone, interval=0.1)
        pyautogui.press('enter')
        time.sleep(4)

        lines = message.split('\n')
        for i, line in enumerate(lines):
            if line:  
                pyautogui.write(line, interval=0.01)
            if i < len(lines) - 1:
                pyautogui.hotkey('shift', 'enter')
        
        time.sleep(1)
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl', 'f')
        print(f"Message sent successfully to Flat {flat}!")

        # UNIFIED COMPACT LOGGING ENGINE
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as file:
            file.write(f"TIMESTAMP: {current_time} | BILL MSG SENT | Flat: {flat} | Tenant: {name} | Amount: ₹{amount}\n")

        cooldown = random.randint(8, 10)
        print(f"Safe Pause: Waiting {cooldown} seconds...")
        time.sleep(cooldown)

    print(f"\nProcess Finished! Sent logs appended inside payment_log.txt.")

if __name__ == "__main__":
    run_logged_excel_blast()
