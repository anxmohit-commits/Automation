import pandas as pd
import os
import glob
import win32com.client as win32
from datetime import datetime, timedelta

# --- Configuration Paths ---
input_folder = r"D:\Automation\Rigger_Attendance\Upload"
output_folder = r"D:\Automation\Rigger_Attendance\HTIS_Pvt_Ltd_Agency\output"
email_config_path = r"D:\Automation\Rigger_Attendance\HTIS_Pvt_Ltd_Agency\Email_Config.xlsx"

# Dedicated Sender Email Address
FROM_EMAIL = "IN_D_Engineering_NIfi_Reports@airtel.com"

# 1. Date Logic: Get yesterday's date for report naming
yesterday = datetime.now() - timedelta(days=1)
date_str = yesterday.strftime("%d-%b-%Y").upper()

# 2. Logic to find the latest file in the folder
def get_latest_file(folder_path):
    files = glob.glob(os.path.join(folder_path, "*.*"))
    if not files:
        return None
    # Sort files by modification time
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

# 3. Email Sending Logic using Outlook Integration
def send_outlook_email(attachment_path, to_list, cc_list):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        # Use SentOnBehalfOfName for the shared mailbox/specific sender
        mail.SentOnBehalfOfName = FROM_EMAIL
            
        mail.To = "; ".join(to_list)
        mail.CC = "; ".join(cc_list)
        mail.Subject = f"Rigger_Attendance Report - HTIS_Pvt_Ltd Agency-{date_str}"
        
        # Updated HTMLBody with specific bold segments
        mail.HTMLBody = f"""
        <html>
        <body>
            <p>Dear Team,</p>
            <p>Please find attached <b>Rigger_Attendance Report</b> for <b>HTIS_Pvt_Ltd Agency</b>.</p>
            <p><b style="color:red;">This is an auto-generated mail.</b></p>
            <p>Regards,<br>
            Central FSO Automation Team</p>
        </body>
        </html>
        """

        mail.Attachments.Add(attachment_path)
        mail.Send()
        print(f"✅ Mail sent successfully from: {FROM_EMAIL}")
    except Exception as e:
        print(f"❌ Outlook Error: {e}")

# 4. Main Execution Logic
def main():
    try:
        # Step A: Locate the latest input file
        latest_input = get_latest_file(input_folder)
        if not latest_input:
            print("⚠️ No files found in the upload folder!")
            return
        
        print(f"Processing latest file: {os.path.basename(latest_input)}")

        # Step B: Read file based on extension (CSV or Excel)
        ext = os.path.splitext(latest_input)[1].lower()
        df = pd.read_csv(latest_input, encoding='latin1') if ext == '.csv' else pd.read_excel(latest_input)
        
        # Step C: Filter data for the specific agency
        filtered_df = df[df['rigger_agency_organisation'].astype(str).str.strip() == 'HTIS Pvt Ltd']

        if not filtered_df.empty:
            # Create output directory if it doesn't exist
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
                
            # Changed extension to .csv
            output_file_name = f"HTIS_Pvt_Ltd Agency_{date_str}.csv"
            full_output_path = os.path.join(output_folder, output_file_name)

            # Step D: Save filtered data to CSV
            filtered_df.to_csv(full_output_path, index=False)
            
            # Step E: Fetch Email recipients from Configuration file
            email_df = pd.read_excel(email_config_path)
            to_list = email_df[email_df['Recipient_Type'].str.upper() == 'TO']['Email_ID'].dropna().tolist()
            cc_list = email_df[email_df['Recipient_Type'].str.upper() == 'CC']['Email_ID'].dropna().tolist()

            # Step F: Send the email if recipients are found
            if to_list:
                send_outlook_email(full_output_path, to_list, cc_list)
            else:
                print("⚠️ No recipients found in Email_Config.xlsx.")
        else:
            print("⚠️ No data found for HTIS Pvt Ltd.")

    except Exception as e:
        print(f"⚠️ Critical Error: {e}")

if __name__ == "__main__":
    main()