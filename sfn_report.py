import pandas as pd
import numpy as np
import glob
import os
import win32com.client
import win32timezone
import re
import requests
from datetime import datetime, timedelta
from getpass import getpass  # Secure password input
from urllib.parse import unquote, urlparse, parse_qs
import logging
import traceback
from urllib.parse import urlparse, parse_qs, unquote
from pathlib import Path
import sys
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
#================================================================================Fetching Latest Report from outlook========================================================
#Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

#**1️⃣ Define Download Path**
save_path = Path(r"D:\Daily_SFN_Dump")
save_path.mkdir(parents=True, exist_ok=True)

#**2️⃣ Connect to Outlook**
try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Inbox
    target_folder = inbox.Folders("Required to Work")  # Subfolder
    logging.info("✅ Connected to Outlook successfully.")
except Exception as e:
    logging.error(f"❌ Outlook connection error: {e}")
    sys.exit()

#**3️⃣ Get Emails (Last 7 Days, Sorted by Newest)**
messages = target_folder.Items
messages.Sort("[ReceivedTime]", True)
date_filter = datetime.now() - timedelta(days=7)
messages = messages.Restrict(f"[ReceivedTime] >= '{date_filter.strftime('%m/%d/%Y %H:%M %p')}'")

#**4️⃣ Subject Keywords to Match**
subject_keywords = {
    "SFN_report": "SFN_report",
}

latest_emails = {}
today = datetime.today().date()  # Get today's date

for message in messages:
    try:
        subject = message.Subject.strip()
        received_date = message.ReceivedTime.date()  # Extract only the date part

        if received_date == today:
            for subject_keyword, filename_part in subject_keywords.items():
                if subject_keyword in subject:
                    if filename_part in latest_emails:
                        if message.ReceivedTime > latest_emails[filename_part].ReceivedTime:
                            latest_emails[filename_part] = message
                            logging.info(f"✅ Updated latest email for: {filename_part} → {subject}")
                    else:
                        latest_emails[filename_part] = message
                        logging.info(f"✅ Found latest email for: {filename_part} → {subject}")

    except Exception as e:
        logging.warning(f"⚠️ Error processing email: {e}")

#**5️⃣ Check if All Required Emails are Found**
required_files = {"SFN_report"}

if not required_files.issubset(set(latest_emails.keys())):
    logging.error("❌ Not all required emails were found. Exiting...")
    sys.exit()

logging.info("✅ All required emails found. Proceeding with download...")

#**6️⃣ Ask for Credentials**
# USERNAME = input("Enter your username: ").strip()
# PASSWORD = getpass("Enter your password: ").strip()


credentials = pd.read_excel(r"D:\Automation\credentials.xlsx")
USERNAME = credentials.iloc[0, 0]  # First row, first column (username)
PASSWORD = credentials.iloc[0, 1]  # First row, second column (password)
print("Credentials loaded successfully.")





# USERNAME = "B0321755"
# PASSWORD = "Bihar@1949"

#**7️⃣ Function to Extract Real URL from SafeLinks**
def extract_real_url(safelink):
    """Extract real URL from Outlook SafeLink"""
    if "safelinks.protection.outlook.com" in safelink:
        parsed_url = urlparse(safelink)
        query_params = parse_qs(parsed_url.query)
        if "url" in query_params:
            return unquote(query_params["url"][0])
    return safelink

#**8️⃣ Extract Download Links & Download Files**
for filename_part, latest_email in latest_emails.items():
    try:
        email_body = latest_email.Body
        match = re.search(r"Link\s*[:\-]?\s*<?(https?://[^\s<>\"']+)>?", email_body)

        if not match:
            logging.warning(f"⚠️ No valid download link found in email: {filename_part}")
            continue

        download_url = extract_real_url(match.group(1))  # Extract real URL from SafeLink if necessary
        logging.info(f"🔗 Downloading from: {download_url}")

 #**🔄 Retry logic for server errors (500)**
        retry_count = 3
        for attempt in range(retry_count):
            try:
                response = requests.get(download_url, auth=(USERNAME, PASSWORD), allow_redirects=True)

                if response.status_code == 200:
                    today_date = datetime.now().strftime('%Y%m%d')
                    file_name = f"Mobility_{filename_part}_{today_date}.csv"
                    file_path = save_path / file_name

                    with open(file_path, "wb") as file:
                        file.write(response.content)

                    logging.info(f"📂 File downloaded successfully: {file_path}")
                    break  # Exit retry loop after success

                elif response.status_code == 401:
                    logging.error(f"❌ Unauthorized access (401) for {filename_part}. Check credentials or token expiry.")
                    break  # No point retrying

                elif response.status_code == 500:
                    logging.warning(f"⚠️ Server error (500) for {filename_part}, retrying... ({attempt+1}/{retry_count})")

                else:
                    logging.error(f"❌ Failed to download {filename_part}, Status Code: {response.status_code}")
                    break  # Stop retries for other unexpected errors

            except Exception as e:
                logging.error(f"❌ Error downloading {filename_part}: {traceback.format_exc()}")

        else:
            logging.error(f"❌ Failed to download {filename_part} after {retry_count} retries.")

    except Exception as e:
        logging.error(f"❌ Error processing email {filename_part}: {traceback.format_exc()}")



#================================================================================Reading and Cleaning data========================================================================================
# Define the list of circles (excluding 'Total' for now)
circle_list = [
    "AP", "AS","BR", "CN", "DL", "GJ", "HP", "HR", "JK", "KK", "KL", "KO",
    "MH", "MP", "MU", "NE", "OR", "PB", "RJ", "TN", "UE", "UW", "WB","Total"
]

# Create the DataFrame with initial Count = 0
df_circles = pd.DataFrame(circle_list)

df_circles.columns = ["Circle"]



# Folder path containing the files
folder_path = "D:/Daily_SFN_Dump"

# Get today's date in YYYYMMDD format
today_date = datetime.now().strftime('%Y%m%d')

# Get all .csv files matching today's date
file_paths = glob.glob(os.path.join(folder_path, f"*{today_date}*.csv"))

if file_paths:
    file_paths.sort(key=os.path.getmtime, reverse=True)
    
    latest_file = file_paths[0]  # Pick the latest one
    df_sfn = pd.read_csv(latest_file)
    print(f"Loaded file: {latest_file}")
    
    df_sfn = df_sfn[df_sfn['Airtel Circle Name'] != "DYMMYTNG"]
    #df_sfn['Created Date'] = pd.to_datetime(df_sfn['Created Date'], errors='coerce')


    try:
        df_sfn['Created Date'] = pd.to_datetime(df_sfn['Created Date'], format='mixed', dayfirst=True, errors='coerce')
    except Exception as e:
        print(f"Date conversion error: {e}")



    #df_sfn['Created Date'] = pd.to_datetime(df_sfn['Created Date'], format='%m/%d/%Y %H:%M', errors='coerce')

    df_sfn['SFN Created Date'] = df_sfn['Created Date'].dt.normalize()
    
    df_sfn.rename(columns={'Airtel Circle Name': 'Circle'}, inplace=True)

    print(df_sfn.head())
    print(df_sfn.info())

    # Define cutoff range: from 15 days ago to yesterday
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    # Define the cutoff date (15 days before today)
    cutoff_date = today - timedelta(days=15)
    # Filter
    df_sfn_15_days = df_sfn[
    (df_sfn['SFN Created Date'] >= cutoff_date) &
    (df_sfn['SFN Created Date'] < today)
    ]
    
    # print(df_sfn_15_days.head())
    # print(df_sfn_15_days.info())



else:
    print("No file found for today.")

#=========================================================================Creating Summary================================================================
# Create Pivot Table
daily_summary = df_sfn_15_days.pivot_table(
    index='Circle', 
    columns='SFN Created Date', 
    values='User OLM ID',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

#daily_summary.rename(columns={'Total': 'Total'}, inplace=True)
# Format the column headers to show dates in 'DD MMM' format
daily_summary.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in daily_summary.columns
]


daily_circlewise_summary = df_circles.merge(daily_summary, on='Circle', how='left').fillna(0)



# Create Pivot Table
daily_summary_oem = df_sfn_15_days.pivot_table(
    index='Circle', 
    columns='OEM', 
    values='User OLM ID',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()


daily_circlewise_summary_oem = df_circles.merge(daily_summary_oem, on='Circle', how='left').fillna(0)


# Create Pivot Table
daily_summary_card = df_sfn_15_days.pivot_table(
    index='Circle', 
    columns='Product Type', 
    values='User OLM ID',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

daily_circlewise_summary_card_type = df_circles.merge(daily_summary_card, on='Circle', how='left').fillna(0)

#daily_circlewise_summary = df_circles.merge(daily_summary, on='Circle', how='left').fillna(0)

#============================defining for mail body=======================================

def get_top_circles(daily_summary, column):
    # Filter out the "Total" row first
    filtered_data = daily_summary[daily_summary["Circle"] != "Total"]
    
    # Get the top 5 circles based on the specified column
    return ", ".join(filtered_data.nlargest(5, column)["Circle"])

# Finding top 5 circles for each category
top5_circles = get_top_circles(daily_summary, "Total")


circles_with_zero_total = ','.join(daily_circlewise_summary[daily_circlewise_summary['Total'] == 0]['Circle'].values)

# #==========================================================================printing==========================================================================================================
output_folder = r'D:\Daily_SFN_Dashboard'
yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
output_file = os.path.join(output_folder, f'SFN_Dashboard_{yesterday}.xlsx')

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)


# Write to Excel
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:     
    #tt_mapped_df.to_excel(writer, sheet_name="tt_mapped_df", index=False)
    #total_filled_rca_df.to_excel(writer, sheet_name="total_filled_rca_df", index=False)
    daily_circlewise_summary.to_excel(writer, sheet_name="Summary", index=False,startrow = 1,startcol = 0)
    daily_circlewise_summary_oem.to_excel(writer, sheet_name="Summary", index=False,startrow = 1, startcol= 19)
    daily_circlewise_summary_card_type.to_excel(writer, sheet_name="Summary", index=False,startrow = 1, startcol= 31)
    df_sfn_15_days.to_excel(writer, sheet_name="SFN Dump", index=False)
    #df_sfn.to_excel(writer, sheet_name="SFN Dump All", index=False)




    # Apply formatting
    workbook  = writer.book
    worksheet = writer.sheets["Summary"]

    merge_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    # Merge cells with labels
    worksheet.merge_range(0, 0, 0, 16, 'SFN Summary', merge_format)
    worksheet.merge_range(0, 19, 0, 28, 'OEM Summary', merge_format)
    worksheet.merge_range(0, 31, 0, 40, 'Cardwise Summary', merge_format)
      

print(f"\n💾 Excel file saved successfully at: {output_file}")

#=================================sending Mail=====================================

# Read recipients from Excel
file_path_receipient = r"D:\Automation\SFN_Report\Receipient.xlsx"
try:
    receipient_df = pd.read_excel(file_path_receipient)
    to_recipients = receipient_df["to"].dropna().tolist()  # List of 'To' recipients
    cc_recipients = receipient_df["cc"].dropna().tolist()  # List of 'CC' recipients
except Exception as e:
    print(f"❌ Error reading Excel file: {e}")
    sys.exit()

# Define sender email
from_email = "IN_D_Engineering_NIfi_Reports@airtel.com"

# Initialize Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
except Exception as e:
    print(f"❌ Error initializing Outlook: {e}")
    sys.exit()

# Today's date in YYYYMMDD format
today_date = datetime.now().strftime('%Y%m%d')

# Yesterday's date in DDMMYYYY format
yesterday_date = (datetime.now() - timedelta(days=1)).strftime('%d%m%Y')


email_subject = f"SFN_Report(Last 15 Days) - {yesterday_date}"
email_body = f"""
Dear Team,

Please find Key observations from the latest SFN Report:
🔹 Top 5 SFN Raised Circles: {top5_circles}
🔹 SFN Not Raised Circles:  {circles_with_zero_total}

Regards,  
Central Team FSO
"""

try:
    # Create a new email
    mail = outlook.CreateItem(0)
    mail.To = ";".join(to_recipients) if to_recipients else ""  # Join multiple recipients with ";"
    mail.CC = ";".join(cc_recipients) if cc_recipients else ""  # Join multiple CC recipients
    mail.Subject = email_subject
    mail.Body = email_body

    # Specify the sender email
    mail.SentOnBehalfOfName = from_email  # Option 1 (if account has permission)

    # Alternative approach (set the sender account explicitly)
    for account in outlook.Session.Accounts:
        if account.SmtpAddress == from_email:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # Option 2 (SendUsingAccount)

    # Attach the file
    mail.Attachments.Add(output_file)
    mail.Send()  # Send the email

    print(f"📧 Email sent successfully from {from_email}")

except Exception as e:
    print(f"❌ Error sending email: {e}")

# Release Outlook objects
del outlook, namespace

print("✅ Email process completed successfully!")
#=================================================================================================End of Code=========================================================================================================