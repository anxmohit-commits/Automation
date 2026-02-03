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
#================================================================================Fetching Latest Report from outlook========================================================
#Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

#**1️⃣ Define Download Path**
save_path = Path(r"D:\Automation\trip_module\Upload")
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
    "MTD_vehicle_usage_Report": "MTD_vehicle_usage_Report",
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
required_files = {"MTD_vehicle_usage_Report"}

if not required_files.issubset(set(latest_emails.keys())):
    logging.error("❌ MTD_vehicle_usage_Report email not found. Exiting...")
    sys.exit()

logging.info("✅ All required emails found. Proceeding with download...")

#**6️⃣ Ask for Credentials**
# USERNAME = input("Enter your username: ").strip()
# PASSWORD = getpass("Enter your password: ").strip()


credentials = pd.read_excel(r"D:\Automation\trip_module\credentials.xlsx")
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

# Folder path containing the files
folder_path = "D:/Automation/trip_module/Upload"
# Get all .xlsx files in the folder
today_date = datetime.now().strftime('%Y%m%d')
file_paths = glob.glob(os.path.join(folder_path,f"*{today_date}*.csv"))
#file_paths = glob.glob(os.path.join(folder_path, f"*{today_yyyymmdd}*.csv"))

# Exit if no today's file found
if not file_paths:
    print(f"❌ No files found for today's date ({today_date}) in folder: {folder_path}")
    sys.exit()

# Columns to load from each file
columns_to_load = ["mtd_for_month" , 
                    "date_range" , 
                    "circle" , 
                    "no_of_days_data_available" , 
                    "no_of_days_vehicle_not_used" , 
                    "vehicle_id" , 
                    "vehicle_number" , 
                    "vehicle_agency_admin" , 
                    "driver_name" , 
                    "driver_mobile_number" , 
                    "day_01_attendance" , 
                    "day_02_attendance" , 
                    "day_03_attendance" , 
                    "day_04_attendance" , 
                    "day_05_attendance" , 
                    "day_06_attendance" , 
                    "day_07_attendance" , 
                    "day_08_attendance" , 
                    "day_09_attendance" , 
                    "day_10_attendance" , 
                    "day_11_attendance" , 
                    "day_12_attendance" , 
                    "day_13_attendance" , 
                    "day_14_attendance" , 
                    "day_15_attendance" , 
                    "day_16_attendance" , 
                    "day_17_attendance" , 
                    "day_18_attendance" , 
                    "day_19_attendance" , 
                    "day_20_attendance" , 
                    "day_21_attendance" , 
                    "day_22_attendance" , 
                    "day_23_attendance" , 
                    "day_24_attendance" , 
                    "day_25_attendance" , 
                    "day_26_attendance" , 
                    "day_27_attendance" , 
                    "day_28_attendance" , 
                    "day_29_attendance" , 
                    "day_30_attendance" , 
                    "day_31_attendance"  
]

# Initialize a list to hold data from each file
all_data = []

# Read each file and append the data to the list
for file_path in file_paths:
    try:
        file_size = os.path.getsize(file_path) / (1024 * 1024)  # Convert bytes to MB
        print(f"Reading file: {os.path.basename(file_path)} | Size: {file_size:.2f} MB")
        
        data = pd.read_csv(file_path, usecols=columns_to_load)
        all_data.append(data)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")


# Concatenate all data into a single DataFrame
if all_data:
    combined_data = pd.concat(all_data, ignore_index=True)
    print("Data successfully combined.")
    # Identify attendance columns
    attendance_columns = [col for col in columns_to_load if "day_" in col]

    # Drop attendance columns where all values are "Absent"
    combined_data = combined_data.drop(
        columns=[col for col in attendance_columns if (combined_data[col] == "Absent").all()]
    )

    print("Columns with only 'Absent' values removed.")
else:
    combined_data = pd.DataFrame(columns=columns_to_load)
    print("No valid data found.")

#=========================================================================Sagrigating Vendorwise==============================================================================================

df_choudhary = combined_data[combined_data['vehicle_agency_admin'] == "Choudhary Tours & Travels Pvt. Ltd"]
df_mahindra = combined_data[combined_data['vehicle_agency_admin'] == "Mahindra Logistics Ltd"]
df_steelman = combined_data[combined_data['vehicle_agency_admin'] == "STEELMAN TELECOM LIMITED"]
df_rohtak_telecom = combined_data[combined_data['vehicle_agency_admin'] == "Rohtak Telecom"]
df_iws =  combined_data[combined_data['vehicle_agency_admin'] == "Integrated Wireless Solutions Private Limited"]
df_matoshree = combined_data[combined_data['vehicle_agency_admin'] == "Matoshree Tours & Travels"]




#==========================================================================printing==========================================================================================================

# Save the data to Excel
today_date = datetime.now().strftime('%Y%m%d')
df_choudhary.to_csv(f"D:\\Automation\\trip_module\\output\\Choudhary Tours & Travels Pvt. Ltd_{today_date}.csv", index=False)
df_mahindra.to_csv(f"D:\\Automation\\trip_module\\output\\Mahindra Logistics Ltd_{today_date}.csv", index=False)
df_steelman.to_csv(f"D:\\Automation\\trip_module\\output\\STEELMAN TELECOM LIMITED_{today_date}.csv", index=False)
df_rohtak_telecom.to_csv(f"D:\\Automation\\trip_module\\output\\Rohtak Telecom_{today_date}.csv", index=False)
df_iws.to_csv(f"D:\\Automation\\trip_module\\output\\Integrated Wireless Solutions Private Limited_{today_date}.csv", index=False)
df_matoshree.to_csv(f"D:\\Automation\\trip_module\\output\\Matoshree Tours & Travels_{today_date}.csv", index=False)

#========================================= Sending Emails with Attachments =========================================


# Read vendor recipients from Excel
file_path_excel = r"D:\Automation\trip_module\Receipient.xlsx"
try:
    Receipient = pd.read_excel(file_path_excel)
    # Convert DataFrame to dictionary
    vendor_recipients = dict(zip(Receipient["vendor"], Receipient["to"]))
    vendor_cc_recipients = dict(zip(Receipient["vendor"], Receipient["cc"]))
except Exception as e:
    print(f"❌ Error reading Excel file: {e}")
    sys.exit()



# Define From email
from_email = "IN_D_Engineering_NIfi_Reports@airtel.com"
#from_email = "kaushal3.kumar@airtel.com"

# Initialize Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
except Exception as e:
    print(f"❌ Error initializing Outlook: {e}")
    sys.exit()

# Get today's date for email subject
today_date = datetime.now().strftime('%Y%m%d')

email_subject_template = "MTD Vehicle Usage Report - {vendor}"
email_body_template = """
Dear Team,

Please find attached MTD Vehicle Usage Report for {vendor}.

Regards,  
Central Team FSO

"""
# Loop through vendors and send emails
for vendor, email in vendor_recipients.items():
    file_path = f"D:\\Automation\\trip_module\\output\\{vendor}_{today_date}.csv"

    # Check if the file exists before sending
    if not os.path.exists(file_path):
        print(f"⚠️ File not found for {vendor}: {file_path}")
        continue  # Skip this vendor if file is missing

    try:
        # Create a new email
        mail = outlook.CreateItem(0)
        mail.To = vendor_recipients.get(vendor, "")
        mail.CC = vendor_cc_recipients.get(vendor, "")  # Add CC recipient if available
        mail.Subject = email_subject_template.format(vendor=vendor, date=today_date)
        mail.Body = email_body_template.format(vendor=vendor)

        # Specify the sender email
        mail.SentOnBehalfOfName = from_email  # Option 1 (if account has permission)


        # Alternative approach (set the sender account explicitly)
        for account in outlook.Session.Accounts:
            if account.SmtpAddress == from_email:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # Option 2 (SendUsingAccount)



        # Attach the file
        mail.Attachments.Add(file_path)
        mail.Send()  # Send the email
        print(f"📧 Email sent successfully from {from_email} to {email} for {vendor}")

    except Exception as e:
        print(f"❌ Error sending email from {from_email} to {email} for {vendor}: {e}")

# Release Outlook objects
del outlook, namespace

print("✅ All emails processed successfully!")