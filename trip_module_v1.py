import pandas as pd
import numpy as np
import glob
import os
import win32com.client
import re
import requests
from datetime import datetime, timedelta
from urllib.parse import urlparse, parse_qs, unquote
import logging
import traceback
from pathlib import Path
import sys
import win32timezone

# ============================ Setup Logging ============================
log_dir = Path(r"D:\Automation\trip_module\log")
log_dir.mkdir(parents=True, exist_ok=True)
log_file = log_dir / f"trip_module_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("🔵 Script Execution Started")


# ============================ Constants & Config ============================
DOWNLOAD_PATH = Path(r"D:\Automation\trip_module\Upload")
OUTPUT_PATH = Path(r"D:\Automation\trip_module\output")
DOWNLOAD_PATH.mkdir(parents=True, exist_ok=True)
OUTPUT_PATH.mkdir(parents=True, exist_ok=True)

REQUIRED_EMAILS = {"MTD_vehicle_usage_Report"}
SUBJECT_KEYWORDS = {
    "MTD_vehicle_usage_Report": "MTD_vehicle_usage_Report",
}

USERNAME, PASSWORD = None, None

# ============================ Outlook Connection ============================
try:
    outlook_ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook_ns.GetDefaultFolder(6)
    target_folder = inbox.Folders("Required to Work")
    logging.info("✅ Connected to Outlook successfully.")
except Exception as e:
    logging.error(f"❌ Failed to connect to Outlook: {e}")
    sys.exit()


# ============================ Fetch Latest Emails ============================
latest_emails = {}
today = datetime.today().date()
date_filter = datetime.now() - timedelta(days=7)

# try:
#     messages = target_folder.Items
#     messages.Sort("[ReceivedTime]", True)
#     messages = messages.Restrict(f"[ReceivedTime] >= '{date_filter.strftime('%m/%d/%Y %H:%M %p')}'")

#     for message in messages:
#         try:
#             subject = message.Subject.strip()
#             received_date = message.ReceivedTime.date()

#             if received_date == today:
#                 for keyword, filename_part in SUBJECT_KEYWORDS.items():
#                     if keyword in subject:
#                         if filename_part in latest_emails:
#                             if message.ReceivedTime > latest_emails[filename_part].ReceivedTime:
#                                 latest_emails[filename_part] = message
#                         else:
#                             latest_emails[filename_part] = message
#         except Exception as e:
#             logging.warning(f"⚠️ Error processing email: {e}")

#     if not REQUIRED_EMAILS.issubset(set(latest_emails.keys())):
#         logging.error("❌ Required emails not found for today. Exiting.")
#         sys.exit()

#     logging.info("✅ All required emails fetched successfully.")

# except Exception as e:
#     logging.error(f"❌ Error fetching emails: {e}")
#     sys.exit()


# # ============================ Credentials ============================
# try:
#     credentials = pd.read_excel(r"D:\Automation\trip_module\credentials.xlsx")
#     USERNAME = credentials.iloc[0, 0].strip()
#     PASSWORD = credentials.iloc[0, 1].strip()
#     logging.info("✅ Credentials loaded successfully.")
# except Exception as e:
#     logging.error(f"❌ Failed to load credentials: {e}")
#     sys.exit()


# # ============================ URL Extractor ============================
# def extract_real_url(safelink):
#     if "safelinks.protection.outlook.com" in safelink:
#         parsed_url = urlparse(safelink)
#         params = parse_qs(parsed_url.query)
#         if "url" in params:
#             return unquote(params["url"][0])
#     return safelink


# # ============================ Download Attachments ============================
# for filename_part, latest_email in latest_emails.items():
#     try:
#         email_body = latest_email.Body
#         match = re.search(r"Link\s*[:\-]?\s*<?(https?://[^\s<>\"']+)>?", email_body)

#         if not match:
#             logging.warning(f"⚠️ No download link found in email: {filename_part}")
#             continue

#         download_url = extract_real_url(match.group(1))
#         retry_count = 3

#         for attempt in range(1, retry_count + 1):
#             try:
#                 response = requests.get(download_url, auth=(USERNAME, PASSWORD), allow_redirects=True)
#                 if response.status_code == 200:
#                     file_name = f"Mobility_{filename_part}_{today.strftime('%Y%m%d')}.csv"
#                     file_path = DOWNLOAD_PATH / file_name
#                     with open(file_path, "wb") as f:
#                         f.write(response.content)
#                     logging.info(f"✅ Downloaded successfully: {file_path}")
#                     break
#                 elif response.status_code == 401:
#                     logging.error(f"❌ Unauthorized (401): {filename_part}")
#                     break
#                 elif response.status_code == 500:
#                     logging.warning(f"⚠️ Server error 500, retrying ({attempt}/{retry_count})...")
#                 else:
#                     logging.error(f"❌ Unexpected error {response.status_code} for {filename_part}")
#                     break
#             except Exception:
#                 logging.error(f"❌ Error downloading {filename_part}: {traceback.format_exc()}")
#         else:
#             logging.error(f"❌ Failed to download {filename_part} after {retry_count} attempts.")
#     except Exception:
#         logging.error(f"❌ Error processing email {filename_part}: {traceback.format_exc()}")


# ============================ Data Processing ============================
try:
    today_str = today.strftime('%Y%m%d')
    yesterday = today - timedelta(days=1)
    yesterday_str = yesterday.strftime('%Y%m%d')

    file_paths = glob.glob(str(DOWNLOAD_PATH / f"*{today_str}*.csv"))
    if not file_paths:
        logging.error(f"❌ No files found for today's date in {DOWNLOAD_PATH}")
        sys.exit()

    columns_to_load = [
        "mtd_for_month", "date_range", "circle", "no_of_days_data_available", "no_of_days_vehicle_not_used",
        "vehicle_id", "vehicle_number", "vehicle_agency_admin", "driver_name", "driver_mobile_number"
    ] + [f"day_{str(i).zfill(2)}_attendance" for i in range(1, 32)]

    combined_data = pd.concat(
        [pd.read_csv(file, usecols=columns_to_load) for file in file_paths], ignore_index=True
    )

    attendance_cols = [col for col in combined_data.columns if "day_" in col]
    combined_data.drop(
        columns=[col for col in attendance_cols if (combined_data[col] == "Absent").all()],
        inplace=True
    )

    logging.info("✅ Data loaded and cleaned successfully.")

except Exception:
    logging.error(f"❌ Data processing error: {traceback.format_exc()}")
    sys.exit()


# ============================ Vendor Wise Segregation ============================
vendors = {
    "Choudhary Tours & Travels Pvt. Ltd": "Choudhary Tours & Travels Pvt. Ltd",
    "Mahindra Logistics Ltd": "Mahindra Logistics Ltd",
    "STEELMAN TELECOM LIMITED": "STEELMAN TELECOM LIMITED",
    "Rohtak Telecom": "Rohtak Telecom",
    "Integrated Wireless Solutions Private Limited": "Integrated Wireless Solutions Private Limited",
    "Matoshree Tours & Travels": "Matoshree Tours & Travels",
    "BHUMI TOUR AND TRAVELS PRIVATE LIMITED": "BHUMI TOUR AND TRAVELS PRIVATE LIMITED",
    "DROR SERVICES PRIVATE LIMITED": "DROR SERVICES PRIVATE LIMITED",
    "AXIOM INFRACOM Mobility Service": "AXIOM INFRACOM Mobility Service",
    "Global Telecom": "Global Telecom"
}

for key, vendor in vendors.items():
    df_vendor = combined_data[combined_data["vehicle_agency_admin"] == vendor]
    df_vendor.to_csv(OUTPUT_PATH / f"{vendor}_{yesterday_str}.csv", index=False)
    logging.info(f"✅ Vendor data saved: {vendor}")


# ============================ Email Sending with Retry & HTML ============================
try:
    recipients_df = pd.read_excel(r"D:\Automation\trip_module\Receipient.xlsx")
    vendor_recipients = dict(zip(recipients_df["vendor"], recipients_df["to"]))
    vendor_cc = dict(zip(recipients_df["vendor"], recipients_df["cc"]))
except Exception as e:
    logging.error(f"❌ Failed to read recipient list: {e}")
    sys.exit()

try:
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    session = outlook_app.GetNamespace("MAPI")
except Exception as e:
    logging.error(f"❌ Failed to initialize Outlook for sending emails: {e}")
    sys.exit()

from_email = "IN_D_Engineering_NIfi_Reports@airtel.com"

email_body_html = """
<html>
<body>
<p>Dear Team,</p>
<p>Please find attached <strong>MTD Vehicle Usage Report</strong> for <strong>{vendor}</strong>.</p>
<p>Regards,<br>Central Team FSO</p>
</body>
</html>
"""

for vendor_name in vendor_recipients.keys():
    file_to_send = OUTPUT_PATH / f"{vendor_name}_{yesterday_str}.csv"
    if not file_to_send.exists():
        logging.warning(f"⚠️ Missing file for vendor: {vendor_name}")
        continue

    for attempt in range(1, 4):  # Retry up to 3 times
        try:
            mail = outlook_app.CreateItem(0)
            mail.To = vendor_recipients[vendor_name]
            mail.CC = vendor_cc.get(vendor_name, "")
            mail.Subject = f"MTD Vehicle Usage Report - {vendor_name} - September"
            mail.HTMLBody = email_body_html.format(vendor=vendor_name)
            mail.SentOnBehalfOfName = from_email

            for account in outlook_app.Session.Accounts:
                if account.SmtpAddress == from_email:
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))

            mail.Attachments.Add(str(file_to_send))
            mail.Send()
            logging.info(f"📧 Email sent successfully to {vendor_recipients[vendor_name]} for {vendor_name}")
            break
        except Exception:
            if attempt == 3:
                logging.error(f"❌ Failed to send email to {vendor_name} after 3 attempts.")
            else:
                logging.warning(f"⚠️ Email send failed for {vendor_name}, retrying ({attempt}/3)...")

# Cleanup
del outlook_app, session
logging.info("✅ Script Execution Completed Successfully")
# ============================ End of Script ============================