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
log_dir = Path(r"D:\Automation\Rigger_Attendance\log")
log_dir.mkdir(parents=True, exist_ok=True)
log_file = log_dir / f"Rigger_Attendance_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("🔵 Script Execution Started")


# ============================ Constants & Config ============================
DOWNLOAD_PATH = Path(r"D:\Automation\Rigger_Attendance\Upload")
OUTPUT_PATH = Path(r"D:\Automation\Rigger_Attendance\output")
DOWNLOAD_PATH.mkdir(parents=True, exist_ok=True)
OUTPUT_PATH.mkdir(parents=True, exist_ok=True)

REQUIRED_EMAILS = {"Rigger_Daily_Attendance_Report"}
SUBJECT_KEYWORDS = {
    "Rigger_Daily_Attendance_Report": "Rigger_Daily_Attendance_Report",
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

try:
    messages = target_folder.Items
    messages.Sort("[ReceivedTime]", True)
    messages = messages.Restrict(f"[ReceivedTime] >= '{date_filter.strftime('%m/%d/%Y %H:%M %p')}'")

    for message in messages:
        try:
            subject = message.Subject.strip()
            received_date = message.ReceivedTime.date()

            if received_date == today:
                for keyword, filename_part in SUBJECT_KEYWORDS.items():
                    if keyword in subject:
                        if filename_part in latest_emails:
                            if message.ReceivedTime > latest_emails[filename_part].ReceivedTime:
                                latest_emails[filename_part] = message
                        else:
                            latest_emails[filename_part] = message
        except Exception as e:
            logging.warning(f"⚠️ Error processing email: {e}")

    if not REQUIRED_EMAILS.issubset(set(latest_emails.keys())):
        logging.error("❌ Required emails not found for today. Exiting.")
        sys.exit()

    logging.info("✅ All required emails fetched successfully.")

except Exception as e:
    logging.error(f"❌ Error fetching emails: {e}")
    sys.exit()


# ============================ Credentials ============================
try:
    credentials = pd.read_excel(r"D:\Automation\Rigger_Attendance\credentials.xlsx")
    USERNAME = credentials.iloc[0, 0].strip()
    PASSWORD = credentials.iloc[0, 1].strip()
    logging.info("✅ Credentials loaded successfully.")
except Exception as e:
    logging.error(f"❌ Failed to load credentials: {e}")
    sys.exit()


# ============================ URL Extractor ============================
def extract_real_url(safelink):
    if "safelinks.protection.outlook.com" in safelink:
        parsed_url = urlparse(safelink)
        params = parse_qs(parsed_url.query)
        if "url" in params:
            return unquote(params["url"][0])
    return safelink


# ============================ Download Attachments ============================
for filename_part, latest_email in latest_emails.items():
    try:
        email_body = latest_email.Body
        match = re.search(r"Link\s*[:\-]?\s*<?(https?://[^\s<>\"']+)>?", email_body)

        if not match:
            logging.warning(f"⚠️ No download link found in email: {filename_part}")
            continue

        download_url = extract_real_url(match.group(1))
        retry_count = 3

        for attempt in range(1, retry_count + 1):
            try:
                response = requests.get(download_url, auth=(USERNAME, PASSWORD), allow_redirects=True)
                if response.status_code == 200:
                    file_name = f"Mobility_{filename_part}_{today.strftime('%Y%m%d')}.csv"
                    file_path = DOWNLOAD_PATH / file_name
                    with open(file_path, "wb") as f:
                        f.write(response.content)
                    logging.info(f"✅ Downloaded successfully: {file_path}")
                    break
                elif response.status_code == 401:
                    logging.error(f"❌ Unauthorized (401): {filename_part}")
                    break
                elif response.status_code == 500:
                    logging.warning(f"⚠️ Server error 500, retrying ({attempt}/{retry_count})...")
                else:
                    logging.error(f"❌ Unexpected error {response.status_code} for {filename_part}")
                    break
            except Exception:
                logging.error(f"❌ Error downloading {filename_part}: {traceback.format_exc()}")
        else:
            logging.error(f"❌ Failed to download {filename_part} after {retry_count} attempts.")
    except Exception:
        logging.error(f"❌ Error processing email {filename_part}: {traceback.format_exc()}")


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
        "name",
        "msisdn",
        "circle",
        "rigger_agency_organisation",
        "zone",
        "rigger_unique_id",
        "attendance_date",
        "attendance_status",
        "timestamp_of_attendance_marked",
        "attendance_approval_status",
        "attendance_approver_name",
        "attendance_approver_msisdn",
        "timestamp_of_approval_action",
        "attendance_marked_latitude",
        "attendance_marked_longitude",
        "full_body_safety_harness_availability",
        "full_body_safety_harness_condition",
        "work_positioning_lanyard_availability",
        "work_positioning_lanyard_condition",
        "safety_shoes_availability",
        "safety_shoes_condition",
        "safety_helmet_availability",
        "safety_helmet_condition",
        "cotton_gloves_availability",
        "cotton_gloves_condition",
        "tool_kit_availability",
        "tool_kit_condition"

    ]

    combined_data = pd.concat(
        [pd.read_csv(file, usecols=columns_to_load) for file in file_paths], ignore_index=True
    )

    combined_data['circle'] = combined_data['circle'].replace({'OD': 'OR', 'UPE': 'UE'})
    logging.info("✅ Data loaded and cleaned successfully.")

except Exception:
    logging.error(f"❌ Data processing error: {traceback.format_exc()}")
    sys.exit()

# Drop unnecessary columns
df_rigger_attendance = combined_data.drop(
    columns=[
        "attendance_approval_status",
        "attendance_approver_name",
        "attendance_approver_msisdn",
        "timestamp_of_approval_action"
    ],
    errors="ignore"   # avoids error if any column is missing
)

#========================Reading Rigger COunt==============================================

df_rigger_count = pd.read_excel(r"D:\Automation\Rigger_Attendance\Rigger_Count.xlsx")
#==============================Creating Summary for Attendance===============================
logging.info("✅ Cleaning data.")
df_summary_dump = df_rigger_attendance[
    df_rigger_attendance['timestamp_of_attendance_marked'].notna() & 
    (df_rigger_attendance['timestamp_of_attendance_marked'].astype(str).str.strip() != "")
]

df_summary_dump['Attendance_Marked_Date'] = pd.to_datetime(df_summary_dump['timestamp_of_attendance_marked'], errors='coerce')
df_summary_dump['Attendance_Marked_Date'] = pd.to_datetime(df_summary_dump['Attendance_Marked_Date']).dt.date

# Calculate date 5 days ago
five_days_ago = datetime.now() - timedelta(days=5)

# Filter for last 5 days
df_last_5_days = df_summary_dump[df_summary_dump['Attendance_Marked_Date'] >= five_days_ago.date()]


# Pivot WITHOUT margins; keep daily unique counts
df_attendance_summary = df_last_5_days.pivot_table(
    index='circle',
    columns='Attendance_Marked_Date',
    values='msisdn',
    aggfunc='nunique',
    fill_value=0
)


# Format the column headers to show date in the desired format
df_attendance_summary.columns = [col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col for col in df_attendance_summary.columns]

# (Optional) sort date columns
df_attendance_summary = df_attendance_summary.reindex(sorted(df_attendance_summary.columns), axis=1)

# ✅ Row-wise Total = sum across days (not unique across the whole period)
df_attendance_summary['Total'] = df_attendance_summary.sum(axis=1)

# ✅ Bottom Total row = column sums
total_row = df_attendance_summary.sum(axis=0).to_frame().T
total_row.index = ['Total']
df_attendance_summary = pd.concat([df_attendance_summary, total_row], axis=0)


# If you need 'circle' as a column:
df_attendance_summary = df_attendance_summary.reset_index()

df_attendance_summary = df_attendance_summary.rename(columns={"index": "circle"})

#================================Merging with Daily ATtendance==================

# Merge the two base dataframes first
df_summary = df_rigger_count.merge(df_attendance_summary, on="circle", how="left")

# Identify date columns (exclude non-date ones)
date_cols = [col for col in df_summary.columns if col not in ["circle", "Qty", "Total"]]

# Add percentage columns with % sign
for col in date_cols:
    df_summary[f"{col}_%age"] = (
        (df_summary[col] / df_summary["Qty"] * 100).round(2).astype(str) + '%'
    )

# ============================ Vendor Wise Segregation ============================
vendors = {
    "Leo Communications_Rigger_Agency": "Leo Communications_Rigger_Agency",
    "Aerial_Rigger_Agency": "Aerial_Rigger_Agency",
    "Steelman_Rigger_Agency": "Steelman_Rigger_Agency",
    "Shivam_Rigger_Agency": "Shivam_Rigger_Agency",
    "Matoshree_Rigger_Agency": "Matoshree_Rigger_Agency",
    "Vedang_Rigger_Agency": "Vedang_Rigger_Agency",
    "Infotech Solution_Rigger_Agency": "Infotech Solution_Rigger_Agency",
    "Accord_Rigger_Agency": "Accord_Rigger_Agency",
    "Neon Mobicam_Rigger_Agency": "Neon Mobicam_Rigger_Agency",
    "RENcomm_Rigger_Agency": "RENcomm_Rigger_Agency",
    "IWS_Rigger_Agency": "IWS_Rigger_Agency",
    "HRJ_Rigger_Agency": "HRJ_Rigger_Agency"
}

for key, vendor in vendors.items():
    df_vendor = df_rigger_attendance[df_rigger_attendance["rigger_agency_organisation"] == vendor]
    df_vendor.to_csv(OUTPUT_PATH / f"{vendor}_{yesterday_str}.csv", index=False)
    logging.info(f"✅ Vendor data saved: {vendor}")





output_file_path = rf"D:\Automation\Rigger_Attendance\Summary\Rigger_Attendance.xlsx"
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
import tkinter as tk
from tkinter import filedialog, Tk

# ============================ Setup Logging ============================
log_dir = Path(r"D:\Automation\Rigger_Attendance\log")
log_dir.mkdir(parents=True, exist_ok=True)
log_file = log_dir / f"Rigger_Attendance_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("🔵 Script Execution Started")


# ============================ Constants & Config ============================
DOWNLOAD_PATH = Path(r"D:\Automation\Rigger_Attendance\Upload")
OUTPUT_PATH = Path(r"D:\Automation\Rigger_Attendance\output")
DOWNLOAD_PATH.mkdir(parents=True, exist_ok=True)
OUTPUT_PATH.mkdir(parents=True, exist_ok=True)

REQUIRED_EMAILS = {"Rigger_Daily_Attendance_Report"}
SUBJECT_KEYWORDS = {
    "Rigger_Daily_Attendance_Report": "Rigger_Daily_Attendance_Report",
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

try:
    messages = target_folder.Items
    messages.Sort("[ReceivedTime]", True)
    messages = messages.Restrict(f"[ReceivedTime] >= '{date_filter.strftime('%m/%d/%Y %H:%M %p')}'")

    for message in messages:
        try:
            subject = message.Subject.strip()
            received_date = message.ReceivedTime.date()

            if received_date == today:
                for keyword, filename_part in SUBJECT_KEYWORDS.items():
                    if keyword in subject:
                        if filename_part in latest_emails:
                            if message.ReceivedTime > latest_emails[filename_part].ReceivedTime:
                                latest_emails[filename_part] = message
                        else:
                            latest_emails[filename_part] = message
        except Exception as e:
            logging.warning(f"⚠️ Error processing email: {e}")

    if not REQUIRED_EMAILS.issubset(set(latest_emails.keys())):
        logging.error("❌ Required emails not found for today. Exiting.")
        sys.exit()

    logging.info("✅ All required emails fetched successfully.")

except Exception as e:
    logging.error(f"❌ Error fetching emails: {e}")
    sys.exit()


# ============================ Credentials ============================
try:
    credentials = pd.read_excel(r"D:\Automation\Rigger_Attendance\credentials.xlsx")
    USERNAME = credentials.iloc[0, 0].strip()
    PASSWORD = credentials.iloc[0, 1].strip()
    logging.info("✅ Credentials loaded successfully.")
except Exception as e:
    logging.error(f"❌ Failed to load credentials: {e}")
    sys.exit()


# ============================ URL Extractor ============================
def extract_real_url(safelink):
    if "safelinks.protection.outlook.com" in safelink:
        parsed_url = urlparse(safelink)
        params = parse_qs(parsed_url.query)
        if "url" in params:
            return unquote(params["url"][0])
    return safelink


# ============================ Download Attachments ============================
for filename_part, latest_email in latest_emails.items():
    try:
        email_body = latest_email.Body
        match = re.search(r"Link\s*[:\-]?\s*<?(https?://[^\s<>\"']+)>?", email_body)

        if not match:
            logging.warning(f"⚠️ No download link found in email: {filename_part}")
            continue

        download_url = extract_real_url(match.group(1))
        retry_count = 3

        for attempt in range(1, retry_count + 1):
            try:
                response = requests.get(download_url, auth=(USERNAME, PASSWORD), allow_redirects=True)
                if response.status_code == 200:
                    file_name = f"Mobility_{filename_part}_{today.strftime('%Y%m%d')}.csv"
                    file_path = DOWNLOAD_PATH / file_name
                    with open(file_path, "wb") as f:
                        f.write(response.content)
                    logging.info(f"✅ Downloaded successfully: {file_path}")
                    break
                elif response.status_code == 401:
                    logging.error(f"❌ Unauthorized (401): {filename_part}")
                    break
                elif response.status_code == 500:
                    logging.warning(f"⚠️ Server error 500, retrying ({attempt}/{retry_count})...")
                else:
                    logging.error(f"❌ Unexpected error {response.status_code} for {filename_part}")
                    break
            except Exception:
                logging.error(f"❌ Error downloading {filename_part}: {traceback.format_exc()}")
        else:
            logging.error(f"❌ Failed to download {filename_part} after {retry_count} attempts.")
    except Exception:
        logging.error(f"❌ Error processing email {filename_part}: {traceback.format_exc()}")


# ============================ USER FILE SELECTION ============================
def select_files_gui():
    """Allow user to select CSV/Excel files from DOWNLOAD_PATH"""
    root = Tk()
    root.withdraw()  # Hide the root window
    root.attributes('-topmost', True)  # Bring dialog to front
    
    file_paths = filedialog.askopenfilenames(
        title="Select CSV/Excel file(s) for processing",
        initialdir=str(DOWNLOAD_PATH),
        filetypes=[
            ("CSV Files", "*.csv"),
            ("Excel Files", "*.xlsx *.xls"),
            ("All Files", "*.*")
        ]
    )
    
    root.destroy()
    return list(file_paths) if file_paths else []


# ============================ Data Processing ============================
try:
    today_str = today.strftime('%Y%m%d')
    yesterday = today - timedelta(days=1)
    yesterday_str = yesterday.strftime('%Y%m%d')

    # Try auto-glob first
    file_paths = glob.glob(str(DOWNLOAD_PATH / f"*{today_str}*.csv"))
    
    # If no files found, ask user to select manually
    if not file_paths:
        logging.warning(f"⚠️ No files found for today's date. Showing file picker...")
        file_paths = select_files_gui()
        
        if not file_paths:
            logging.error(f"❌ No files selected by user. Exiting.")
            sys.exit()
    
    logging.info(f"✅ Using files: {file_paths}")

    columns_to_load = [
        "name",
        "msisdn",
        "circle",
        "rigger_agency_organisation",
        "zone",
        "rigger_unique_id",
        "date",
        "attendance_status",
        "timestamp_of_attendance_marked",
        "attendance_approval_status",
        "attendance_approver_name",
        "attendance_approver_msisdn",
        "timestamp_of_approval_action",
        "attendance_marked_latitude",
        "attendance_marked_longitude",
        "full_body_safety_harness_availability",
        "full_body_safety_harness_condition",
        "work_positioning_lanyard_availability",
        "work_positioning_lanyard_condition",
        "safety_shoes_availability",
        "safety_shoes_condition",
        "safety_helmet_availability",
        "safety_helmet_condition",
        "cotton_gloves_availability",
        "cotton_gloves_condition",
        "tool_kit_availability",
        "tool_kit_condition"
    ]

    combined_data = pd.concat(
        [pd.read_csv(file, usecols=columns_to_load) for file in file_paths], ignore_index=True
    )

    combined_data['circle'] = combined_data['circle'].replace({'OD': 'OR', 'UPE': 'UE'})
    logging.info("✅ Data loaded and cleaned successfully.")

except Exception:
    logging.error(f"❌ Data processing error: {traceback.format_exc()}")
    sys.exit()

# Drop unnecessary columns
df_rigger_attendance = combined_data.drop(
    columns=[
        "attendance_approval_status",
        "attendance_approver_name",
        "attendance_approver_msisdn",
        "timestamp_of_approval_action"
    ],
    errors="ignore"
)

#========================Reading Rigger Count==============================================

df_rigger_count = pd.read_excel(r"D:\Automation\Rigger_Attendance\Rigger_Count.xlsx")
#==============================Creating Summary for Attendance===============================
logging.info("✅ Cleaning data.")
df_summary_dump = df_rigger_attendance[
    df_rigger_attendance['timestamp_of_attendance_marked'].notna() & 
    (df_rigger_attendance['timestamp_of_attendance_marked'].astype(str).str.strip() != "")
]

df_summary_dump['Attendance_Marked_Date'] = pd.to_datetime(df_summary_dump['timestamp_of_attendance_marked'], errors='coerce')
df_summary_dump['Attendance_Marked_Date'] = pd.to_datetime(df_summary_dump['Attendance_Marked_Date']).dt.date

# Calculate date 5 days ago
five_days_ago = datetime.now() - timedelta(days=5)

# Filter for last 5 days
df_last_5_days = df_summary_dump[df_summary_dump['Attendance_Marked_Date'] >= five_days_ago.date()]


# Pivot WITHOUT margins; keep daily unique counts
df_attendance_summary = df_last_5_days.pivot_table(
    index='circle',
    columns='Attendance_Marked_Date',
    values='msisdn',
    aggfunc='nunique',
    fill_value=0
)


# Format the column headers to show date in the desired format
df_attendance_summary.columns = [col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col for col in df_attendance_summary.columns]

# (Optional) sort date columns
df_attendance_summary = df_attendance_summary.reindex(sorted(df_attendance_summary.columns), axis=1)

# ✅ Row-wise Total = sum across days (not unique across the whole period)
df_attendance_summary['Total'] = df_attendance_summary.sum(axis=1)

# ✅ Bottom Total row = column sums
total_row = df_attendance_summary.sum(axis=0).to_frame().T
total_row.index = ['Total']
df_attendance_summary = pd.concat([df_attendance_summary, total_row], axis=0)


# If you need 'circle' as a column:
df_attendance_summary = df_attendance_summary.reset_index()

df_attendance_summary = df_attendance_summary.rename(columns={"index": "circle"})

#================================Merging with Daily Attendance==================

# Merge the two base dataframes first
df_summary = df_rigger_count.merge(df_attendance_summary, on="circle", how="left")

# Identify date columns (exclude non-date ones)
date_cols = [col for col in df_summary.columns if col not in ["circle", "Qty", "Total"]]

# Add percentage columns with % sign
for col in date_cols:
    df_summary[f"{col}_%age"] = (
        (df_summary[col] / df_summary["Qty"] * 100).round(2).astype(str) + '%'
    )

# ============================ Vendor Wise Segregation ============================
vendors = {
    "Leo Communications_Rigger_Agency": "Leo Communications_Rigger_Agency",
    "Aerial_Rigger_Agency": "Aerial_Rigger_Agency",
    "Steelman_Rigger_Agency": "Steelman_Rigger_Agency",
    "Shivam_Rigger_Agency": "Shivam_Rigger_Agency",
    "Matoshree_Rigger_Agency": "Matoshree_Rigger_Agency",
    "Vedang_Rigger_Agency": "Vedang_Rigger_Agency",
    "Infotech Solution_Rigger_Agency": "Infotech Solution_Rigger_Agency",
    "Accord_Rigger_Agency": "Accord_Rigger_Agency",
    "Neon Mobicam_Rigger_Agency": "Neon Mobicam_Rigger_Agency",
    "RENcomm_Rigger_Agency": "RENcomm_Rigger_Agency",
    "IWS_Rigger_Agency": "IWS_Rigger_Agency",
    "HRJ_Rigger_Agency": "HRJ_Rigger_Agency"
}

for key, vendor in vendors.items():
    df_vendor = df_rigger_attendance[df_rigger_attendance["rigger_agency_organisation"] == vendor]
    df_vendor.to_csv(OUTPUT_PATH / f"{vendor}_{yesterday_str}.csv", index=False)
    logging.info(f"✅ Vendor data saved: {vendor}")


output_file_path = rf"D:\Automation\Rigger_Attendance\Summary\Rigger_Attendance.xlsx"


with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer: 
    df_summary.to_excel(writer, sheet_name="Attendance Summary", index=False)
    combined_data.to_excel(writer, sheet_name="Dump", index=False)


# ============================ CONFIG ============================
OUTPUT_PATH = Path(r"D:\Automation\Rigger_Attendance\Output")
VENDOR_RECIPIENT_FILE = r"D:\Automation\Rigger_Attendance\Receipient.xlsx"
CENTRAL_RECIPIENT_FILE = r"D:\Automation\Rigger_Attendance\Central_Receipient.xlsx"
from_email = "IN_D_Engineering_NIfi_Reports@airtel.com"

yesterday_str = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
yesterday_display = (datetime.now() - timedelta(days=1)).strftime('%d-%b-%y')

# ============================ LOGGING ============================
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# ============================ READ RECIPIENT FILES ============================
try:
    recipients_df = pd.read_excel(VENDOR_RECIPIENT_FILE)
    vendor_recipients = dict(zip(recipients_df["vendor"], recipients_df["to"]))
    vendor_cc = dict(zip(recipients_df["vendor"], recipients_df["cc"]))

    central_df = pd.read_excel(CENTRAL_RECIPIENT_FILE)
    central_to = central_df["to"].dropna().tolist()
    central_cc = central_df["cc"].dropna().tolist()

except Exception as e:
    logging.error(f"❌ Failed to read recipient list: {e}")
    sys.exit()

# ============================ OUTLOOK INITIALIZATION ============================
try:
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    session = outlook_app.GetNamespace("MAPI")
except Exception as e:
    logging.error(f"❌ Failed to initialize Outlook: {e}")
    sys.exit()

# ============================ COMMON SEND FUNCTION ============================
def send_email_with_retry(to_list, cc_list, subject, body_html, attachment_path=None, retries=3):
    for attempt in range(1, retries + 1):
        try:
            mail = outlook_app.CreateItem(0)
            mail.To = ";".join(to_list) if isinstance(to_list, list) else to_list
            mail.CC = ";".join(cc_list) if isinstance(cc_list, list) else cc_list
            mail.Subject = subject
            mail.HTMLBody = body_html
            mail.SentOnBehalfOfName = from_email

            # Force using correct account
            for account in outlook_app.Session.Accounts:
                if account.SmtpAddress == from_email:
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))

            if attachment_path and Path(attachment_path).exists():
                mail.Attachments.Add(str(attachment_path))

            mail.Send()
            logging.info(f"📧 Email sent successfully to {to_list}")
            break

        except Exception as e:
            if attempt == retries:
                logging.error(f"❌ Failed to send email to {to_list} after {retries} attempts. Error: {e}")
            else:
                logging.warning(f"⚠️ Email send failed to {to_list}, retrying ({attempt}/{retries})...")

# ============================ SEND VENDOR EMAILS ============================
email_body_vendor = """
<html>
<body>
<p>Dear Team,</p>
<p>Please find attached <strong>Rigger_Attendance Report</strong> for <strong>{vendor}</strong>.</p>
<p>Regards,<br>Central Team FSO</p>
</body>
</html>
"""

for vendor_name in vendor_recipients.keys():
    file_to_send = OUTPUT_PATH / f"{vendor_name}_{yesterday_str}.csv"
    if not file_to_send.exists():
        logging.warning(f"⚠️ Missing file for vendor: {vendor_name}")
        continue

    send_email_with_retry(
        to_list=vendor_recipients[vendor_name],
        cc_list=vendor_cc.get(vendor_name, ""),
        subject=f"Rigger_Attendance Report - {vendor_name} - {yesterday_display}",
        body_html=email_body_vendor.format(vendor=vendor_name),
        attachment_path=file_to_send
    )

# ============================ SEND CENTRAL EMAIL ============================
email_body_central = """
<html>
<body>
<p>Dear Team,</p>
<p>Please find attached <strong>Rigger Attendance Report</strong> for the last 5 days.</p>
<p>Regards,<br>Central Team FSO</p>
</body>
</html>
"""

send_email_with_retry(
    to_list=central_to,
    cc_list=central_cc,
    subject=f"Rigger_Attendance_Report - {yesterday_display}",
    body_html=email_body_central,
    attachment_path=output_file_path
)

# ============================ CLEANUP ============================
del outlook_app, session
logging.info("✅ All emails processed successfully!")