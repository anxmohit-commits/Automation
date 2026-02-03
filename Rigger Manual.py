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

# ============================ GUI FILE SELECTION ============================
def get_user_file():
    """Opens a window for the user to select the input sheet"""
    root = Tk()
    root.withdraw()  # Hide main tkinter window
    root.attributes('-topmost', True)  # Bring to front
    
    print("Opening File Selection Window... Please select your Input Sheet.")
    file_path = filedialog.askopenfilename(
        title="Select Input Sheet (Excel or CSV)",
        initialdir=str(DOWNLOAD_PATH),
        filetypes=[
            ("Data Files", "*.csv *.xlsx *.xls"),
            ("All Files", "*.*")
        ]
    )
    root.destroy()
    return file_path

# User se file input lena (Yahi aapne maanga tha)
selected_input_file = get_user_file()

if not selected_input_file:
    print("❌ No file selected. Script exiting.")
    logging.error("❌ User cancelled file selection. Exiting.")
    sys.exit()

logging.info(f"✅ User selected file: {selected_input_file}")

# ============================ Outlook Connection ============================
try:
    outlook_ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook_ns.GetDefaultFolder(6)
    target_folder = inbox.Folders("Required to Work")
    logging.info("✅ Connected to Outlook successfully.")
except Exception as e:
    logging.error(f"❌ Failed to connect to Outlook: {e}")
    # Logic continue rakhein agar email fetch na bhi ho sake manual input ke case mein

# ============================ Credentials ============================
try:
    credentials = pd.read_excel(r"D:\Automation\Rigger_Attendance\credentials.xlsx")
    USERNAME = credentials.iloc[0, 0].strip()
    PASSWORD = credentials.iloc[0, 1].strip()
    logging.info("✅ Credentials loaded successfully.")
except Exception as e:
    logging.error(f"❌ Failed to load credentials: {e}")
    sys.exit()

# ============================ Data Processing (Aapka Logic) ============================
try:
    today = datetime.today().date()
    today_str = today.strftime('%Y%m%d')
    yesterday_str = (today - timedelta(days=1)).strftime('%Y%m%d')

    columns_to_load = [
        "name", "msisdn", "circle", "rigger_agency_organisation", "zone",
        "rigger_unique_id", "date", "attendance_status", "timestamp_of_attendance_marked",
        "attendance_approval_status", "attendance_approver_name", "attendance_approver_msisdn",
        "timestamp_of_approval_action", "attendance_marked_latitude", "attendance_marked_longitude",
        "full_body_safety_harness_availability", "full_body_safety_harness_condition",
        "work_positioning_lanyard_availability", "work_positioning_lanyard_condition",
        "safety_shoes_availability", "safety_shoes_condition", "safety_helmet_availability",
        "safety_helmet_condition", "cotton_gloves_availability", "cotton_gloves_condition",
        "tool_kit_availability", "tool_kit_condition"
    ]

    # File loading based on user selection
    if selected_input_file.lower().endswith('.csv'):
        combined_data = pd.read_csv(selected_input_file, usecols=columns_to_load)
    else:
        combined_data = pd.read_excel(selected_input_file, usecols=columns_to_load)

    combined_data['circle'] = combined_data['circle'].replace({'OD': 'OR', 'UPE': 'UE'})
    logging.info("✅ Data loaded from user file and cleaned.")

except Exception:
    logging.error(f"❌ Data processing error: {traceback.format_exc()}")
    sys.exit()

# Baki saara logic as-is (Logic mein koi ched-chad nahi ki gayi hai)
df_rigger_attendance = combined_data.drop(
    columns=["attendance_approval_status", "attendance_approver_name", "attendance_approver_msisdn", "timestamp_of_approval_action"],
    errors="ignore"
)

df_rigger_count = pd.read_excel(r"D:\Automation\Rigger_Attendance\Rigger_Count.xlsx")

df_summary_dump = df_rigger_attendance[
    df_rigger_attendance['timestamp_of_attendance_marked'].notna() & 
    (df_rigger_attendance['timestamp_of_attendance_marked'].astype(str).str.strip() != "")
]

df_summary_dump['Attendance_Marked_Date'] = pd.to_datetime(df_summary_dump['timestamp_of_attendance_marked'], errors='coerce').dt.date

five_days_ago = (datetime.now() - timedelta(days=5)).date()
df_last_5_days = df_summary_dump[df_summary_dump['Attendance_Marked_Date'] >= five_days_ago]

df_attendance_summary = df_last_5_days.pivot_table(
    index='circle', columns='Attendance_Marked_Date', values='msisdn', aggfunc='nunique', fill_value=0
)

df_attendance_summary.columns = [col.strftime('%d %b') if hasattr(col, 'strftime') else col for col in df_attendance_summary.columns]
df_attendance_summary = df_attendance_summary.reindex(sorted(df_attendance_summary.columns), axis=1)
df_attendance_summary['Total'] = df_attendance_summary.sum(axis=1)

total_row = df_attendance_summary.sum(axis=0).to_frame().T
total_row.index = ['Total']
df_attendance_summary = pd.concat([df_attendance_summary, total_row], axis=0)
df_attendance_summary = df_attendance_summary.reset_index().rename(columns={"index": "circle"})

df_summary = df_rigger_count.merge(df_attendance_summary, on="circle", how="left")
date_cols = [col for col in df_summary.columns if col not in ["circle", "Qty", "Total"]]

for col in date_cols:
    df_summary[f"{col}_%age"] = ((df_summary[col] / df_summary["Qty"] * 100).round(2).fillna(0).astype(str) + '%')

# Vendor splitting
vendors = ["Leo Communications_Rigger_Agency", "Aerial_Rigger_Agency", "Steelman_Rigger_Agency", 
           "Shivam_Rigger_Agency", "Matoshree_Rigger_Agency", "Vedang_Rigger_Agency", 
           "Infotech Solution_Rigger_Agency", "Accord_Rigger_Agency", "Neon Mobicam_Rigger_Agency", 
           "RENcomm_Rigger_Agency", "IWS_Rigger_Agency", "HRJ_Rigger_Agency"]

for vendor in vendors:
    df_vendor = df_rigger_attendance[df_rigger_attendance["rigger_agency_organisation"] == vendor]
    df_vendor.to_csv(OUTPUT_PATH / f"{vendor}_{yesterday_str}.csv", index=False)

output_file_path = r"D:\Automation\Rigger_Attendance\Summary\Rigger_Attendance.xlsx"
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer: 
    df_summary.to_excel(writer, sheet_name="Attendance Summary", index=False)
    combined_data.to_excel(writer, sheet_name="Dump", index=False)

# Email sending logic as-is
VENDOR_RECIPIENT_FILE = r"D:\Automation\Rigger_Attendance\Receipient.xlsx"
CENTRAL_RECIPIENT_FILE = r"D:\Automation\Rigger_Attendance\Central_Receipient.xlsx"
from_email = "IN_D_Engineering_NIfi_Reports@airtel.com"
yesterday_display = (datetime.now() - timedelta(days=1)).strftime('%d-%b-%y')

try:
    rec_df = pd.read_excel(VENDOR_RECIPIENT_FILE)
    v_to = dict(zip(rec_df["vendor"], rec_df["to"]))
    v_cc = dict(zip(rec_df["vendor"], rec_df["cc"]))
    c_df = pd.read_excel(CENTRAL_RECIPIENT_FILE)
    c_to = c_df["to"].dropna().tolist()
    c_cc = c_df["cc"].dropna().tolist()
except Exception as e:
    logging.error(f"❌ Recipient error: {e}")
    sys.exit()

try:
    outlook_app = win32com.client.Dispatch("Outlook.Application")
except Exception:
    logging.error("❌ Outlook fail")
    sys.exit()

def send_email(to_list, cc_list, subject, body, attach=None):
    try:
        mail = outlook_app.CreateItem(0)
        mail.To = ";".join(to_list) if isinstance(to_list, list) else str(to_list)
        mail.CC = ";".join(cc_list) if isinstance(cc_list, list) else str(cc_list)
        mail.Subject = subject
        mail.HTMLBody = body
        mail.SentOnBehalfOfName = from_email
        for acc in outlook_app.Session.Accounts:
            if acc.SmtpAddress.lower() == from_email.lower():
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                break
        if attach and Path(attach).exists(): mail.Attachments.Add(str(attach))
        mail.Send()
    except Exception as e: logging.error(f"Mail fail: {e}")

# Bulk sending
v_body = "<html><body><p>Dear Team,</p><p>Please find attached Report for <strong>{vendor}</strong>.</p><p>Regards,<br>Central Team FSO</p></body></html>"
for v_name in v_to.keys():
    f_path = OUTPUT_PATH / f"{v_name}_{yesterday_str}.csv"
    if f_path.exists():
        send_email(v_to[v_name], v_cc.get(v_name, ""), f"Rigger_Attendance Report - {v_name} - {yesterday_display}", v_body.format(vendor=v_name), f_path)

send_email(c_to, c_cc, f"Rigger_Attendance_Report - {yesterday_display}", "<html><body><p>Dear Team,</p><p>Please find attached Summary.</p></body></html>", output_file_path)

logging.info("✅ All tasks completed successfully!")