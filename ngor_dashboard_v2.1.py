import os
import sys
import re
import glob
import shutil
import zipfile
import tempfile
import numpy as np
import pandas as pd
import openpyxl
import win32com.client
import win32timezone
from datetime import datetime, timedelta
import logging
from collections import defaultdict
import sys

#=============================
# LOGGING SETUP (UTF-8 Safe)
#=============================

log_folder = r'D:\Automation\One_Outage\log'
os.makedirs(log_folder, exist_ok=True)
log_file = os.path.join(log_folder, 'process_log.txt')

logger = logging.getLogger("OneOutageLogger")
logger.setLevel(logging.INFO)

if not logger.handlers:
    file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter('%(message)s'))
    logger.addHandler(console_handler)

def log_print(message, level="info"):
    if level == "info":
        logger.info(message)
    elif level == "error":
        logger.error(message)
    elif level == "warning":
        logger.warning(message)

#=============================
# CONFIGURATION
#=============================

SAVE_PATH_TT = r"D:\Ericsson TT Daily Report"
SAVE_PATH_NGOR = r"D:\NGOR_Report"
OUTLOOK_FOLDER = "Required to Work"
SENDER_NGOR = "performance.management@ericsson.com"

os.makedirs(SAVE_PATH_TT, exist_ok=True)
os.makedirs(SAVE_PATH_NGOR, exist_ok=True)

#=============================
# OUTLOOK SETUP
#=============================

try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
except Exception as e:
    log_print(f"❌ Error connecting to Outlook: {e}", "error")
    sys.exit(1)

try:
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    target_folder = inbox.Folders(OUTLOOK_FOLDER)
except Exception as e:
    log_print(f"❌ Error accessing '{OUTLOOK_FOLDER}' folder: {e}", "error")
    sys.exit(1)

messages = target_folder.Items
messages.Sort("[ReceivedTime]", True)

today = datetime.now().date()

#=============================
# PROCESS ERICSSON DAILY REPORT
#=============================

log_print("🔍 Checking for 'Ericsson Daily Report' emails...")

found_tt = False

for msg in messages:
    try:
        if msg.ReceivedTime.date() != today:
            break

        if "Ericsson Daily Report" in msg.Subject:
            log_print(f"✅ Found TT Email: {msg.Subject} | {msg.ReceivedTime.strftime('%H:%M')}")
            for att in msg.Attachments:
                if att.FileName.lower().endswith(".xlsx"):
                    base, ext = os.path.splitext(att.FileName)
                    date_str = msg.ReceivedTime.strftime('%Y%m%d_%H%M%S')
                    safe_name = re.sub(r'[\\/*?:"<>|]', "_", base)
                    final_name = f"{safe_name}_{date_str}{ext}"
                    att.SaveAsFile(os.path.join(SAVE_PATH_TT, final_name))
                    log_print(f"📂 Saved: {final_name}")
                    found_tt = True
    except Exception as e:
        log_print(f"⚠️ Error processing TT email: {e}", "error")

if not found_tt:
    log_print("❌ No 'Ericsson Daily Report' emails found for today.", "warning")

#=============================
# PROCESS NGOR REPORT
#=============================

log_print("\n🔍 Checking for 'NGOR Report' emails...")

circle_emails = defaultdict(dict)

for msg in messages:
    try:
        if msg.ReceivedTime.date() != today:
            break

        if msg.SenderEmailAddress.lower() == SENDER_NGOR:
            subject = msg.Subject.lower()
            if "bharti one outage report" in subject:
                parts = subject.split()
                circle = parts[-1].replace("-update", "")

                if "-update" in subject:
                    circle_emails[circle]["update"] = msg
                else:
                    if "update" not in circle_emails[circle]:
                        circle_emails[circle]["plain"] = msg

    except Exception as e:
        log_print(f"⚠️ Error reading email: {e}", "error")

found_ngor = False

for circle, records in circle_emails.items():
    msg_to_use = records.get("update") or records.get("plain")

    if msg_to_use:
        is_update = "update" in records
        status = "✅ Found Preferred Email" if is_update else "📂 Downloading Fallback Email"
        log_print(f"{status}: {msg_to_use.Subject} | {msg_to_use.ReceivedTime.strftime('%H:%M')}")

        try:
            for att in msg_to_use.Attachments:
                if att.FileName.lower().endswith(".xlsx"):
                    base, ext = os.path.splitext(att.FileName)
                    time_str = msg_to_use.ReceivedTime.strftime('%Y%m%d_%H%M%S')
                    safe_name = re.sub(r'[\\/*?:"<>|]', "_", base)
                    final_name = f"{safe_name}_{time_str}{ext}"
                    att.SaveAsFile(os.path.join(SAVE_PATH_NGOR, final_name))
                    log_print(f"📂 Saved: {final_name}")
                    found_ngor = True
        except Exception as e:
            log_print(f"⚠️ Error saving attachment for {circle}: {e}", "error")

if not found_ngor:
    log_print("❌ No NGOR Report emails found for today.", "warning")

log_print("\n🚀 Script completed successfully.")

# Cleanup COM objects
del messages, target_folder, inbox, outlook

# ============================
# Your Ericsson TT Report processing with logging
# ============================

# Define circle list and DataFrame
circle_list = [
    'AP', 'AS', 'BH', 'CH', 'DL', 'GJ', 'HP', 'HR', 'JK', 'KK', 'KL', 'KOL',
    'MH', 'MP', 'MU', 'NE', 'OR', 'PB', 'RJ', 'TN', 'UE', 'UW', 'WB'
]
df_circles = pd.DataFrame({'Circle': circle_list})





folder_path = r'D:\Ericsson TT Daily Report'

try:
    file_list = glob.glob(os.path.join(folder_path, '*.xlsx'))
    if not file_list:
        log_print(f"❌ No Excel files found in folder: {folder_path}", "error")
        raise FileNotFoundError(f"No files found in {folder_path}")

    latest_file = sorted(file_list)[-1]
    log_print(f"📄 Reading file: {os.path.basename(latest_file)}")

    df_ericsson_tt = pd.read_excel(latest_file, header=3)

    # Convert and clean dates
    df_ericsson_tt['Create Date'] = df_ericsson_tt['Create Date Time of TT'].dt.date
    df_ericsson_tt['Create Date'] = pd.to_datetime(df_ericsson_tt['Create Date'], errors='coerce')
    report_date_ericsson_tt = df_ericsson_tt['Create Date'].max().strftime('%d-%m-%Y')
    log_print(f"✅ Report date found: {report_date_ericsson_tt}")

    # Filtering Conditions
    outage_down_condition = df_ericsson_tt['Description'].str.contains(r'(_OUTAGE|Outage|Down)', case=False, na=False)
    not_l1la_condition = ~df_ericsson_tt['Description'].str.contains(r'L1LA member', case=False, na=False)
    base_stn_con = ~df_ericsson_tt['Description'].str.contains(r'Base Station Connectivity Down Minor', case=False, na=False)
    qia = ~df_ericsson_tt['Description'].str.contains(r'QIA', case=False, na=False)
    career_downlink = ~df_ericsson_tt['Description'].str.contains(r'AUTO ZTE 4G NVR-Carrier downlink data problem', case=False, na=False)
    ser_degr = ~df_ericsson_tt['Description'].str.contains(r'AUTO 5G OUTAGE-Service Degraded', case=False, na=False)

    # Apply conditions and create outage type column
    df_ericsson_tt['Outage Type'] = ((outage_down_condition) & (not_l1la_condition) & (base_stn_con) & (qia) & (career_downlink) & (ser_degr)) \
        .apply(lambda x: 'Outage/Down' if x else '')

    df_ericsson_tt.rename(columns={'Site Group/Circle': 'Circle'}, inplace=True)

    # Replace circle names with codes
    df_ericsson_tt['Circle'] = df_ericsson_tt['Circle'].replace({
        'Tamilnadu': 'TN',
        'Assam': 'AS',
        'UPE': 'UE',
        'Rajasthan': 'RJ',
        'Delhi': 'DL',
        'Chennai': 'CH',
        'Kerala': 'KL',
        'KO': 'KOL',
        'Jammu and Kashmir': 'JK',
        'North East': 'NE'
    })

    df_outage_tt = df_ericsson_tt[df_ericsson_tt['Outage Type'] == 'Outage/Down']
    df_outage_tt_count = df_outage_tt.groupby('Circle').size().reset_index(name='Outage Category TTs raised')

    log_print(f"✅ Total outage TTs counted: {df_outage_tt_count['Outage Category TTs raised'].sum()}")

except Exception as e:
    log_print(f"⚠️ Error processing Ericsson TT Report: {e}", "error")

#==============================================Reading Today received files (outage report)====================
input_folder_one_outage = r'D:\NGOR_Report'

sheet_to_df = {
    'Common Outage Format': [],
    'Locked_LODS': [],
    'Not in Site DB': [],
}

today_date = datetime.now().date()

try:
    excel_files = [
        f for f in os.listdir(input_folder_one_outage)
        if f.lower().endswith(('.xlsx', '.xls'))
    ]

    if not excel_files:
        log_print("🚫 No Excel files found in the folder.", "warning")
    else:
        today_files = []
        for file in excel_files:
            file_path = os.path.join(input_folder_one_outage, file)
            file_mod_date = datetime.fromtimestamp(os.path.getmtime(file_path)).date()
            if file_mod_date == today_date:
                today_files.append(file)

        log_print(f"📁 Found {len(excel_files)} Excel files, {len(today_files)} received today.\n")

        for file in excel_files:
            file_path = os.path.join(input_folder_one_outage, file)
            file_modified_date = datetime.fromtimestamp(os.path.getmtime(file_path)).date()

            if file_modified_date != today_date:
                continue

            try:
                with pd.ExcelFile(file_path) as xls:
                    for sheet_name in sheet_to_df.keys():
                        if sheet_name in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet_name)
                            sheet_to_df[sheet_name].append(df)
                            log_print(f"✔️ Sheet '{sheet_name}' loaded from: {file}")
                        else:
                            log_print(f"⚠️ Sheet '{sheet_name}' not found in: {file}", "warning")
            except Exception as e:
                log_print(f"❌ Error reading {file}: {e}", "error")

    combined_df = (
        pd.concat(sheet_to_df['Common Outage Format'], ignore_index=True)
        if sheet_to_df['Common Outage Format'] else pd.DataFrame()
    )
    combined_df_Locked_LODS = (
        pd.concat(sheet_to_df['Locked_LODS'], ignore_index=True)
        if sheet_to_df['Locked_LODS'] else pd.DataFrame()
    )
    combined_df_not_in_site_db = (
        pd.concat(sheet_to_df['Not in Site DB'], ignore_index=True)
        if sheet_to_df['Not in Site DB'] else pd.DataFrame()
    )

    log_print("\n✅ DataFrames are ready!")
    log_print(f"📊 Common Outage Format rows: {combined_df.shape[0]}")
    log_print(f"📊 Locked_LODS rows: {combined_df_Locked_LODS.shape[0]}")
    log_print(f"📊 Not in Site DB rows: {combined_df_not_in_site_db.shape[0]}")

    # Sanity Testing
    present_circles = combined_df['Circle'].unique().tolist()
    missing_circles = [circle for circle in circle_list if circle not in present_circles]

    if missing_circles:
        log_print(f"Missing Circles: {', '.join(missing_circles)}", "error")
        sys.exit()
    else:
        log_print("All Circles are present. Proceeding...")

except Exception as e:
    log_print(f"❌ Unexpected error: {e}", "error")





log_print("Cleaning column names...")
combined_df.columns = [col.strip() for col in combined_df.columns]

mttr_col_candidates = [col for col in combined_df.columns if col.startswith("MTTR Category")]
if not mttr_col_candidates:
    log_print("❌ MTTR Category column not found!", "error")
    mttr_col = None
else:
    mttr_col = mttr_col_candidates[0]
    log_print(f"Using MTTR column: {mttr_col}")

# Process common_outage_tt from combined_df
log_print("Processing Common Outage TT from combined_df...")
common_outage_tt = combined_df[['Circle', 'TT ID']].copy()
common_outage_tt['TT Numbers'] = common_outage_tt['TT ID'].str.findall(r'INC\d+')
common_outage_tt = common_outage_tt.explode('TT Numbers').dropna(subset=['TT Numbers']).reset_index(drop=True)
common_outage_tt = common_outage_tt.drop_duplicates(subset=['Circle', 'TT Numbers'])

# Process common_outage_tt_lod from combined_df_Locked_LODS
log_print("Processing Common Outage TT from Locked_LODS...")
common_outage_tt_lod = combined_df_Locked_LODS[['Circle', 'TT ID']].copy()
common_outage_tt_lod['TT Numbers'] = common_outage_tt_lod['TT ID'].str.findall(r'INC\d+')
common_outage_tt_lod = common_outage_tt_lod.explode('TT Numbers').dropna(subset=['TT Numbers']).reset_index(drop=True)
common_outage_tt_lod = common_outage_tt_lod.drop_duplicates(subset=['Circle', 'TT Numbers'])

# Process common_outage_tt_not_in_sdb from combined_df_not_in_site_db
log_print("Processing Common Outage TT from Not in Site DB...")
common_outage_tt_not_in_sdb = combined_df_not_in_site_db[['Circle', 'WorK Order ID']].copy()
common_outage_tt_not_in_sdb['TT Numbers'] = common_outage_tt_not_in_sdb['WorK Order ID'].str.findall(r'INC\d+')
common_outage_tt_not_in_sdb = common_outage_tt_not_in_sdb.explode('TT Numbers').dropna(subset=['TT Numbers']).reset_index(drop=True)
common_outage_tt_not_in_sdb = common_outage_tt_not_in_sdb.drop_duplicates(subset=['Circle', 'TT Numbers'])


def assign_tt_status(df, column_to_check, tt_list_df, status_column_name):
    tt_set = set(tt_list_df['TT Numbers'].dropna().astype(str))
    df[status_column_name] = df[column_to_check].astype(str).apply(
        lambda x: 'Matched' if x in tt_set else 'Not matched'
    )
    log_print(f"Assigned TT status for {status_column_name}")


log_print("Assigning TT statuses in df_ericsson_tt...")
assign_tt_status(df_ericsson_tt, 'Ericsson Incident Number', common_outage_tt, 'TT status COR')
assign_tt_status(df_ericsson_tt, 'Ericsson Incident Number', common_outage_tt_lod, 'TT status LODS')
assign_tt_status(df_ericsson_tt, 'Ericsson Incident Number', common_outage_tt_not_in_sdb, 'TT status (Not in Site DB)')


def get_final_tt_status_combination(row):
    cor = row['TT status COR'] == 'Matched'
    lods = row['TT status LODS'] == 'Matched'
    not_sdb = row['TT status (Not in Site DB)'] == 'Matched'
    if cor and lods and not_sdb:
        return 'Matched in All'
    elif cor and lods:
        return 'Matched in COR & LODS'
    elif cor and not_sdb:
        return 'Matched in COR & Not in Site DB'
    elif lods and not_sdb:
        return 'Matched in LODS & Not in Site DB'
    elif cor:
        return 'Matched in COR'
    elif lods:
        return 'Matched in LODS'
    elif not_sdb:
        return 'Matched in Not in Site DB'
    else:
        return 'Not matched'

log_print("Computing final TT status combination...")
df_ericsson_tt['Final TT Status'] = df_ericsson_tt.apply(get_final_tt_status_combination, axis=1)


combined_df['TT Number'] = combined_df['TT ID'].str.extract(r'(INC\d+)')

combined_df['Date'] = pd.to_datetime(combined_df['Date'], errors='coerce')
report_date_one_outage = combined_df['Date'].max().strftime('%d-%m-%Y')
log_print(f"Report Date (One Outage): {report_date_one_outage}")

log_print("Filling blank Physical ID from other columns...")

def fill_physical_id(row):
    if pd.isna(row['Physical ID']) or str(row['Physical ID']).strip() == '-':
        for col in ['Site ID - 2G', 'Site ID - 4G TDD', 'Site ID - 4G FDD', 'Site ID - 5G']:
            if pd.notna(row.get(col, None)) and str(row[col]).strip() != '-':
                return row[col]
    return row['Physical ID']

combined_df['Physical ID'] = combined_df.apply(fill_physical_id, axis=1)

log_print("Defining RCA detail from other columns...")


def classify_rca(oem, bb, bk, aq):
    # Normalize input values
    oem = str(oem).lower() if pd.notnull(oem) else ""
    bb = str(bb).lower() if pd.notnull(bb) else ""
    bk = str(bk).lower() if pd.notnull(bk) else ""       # ✅ NEW: Ensure 'bk' is lowercased
    aq = str(aq).lower() if pd.notnull(aq) else ""       # ✅ NEW: Added 'aq' for CR-ID logic

    if "ericsson" in oem:                                # ✅ CHANGED: from == to in for partial match
        if "system restart alert requested restart type: power" in bb:
            return "Infra-BTS Booted due to power"
        elif "input power failure" in bb:
            return "Infra-Input Power Failure/Unit Power Reset"    # ✅ CHANGED: Label matches Excel logic
        elif ("input voltage below configured threshold" in bb or 
              "pmvoltage down" in bb):                              # ✅ ADDED: pmVoltage condition
            return "Infra-Low Pdu input voltage/Voltage low than threshold/pmVoltage Down"
        elif any(f"voltage is {v}" in bb for v in range(36, 46)):  # ✅ UNCHANGED but matched to Excel
            return "Infra-voltage <46"
        elif ("critical temperature taken out of service" in bb or 
              "temperature high" in bb):                            # ✅ UNCHANGED
            return "Active(BSS)-High Temperature(Unit)"
        elif ("system restart alert requested restart type: cold" in bb or 
              "system restart alert requested restart type: data restore" in bb):
            return "Active(BSS)-Active(Cold Restart/Data Restore)"
        elif (":link failure" in bb or "link degraded" in bb):      # ✅ UNCHANGED
            return "Active(BSS)-Hardware"
        elif "cr-id :" in aq:                                       # ✅ NEW: CR-ID logic added for Planned
            return "Planned-Planned(CR)"
        else:
            return "Other/NORCA"

    elif "nsn" in oem:                                              # ✅ CHANGED: from == to in
        if "bts booted" in bb and "due to power on" in bb:
            return "Infra-BTS Booted due to power"
        elif "unit power reset" in bb:
            return "Infra-Input Power Failure/Unit Power Reset"
        elif "low pdu input" in bb:
            return "Infra-Low Pdu input voltage/Voltage low than threshold/pmVoltage Down"
        elif any(f"voltage: {v}" in bb for v in range(36, 46)):     # ✅ CHANGED: voltage pattern to "Voltage: X"
            return "Infra-voltage <46"
        elif "overheating" in bk:                                   # ✅ NEW: match from Excel formula
            return "Active(BSS)-High Temperature(Unit)"
        elif ("bts booted" in bb and any(phrase in bb for phrase in [
            "due to software reset", "due to commissioning change", "due to recovery"
        ])):
            return "Active(BSS)-BTS Boot-Active(Sw reset/Cmm change/Recovery)"
        elif "bts booted" in bb and "due to user request" in bb:
            return "Active(BSS)-BTS Booted-Planned(User request)"
        elif "administrative state change" in bb:
            return "Active(BSS)-Active-Planned"
        elif "cr-id :" in aq:                                       # ✅ NEW: CR-ID for NSN
            return "Planned-Planned(CR)"
        else:
            return "Other/NORCA"

    elif "huawei" in oem:                                           # ✅ NEW BLOCK: Huawei logic from Excel
        if any(x in bb for x in [
            "board power-off",
            "rf unit dc input power failure",
            "base station dc power supply abnormal",
            "unit external power supply insufficient"
        ]):
            return "Infra-Board Power Of/RF Unit Power Failure"
        elif "cr-id :" in aq:
            return "Planned-Planned(CR)"
        else:
            return "Other/NORCA"

    else:
        return "OEM not added"





combined_df['RCA detail'] = combined_df.apply(
    lambda row: classify_rca(row['OEM Name'], row['Auto RCA - From Internal alarm'], row['Alarm Description'], row['CR No.']),
    axis=1
)

log_print("Defining RCA conclusive from RCA detail...")

def classify_rca_conclusive(detail):
    if pd.isna(detail):
        return "NO RCA/Not conclusive"
    detail_lower = detail.lower()
    if "infra" in detail_lower:
        return "Infra"
    elif "active(bss)" in detail_lower:
        return "Active(BSS)"
    elif "oem not added" in detail_lower:
        return "OEM Not included yet"
    else:
        return "NO RCA/Not conclusive"

combined_df["RCA conclusive"] = combined_df["RCA detail"].apply(classify_rca_conclusive)

combined_df['Tech down'] = combined_df['TechG'].apply(lambda x: "2G Only" if x in ["N2G", "N2G,N2G"] else "2G+others")

# Filtering MS1 and active MS2
df_ms1_sites = combined_df[combined_df['CBS Status (MS1 / MS2)'].str.contains('MS1|MS0|skipped', na=False, case=False)]
df_active_ms2 = combined_df[~combined_df['CBS Status (MS1 / MS2)'].str.contains('MS1', na=False, case=False)]
df_active_ms2 = df_active_ms2[~df_active_ms2['Status'].str.contains('locked|theft|lock', na=False, case=False)]

# Grouping examples with logging
log_print("Grouping Total incidents by Circle...")
df_total_inc = combined_df.groupby('Circle').size().reset_index(name='Total Incidents')

log_print("Grouping Total SDO incidents by Circle...")
df_sdo = combined_df[combined_df[mttr_col].str.strip() == "SDO"]
df_sdo_count = df_sdo.groupby('Circle')[mttr_col].count().reset_index(name="Total SDO")

log_print("Grouping RCA from External Alarm...")
df_auto_rca_ext = combined_df[
    (combined_df["Auto RCA - From Ext alarm"].notna()) &
    (combined_df["Auto RCA - From Ext alarm"].str.strip() != "") &
    (combined_df["Auto RCA - From Ext alarm"].str.strip().str.lower() != "no alarm")
]
df_auto_rca_ext_count = df_auto_rca_ext.groupby('Circle').size().reset_index(name='RCA From External Alarm')

log_print("Grouping RCA from Internal Alarm...")
df_auto_rca_int = combined_df[
    (combined_df["Auto RCA - From Internal alarm"].notna()) &
    (combined_df["Auto RCA - From Internal alarm"].str.strip() != "") &
    (combined_df["Auto RCA - From Internal alarm"].str.strip().str.lower() != "no alarm")
]
df_auto_rca_int_count = df_auto_rca_int.groupby('Circle').size().reset_index(name='RCA From Internal Alarm')

log_print("Grouping RCA from Transmission Equipment...")
df_auto_rca_txeq = combined_df[
    (combined_df["Auto RCA - From Transmission Equipment"].notna()) &
    (combined_df["Auto RCA - From Transmission Equipment"].str.strip() != "") &
    (combined_df["Auto RCA - From Transmission Equipment"].str.strip().str.lower() != "no alarm")
]
df_auto_rca_txeq_count = df_auto_rca_txeq.groupby('Circle').size().reset_index(name='RCA From Transmission Equipment')

# Total Filled RCA
log_print("Grouping Total Filled RCA...")
total_filled_rca_df = combined_df[
    ((combined_df["Auto RCA - From Ext alarm"].notna()) & (combined_df["Auto RCA - From Ext alarm"] != "No Alarm")) |
    ((combined_df["Auto RCA - From Internal alarm"].notna()) & (combined_df["Auto RCA - From Internal alarm"] != "No Alarm")) |
    ((combined_df["Auto RCA - From Transmission Equipment"].notna()) & (combined_df["Auto RCA - From Transmission Equipment"] != "No Alarm"))
]
total_filled_rca_df_count = total_filled_rca_df.groupby('Circle').size().reset_index(name='Total Filled RCA')

# Prepare TT mapped dataframes
tt_mapped_df = combined_df[combined_df["TT ID"].notna() & (combined_df["TT ID"].astype(str).str.strip() != "")]

combined_df["Duration of Cell Outage Minutes"] = pd.to_numeric(combined_df["Duration of Cell Outage Minutes"], errors='coerce')

log_print("Grouping incidents less than 15 minutes...")
df_less_than_15_all = combined_df[combined_df["Duration of Cell Outage Minutes"] < 15]
df_less_than_15_all_count = df_less_than_15_all.groupby('Circle').size().reset_index(name='Total Incidents Less Than 15 Minutes')

log_print("Grouping incidents greater than or equal to 15 minutes...")
df_greater_than_15_all = combined_df[combined_df["Duration of Cell Outage Minutes"] >= 15]
df_greater_than_15_all_count = df_greater_than_15_all.groupby('Circle').size().reset_index(name='Total Incidents Greater Than equal to 15 Minutes')

log_print("Grouping TT Mapping less than 15 minutes...")
df_less_than_15 = tt_mapped_df[tt_mapped_df["Duration of Cell Outage Minutes"] < 15]
df_less_than_15_count = df_less_than_15.groupby('Circle').size().reset_index(name='TT Mapping less than 15 minute')

log_print("Grouping TT Mapping greater than or equal to 15 minutes...")
df_greater_than_15 = tt_mapped_df[tt_mapped_df["Duration of Cell Outage Minutes"] >= 15]
df_greater_than_15_count = df_greater_than_15.groupby('Circle').size().reset_index(name='TT Mapping greater than equal to 15 minute')

log_print("Grouping Infra RCA...")
df_infra = combined_df[combined_df["RCA conclusive"] == "Infra"]
df_infra_count = df_infra.groupby('Circle').size().reset_index(name='Infra')

log_print("Grouping Active(BSS) RCA...")
df_active_bss = combined_df[combined_df["RCA conclusive"] == "Active(BSS)"]
df_active_bss_count = df_active_bss.groupby('Circle').size().reset_index(name='Active(BSS)')

log_print("Grouping NO RCA/Not conclusive...")
df_no_rca = combined_df[combined_df["RCA conclusive"] == "NO RCA/Not conclusive"]
df_no_rca_count = df_no_rca.groupby('Circle').size().reset_index(name='NO RCA/Not conclusive')

log_print("Grouping OEM Not Included Yet RCA...")
df_oem_not_included = combined_df[combined_df["RCA conclusive"] == "OEM Not included yet"]
df_oem_not_included_count = df_oem_not_included.groupby('Circle').size().reset_index(name='OEM Not included yet')

# Date formatting for outage duration columns
formatted_date = combined_df['Date'].max().strftime('%d %b')

combined_df['Total Duration of Outage (Minutes) - All Tech'] = pd.to_numeric(
    combined_df['Total Duration of Outage (Minutes) - All Tech'], errors='coerce'
)

log_print("Grouping total outage duration by Circle and Physical ID...")
df_all_tech_outage = combined_df.groupby(['Circle', 'Physical ID'], as_index=False)[
    'Total Duration of Outage (Minutes) - All Tech'
].sum()

df_all_tech_outage.rename(
    columns={'Total Duration of Outage (Minutes) - All Tech': f'{formatted_date}'},
    inplace=True
)

log_print("Grouping activity count by Circle and Physical ID...")
df_activity_count = combined_df.groupby(['Circle', 'Physical ID']).size().reset_index(name='Count')
df_activity_count.rename(columns={'Count': f'{formatted_date}'}, inplace=True)

log_print("Grouping total outage duration by Circle...")
df_all_tech_outage_circle = combined_df.groupby(['Circle'], as_index=False)[
    'Total Duration of Outage (Minutes) - All Tech'
].sum()

df_all_tech_outage_circle.rename(
    columns={'Total Duration of Outage (Minutes) - All Tech': f'{formatted_date}'},
    inplace=True
)

total_sum = df_all_tech_outage_circle[f'{formatted_date}'].sum()
total_row = pd.DataFrame({'Circle': ['Total'], f'{formatted_date}': [total_sum]})
df_all_tech_outage_circle = pd.concat([df_all_tech_outage_circle, total_row], ignore_index=True)

log_print("Grouping total outage duration for active MS2 by Circle...")
df_all_tech_outage_circle_active = df_active_ms2.groupby(['Circle'], as_index=False)[
    'Total Duration of Outage (Minutes) - All Tech'
].sum()

# Rename column using formatted date
log_print("Renaming active MS2 outage duration column with formatted date...")
df_all_tech_outage_circle_active.rename(
    columns={'Total Duration of Outage (Minutes) - All Tech': f'{formatted_date}'},
    inplace=True
)

# Calculate total sum of outage durations across all circles
log_print("Calculating total outage duration for active MS2 across all circles...")
total_sum = df_all_tech_outage_circle_active[f'{formatted_date}'].sum()

# Create a Total row as a DataFrame
log_print("Appending Total row to active MS2 outage DataFrame...")
total_row = pd.DataFrame({
    'Circle': ['Total'],
    f'{formatted_date}': [total_sum]
})

df_all_tech_outage_circle_active = pd.concat([df_all_tech_outage_circle_active, total_row], ignore_index=True)

log_print(f"✅ Total outage duration (active MS2): {total_sum}")


#===============================================Matching TT Numbers========================================

log_print("Filtering Ericsson TT records for matched TT numbers with Outage/Down type...")
df_matched_tt = df_ericsson_tt[
    (df_ericsson_tt['Final TT Status'] != 'Not matched') &
    (df_ericsson_tt['Outage Type'] == 'Outage/Down')
]

log_print(f"✅ Total matched TTs before removing duplicates: {df_matched_tt.shape[0]}")

df_matched_tt = df_matched_tt.drop_duplicates(subset='Ericsson Incident Number', keep='first')

log_print(f"✅ Total matched TTs after removing duplicates: {df_matched_tt.shape[0]}")

log_print("Grouping matched TTs by Circle...")
df_matched_tt_count = df_matched_tt.groupby('Circle').size().reset_index(name='TTs found in NGOR')

log_print(f"✅ Total TTs found in NGOR: {df_matched_tt_count['TTs found in NGOR'].sum()}")


tt_mapped_bucket = df_matched_tt.pivot_table(
    index=['Circle'],
    columns='Final TT Status',
    values='Ericsson Incident Number',
    aggfunc='count',
    fill_value=0,
    margins=True,
    margins_name='Total'
).reset_index()

log_print(f"✅ Pivot table created with {tt_mapped_bucket.shape[0]} rows.")

log_print("Merging outage TT count with matched TT status pivot table...")

tt_mapped_bucket = df_outage_tt_count.merge(tt_mapped_bucket, on='Circle', how='left').fillna(0)

log_print(f"✅ Merged DataFrame shape: {tt_mapped_bucket.shape[0]} rows and {tt_mapped_bucket.shape[1]} columns.")


log_print("🔍 Searching for yesterday's dashboard file in Sitewise_CMO_Trend...")

folder_path_yesterday = r'D:\Automation\One_Outage\Output\Ngor_Dashboard'
file_list_yesterday = glob.glob(os.path.join(folder_path_yesterday, '*.xlsx'))
valid_files = [f for f in file_list_yesterday if not os.path.basename(f).startswith('~$')]
valid_files = sorted(valid_files, key=os.path.getmtime, reverse=True)

for file_path in valid_files:
    try:
        if os.path.getsize(file_path) < 1000:
            log_print(f"⚠️ Skipping small or empty file: {os.path.basename(file_path)}", "warning")
            continue

        log_print(f"📄 Reading file: {file_path}")
        df_daily_sum_yesterday = pd.read_excel(file_path, sheet_name='Sitewise_CMO_Trend', engine='openpyxl')
        log_print(f"✅ Successfully read: {os.path.basename(file_path)}")
        break
    except Exception as e:
        log_print(f"❌ Failed to read {os.path.basename(file_path)}: {e}", "error")
        continue
else:
    raise FileNotFoundError("❌ No valid Excel file found in yesterday's dashboard folder.")

log_print("🗑️ Dropping unwanted columns from Sitewise_CMO_Trend...")
df_daily_sum_yesterday = df_daily_sum_yesterday.drop(df_daily_sum_yesterday.columns[2], axis=1)
df_daily_sum_yesterday = df_daily_sum_yesterday.drop(columns=['Total', 'Days ≥1000', 'Airtel ID'], errors='ignore')

log_print("🗑️ Dropping rows where all date columns are 0...")
date_columns = df_daily_sum_yesterday.columns[2:]
df_daily_sum_yesterday = df_daily_sum_yesterday.loc[~(df_daily_sum_yesterday[date_columns] == 0).all(axis=1)]



log_print("🔄 Merging yesterday's data with today's sitewise outage summary...")
df_daily_cmo = df_daily_sum_yesterday.merge(
    df_all_tech_outage,
    on=['Circle', 'Physical ID'],
    how='outer'
).fillna(0)


df_daily_cmo['Total'] = df_daily_cmo.iloc[:, 2:].sum(axis=1)
date_columns = df_daily_cmo.columns.difference(['Circle', 'Physical ID', 'Total'])
df_daily_cmo['Days ≥1000'] = (df_daily_cmo[date_columns] >= 1000).sum(axis=1)

df_daily_cmo['Airtel ID'] = df_daily_cmo['Physical ID'].apply(
    lambda x: x[2:8] if isinstance(x, str) and len(x) == 15 else x
)



log_print("📊 Processing completed for Sitewise_CMO_Trend.")


#====================== Circlewise_CMO_Trend ======================

log_print("📄 Reading Circlewise_CMO_Trend sheet from the same file...")
try:
    df_daily_sum_yesterday_circle = pd.read_excel(file_path, sheet_name='Circlewise_CMO_Trend', engine='openpyxl')
    df_daily_cmo_circle = df_daily_sum_yesterday_circle.merge(
        df_all_tech_outage_circle,
        on=['Circle'],
        how='left'
    ).fillna(0)
    log_print("📊 Processing completed for Circlewise_CMO_Trend.")
except Exception as e:
    log_print(f"❌ Failed to process Circlewise_CMO_Trend: {e}", "error")


#====================== Sitewise_Activity_Count ======================

log_print("📄 Reading Sitewise_incidents_Trend sheet from the same file...")
try:
    df_circlewise_activity_count = pd.read_excel(file_path, sheet_name='Sitewise_incidents_Trend', engine='openpyxl')
    df_circlewise_activity_count = df_circlewise_activity_count.drop(df_circlewise_activity_count.columns[2], axis=1)

    log_print("🗑️ Dropping rows where all date columns are 0...")
    date_columns = df_circlewise_activity_count.columns[2:]
    df_circlewise_activity_count = df_circlewise_activity_count.loc[~(df_circlewise_activity_count[date_columns] == 0).all(axis=1)]

    df_circlewise_activity_count = df_circlewise_activity_count.merge(
        df_activity_count,
        on=['Circle', 'Physical ID'],
        how='outer'
    ).fillna(0)

    log_print("📊 Processing completed for Sitewise_Activity_Count.")
except Exception as e:
    log_print(f"❌ Failed to process Sitewise_Activity_Count: {e}", "error")


log_print("📄 Reading Circlewise_RNA_Trend sheet...")
try:
    df_daily_rna_yesterday_circle = pd.read_excel(file_path, sheet_name='Circlewise_RNA_Trend', engine='openpyxl')
    df_cell_count = df_daily_rna_yesterday_circle[['Circle', 'Total Cells']]
    df_all_tech_rna_circle = df_cell_count.merge(df_all_tech_outage_circle, on='Circle', how='left').fillna(0)

    df_all_tech_rna_circle['RNA %'] = (
        ((df_all_tech_rna_circle['Total Cells'] * 1440) - df_all_tech_rna_circle[formatted_date]) /
        (df_all_tech_rna_circle['Total Cells'] * 1440)
    ) * 100

    df_all_tech_rna_circle.drop(columns=[formatted_date, "Total Cells"], inplace=True)
    df_all_tech_rna_circle.rename(columns={'RNA %': formatted_date}, inplace=True)

    df_rna = df_daily_rna_yesterday_circle.merge(df_all_tech_rna_circle, on='Circle', how='left').fillna(0)
    log_print("📊 Processing completed for Circlewise_RNA_Trend.")
except Exception as e:
    log_print(f"❌ Failed to process Circlewise_RNA_Trend: {e}", "error")


log_print("📄 Reading Circlewise_RNA_Trend_Excluding sheet...")
try:
    df_daily_rna_yesterday_circle_ms2 = pd.read_excel(file_path, sheet_name='Circlewise_RNA_Trend_Excluding', engine='openpyxl')
    df_cell_count_ms2 = df_daily_rna_yesterday_circle_ms2[['Circle', 'Total Cells']]
    df_all_tech_rna_circle_ms2 = df_cell_count_ms2.merge(df_all_tech_outage_circle_active, on='Circle', how='left').fillna(0)

    df_all_tech_rna_circle_ms2['RNA %'] = (
        ((df_all_tech_rna_circle_ms2['Total Cells'] * 1440) - df_all_tech_rna_circle_ms2[formatted_date]) /
        (df_all_tech_rna_circle_ms2['Total Cells'] * 1440)
    ) * 100

    df_all_tech_rna_circle_ms2.drop(columns=[formatted_date, "Total Cells"], inplace=True)
    df_all_tech_rna_circle_ms2.rename(columns={'RNA %': formatted_date}, inplace=True)

    df_rna_ms2 = df_daily_rna_yesterday_circle_ms2.merge(df_all_tech_rna_circle_ms2, on='Circle', how='left').fillna(0)
    log_print("📊 Processing completed for Circlewise_RNA_Trend (excluding).")
except Exception as e:
    log_print(f"❌ Failed to process Circlewise_RNA_Trend_Excluding: {e}", "error")





log_print("📄 Processing RNA excluding inactive sites...")
try:
    df_all_tech_rna_circle_active = df_cell_count.merge(df_all_tech_outage_circle_active, on='Circle', how='left').fillna(0)

    df_all_tech_rna_circle_active['RNA %'] = (
        ((df_all_tech_rna_circle_active['Total Cells'] * 1440) - df_all_tech_rna_circle_active[formatted_date]) /
        (df_all_tech_rna_circle_active['Total Cells'] * 1440)
    )

    df_all_tech_rna_circle_active.drop(columns=[formatted_date, "Total Cells"], inplace=True)
    df_all_tech_rna_circle_active.rename(columns={'RNA %': formatted_date}, inplace=True)

    df_rna_active = df_daily_rna_yesterday_circle.merge(df_all_tech_rna_circle_active, on='Circle', how='left').fillna(0)
    log_print("📊 Processing completed for RNA excluding inactive sites.")
except Exception as e:
    log_print(f"❌ Failed to process RNA excluding inactive sites: {e}", "error")


log_print("📊 Preparing Dashboard DataFrame...")
try:
    df_dashboard = df_circles.merge(df_total_inc, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_sdo_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_auto_rca_ext_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_auto_rca_int_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_auto_rca_txeq_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(total_filled_rca_df_count, on='Circle', how='left').fillna(0)
    df_dashboard['Not Filled RCA'] = df_dashboard['Total Incidents'] - df_dashboard['Total Filled RCA']
    df_dashboard = df_dashboard.merge(df_less_than_15_all_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_greater_than_15_all_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_less_than_15_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_greater_than_15_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_outage_tt_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_matched_tt_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_infra_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_active_bss_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_no_rca_count, on='Circle', how='left').fillna(0)
    df_dashboard = df_dashboard.merge(df_oem_not_included_count, on='Circle', how='left').fillna(0)

    total_row = df_dashboard.select_dtypes(include="number").sum().to_frame().T
    total_row[df_dashboard.columns[0]] = "Total"
    total_row = total_row[df_dashboard.columns]
    df_dashboard = pd.concat([df_dashboard, total_row], ignore_index=True)

    def calc_percent(numerator, denominator):
        return f"{round((numerator / denominator) * 100, 2)}%" if denominator and not pd.isna(denominator) else "0.0%"

    df_dashboard["RCA External %"] = df_dashboard.apply(lambda row: calc_percent(row['RCA From External Alarm'], row['Total Incidents']), axis=1)
    df_dashboard["RCA Internal %"] = df_dashboard.apply(lambda row: calc_percent(row['RCA From Internal Alarm'], row['Total Incidents']), axis=1)
    df_dashboard["RCA Transmission %"] = df_dashboard.apply(lambda row: calc_percent(row['RCA From Transmission Equipment'], row['Total Incidents']), axis=1)
    df_dashboard["Total Filled RCA %"] = df_dashboard.apply(lambda row: calc_percent(row['Total Filled RCA'], row['Total Incidents']), axis=1)
    df_dashboard["TT Mapping less than 15 minute %"] = df_dashboard.apply(lambda row: calc_percent(row['TT Mapping less than 15 minute'], row['Total Incidents Less Than 15 Minutes']), axis=1)
    df_dashboard["TT Mapping greater than equal to 15 minute %"] = df_dashboard.apply(lambda row: calc_percent(row['TT Mapping greater than equal to 15 minute'], row['Total Incidents Greater Than equal to 15 Minutes']), axis=1)
    df_dashboard["TT created and Mapped %"] = df_dashboard.apply(lambda row: calc_percent(row['TTs found in NGOR'], row['Outage Category TTs raised']), axis=1)
    df_dashboard["Infra %"] = df_dashboard.apply(lambda row: calc_percent(row['Infra'], row['Total Incidents']), axis=1)
    df_dashboard["Active(BSS) %"] = df_dashboard.apply(lambda row: calc_percent(row['Active(BSS)'], row['Total Incidents']), axis=1)
    df_dashboard["NO RCA/Not conclusive %"] = df_dashboard.apply(lambda row: calc_percent(row['NO RCA/Not conclusive'], row['Total Incidents']), axis=1)
    df_dashboard["OEM Not included yet %"] = df_dashboard.apply(lambda row: calc_percent(row['OEM Not included yet'], row['Total Incidents']), axis=1)

    log_print("📊 Dashboard DataFrame processing completed successfully.")
except Exception as e:
    log_print(f"❌ Failed to process Dashboard DataFrame: {e}", "error")




log_print("📄 Reading RBS_Type Excel file...")
try:
    rbs_type_file_path = r'D:\Automation\One_Outage\RBS_Type.xlsx'
    df_rbs_type = pd.read_excel(rbs_type_file_path, engine='openpyxl')
    log_print("📊 Successfully read RBS_Type Excel file.")
except Exception as e:
    log_print(f"❌ Failed to read RBS_Type Excel file: {e}", "error")



log_print("📄 Starting daily CMO and Activity count trend processing...")

try:
    df_cmo_till_date = df_daily_cmo.drop(columns=['Total', 'Days ≥1000', 'Airtel ID'], errors='ignore')

    df_cmo_activity = df_cmo_till_date.merge(
        df_circlewise_activity_count,
        on=['Circle', 'Physical ID'],
        how='left',
        suffixes=('_cmo', '_inc')
    )

    cmo_cols = [col for col in df_cmo_activity.columns if col.endswith('_cmo')]
    selected_cmo_cols = cmo_cols[-7:]
    count_series = (df_cmo_activity[selected_cmo_cols] > 1000).sum(axis=1)
    df_cmo_activity['CMO Priority'] = np.where(count_series > 2, count_series, 0)

    col_br = df_cmo_activity.columns[63]
    df_cmo_activity[col_br] = pd.to_numeric(df_cmo_activity[col_br], errors='coerce')

    conditions = [
        (df_cmo_activity[col_br] > 10) & (df_cmo_activity[col_br] <= 20),
        (df_cmo_activity[col_br] > 20)
    ]
    choices = [">10 & <=20", ">20"]

    df_cmo_activity['Incident Priority'] = np.select(conditions, choices, default="")

    df_cmo_activity['CMO_Incident_Category'] = np.select(
        [
            (df_cmo_activity['CMO Priority'] > 2) & (df_cmo_activity['Incident Priority'].isin([">10 & <=20", ">20"])),
            (df_cmo_activity['CMO Priority'] > 0) & (df_cmo_activity['Incident Priority'] == ""),
            (df_cmo_activity['CMO Priority'] <= 2) & (df_cmo_activity['Incident Priority'].isin([">10 & <=20", ">20"])),
        ],
        [
            "Both-CMO&Incident",
            "High CMO Only",
            "High Incidents Only"
        ],
        default=""
    )

    try:
        df_cmo_activity = df_cmo_activity.merge(
            df_rbs_type[['Physical ID', 'Toco Name', 'RBS']],
            on='Physical ID',
            how='left'
        )
    except KeyError:
        df_cmo_activity = df_cmo_activity.merge(
            df_rbs_type[['Site_ID', 'Toco Name', 'RBS']],
            left_on='Physical ID',
            right_on='Site_ID',
            how='left'
        )

    df_cmo_activity[['Toco Name', 'RBS']] = df_cmo_activity[['Toco Name', 'RBS']].replace(['', None, np.nan], '#N/A')

    priority_sites_cmo = df_cmo_activity[df_cmo_activity['CMO_Incident_Category'].isin(["Both-CMO&Incident", "High Incidents Only"])]
    priority_sites_all = df_cmo_activity[df_cmo_activity['CMO_Incident_Category'].isin(["Both-CMO&Incident", "High Incidents Only", "High CMO Only"])]

    priority_sites_cmo_pivot = priority_sites_cmo.pivot_table(
        index='Circle',
        columns='Incident Priority',
        values='Physical ID',
        aggfunc='count',
        fill_value=0,
        margins=True,
        margins_name='Total'
    ).reset_index()

    priority_sites_all_pivot = priority_sites_cmo.pivot_table(
        index='Circle',
        columns='RBS',
        values='Physical ID',
        aggfunc='count',
        fill_value=0,
        margins=True,
        margins_name='Total'
    ).reset_index()

    priority_sites_cmo_inc = priority_sites_all.pivot_table(
        index='Circle',
        columns='CMO_Incident_Category',
        values='Physical ID',
        aggfunc='count',
        fill_value=0,
        margins=True,
        margins_name='Total'
    ).reset_index()

    df_cmo_inc_dashboard = priority_sites_cmo_pivot.merge(priority_sites_all_pivot, on='Circle', how='left')
    df_cmo_inc_dashboard = df_cmo_inc_dashboard.merge(priority_sites_cmo_inc, on='Circle', how='left')

    log_print("📊 Daily CMO and Activity count trend processed successfully.")

except Exception as e:
    log_print(f"❌ Failed in daily CMO and Activity count trend processing: {e}", "error")


log_print("📄 Starting Project Harmony processing...")

try:
    folder_path_harmony = r'D:\Automation\One_Outage\Harmony_Sites\Harmony Indus Sites & Airtel Owned _Physcial IDs.xlsx'

    # Reading the "Harmony Indus Site detail" sheet
    df_indus_harmony_sites = pd.read_excel(folder_path_harmony, sheet_name='Harmony Indus Site detail')
    log_print("✅ Successfully read 'Harmony Indus Site detail' sheet.")

     # Create helper 8-char columns
    df_indus_harmony_sites['Physical 8_Char'] = df_indus_harmony_sites['Physical ID'].str[:8]
    df_daily_cmo['Physical 8_Char'] = df_daily_cmo['Physical ID'].str[:8]


    df_indus_harmony_sites_matched_full = df_daily_cmo[df_daily_cmo['Physical ID'].isin(df_indus_harmony_sites['Physical ID'])]
    df_match_8char = df_daily_cmo[df_daily_cmo['Physical 8_Char'].isin(df_indus_harmony_sites['Physical 8_Char'])]

    # Combine both matches (union) and remove duplicates
    df_indus_harmony_sites_matched = pd.concat([df_indus_harmony_sites_matched_full, df_match_8char]).drop_duplicates()
      
    df_indus_harmony_sites_matched = df_indus_harmony_sites_matched.drop(columns=['Airtel ID','Physical 8_Char'], errors='ignore')
    df_indus_harmony_sites_matched = df_indus_harmony_sites_matched.merge(df_indus_harmony_sites, on='Physical ID', how='left')
    log_print(f"🔍 Indus Harmony matched sites: {len(df_indus_harmony_sites_matched)} records found.")

    # Reading the "Airtel Own_Esc Matrix" sheet
    df_airtel_owned_sites = pd.read_excel(folder_path_harmony, sheet_name='Airtel Own_Esc Matrix')
    log_print("✅ Successfully read 'Airtel Own_Esc Matrix' sheet.")

    df_airtel_owned_sites['Physical 8_Char'] = df_airtel_owned_sites['Physical ID'].str[:8]


    df_airtel_owned_sites_matched_full = df_daily_cmo[df_daily_cmo['Physical ID'].isin(df_airtel_owned_sites['Physical ID'])]

    df_match_8char_airtel = df_daily_cmo[df_daily_cmo['Physical 8_Char'].isin(df_airtel_owned_sites['Physical 8_Char'])]

    # Combine both matches (union) and remove duplicates
    df_airtel_owned_sites_matched = pd.concat([df_airtel_owned_sites_matched_full, df_match_8char_airtel]).drop_duplicates()





    df_airtel_owned_sites_matched = df_airtel_owned_sites_matched.drop(columns=['Airtel ID','Physical 8_Char'], errors='ignore')
    df_daily_cmo = df_daily_cmo.drop(columns=['Physical 8_Char'], errors='ignore')

    df_airtel_owned_sites_matched = df_airtel_owned_sites_matched.merge(df_airtel_owned_sites, on='Physical ID', how='left')
    log_print(f"🔍 Airtel Owned sites matched: {len(df_airtel_owned_sites_matched)} records found.")

    log_print("📊 Project Harmony processing completed successfully.")

except Exception as e:
    log_print(f"❌ Error during Project Harmony processing: {e}", "error")





#=====================================Project Relocation Sites trend=================================================




log_print("📄 Starting Project Relocation SItes trend......")

# try:
#     folder_path_relocation = r'D:\Automation\One_Outage\Relocated Sites\relocated_sites.xlsx'

#     # Reading the "Relocated Site detail" sheet
#     df_relocated_sites = pd.read_excel(folder_path_relocation, sheet_name='Relocated_Sites')
#     log_print("✅ Successfully read 'Relocated Site detail' sheet.")

#     df_relocated_sites_matched = df_daily_cmo[df_daily_cmo['Physical ID'].isin(df_relocated_sites['Physical ID'])]
      
#     df_relocated_sites_matched = df_relocated_sites_matched.drop(columns=['Airtel ID'], errors='ignore')
#     log_print(f"🔍 Relocated matched sites: {len(df_relocated_sites_matched)} records found.")

#     log_print("📊 Project Harmony processing completed successfully.")

# except Exception as e:
#     log_print(f"❌ Error during Relocated Sites processing: {e}", "error")


try:
    folder_path_relocation = r'D:\Automation\One_Outage\Relocated Sites\relocated_sites.xlsx'

    # Reading the "Relocated_Sites" sheet
    df_relocated_sites = pd.read_excel(folder_path_relocation, sheet_name='Relocated_Sites')
    log_print("✅ Successfully read 'Relocated_Sites' sheet.")

    # Ensure Physical ID is string
    df_daily_cmo['Physical ID'] = df_daily_cmo['Physical ID'].astype(str)
    df_relocated_sites['Physical ID'] = df_relocated_sites['Physical ID'].astype(str)

    # Create helper 8-char columns
    df_daily_cmo['Physical 8_Char'] = df_daily_cmo['Physical ID'].str[:8]
    df_relocated_sites['Physical 8_Char'] = df_relocated_sites['Physical ID'].str[:8]

    # Match using full Physical ID
    df_match_full = df_daily_cmo[df_daily_cmo['Physical ID'].isin(df_relocated_sites['Physical ID'])]

    # Match using 8-char Physical ID
    df_match_8char = df_daily_cmo[df_daily_cmo['Physical 8_Char'].isin(df_relocated_sites['Physical 8_Char'])]

    # Combine both matches (union) and remove duplicates
    df_relocated_sites_matched = pd.concat([df_match_full, df_match_8char]).drop_duplicates()

    # Drop Airtel ID if present
    df_relocated_sites_matched = df_relocated_sites_matched.drop(columns=['Airtel ID','Physical 8_Char'], errors='ignore')
    df_daily_cmo = df_daily_cmo.drop(columns=['Physical 8_Char'], errors='ignore')

    log_print(f"🔍 Relocated matched sites: {len(df_relocated_sites_matched)} records found.")

    log_print("📊 Project Relocation SItes processing completed successfully.")

except Exception as e:
    log_print(f"❌ Error during Relocated Sites processing: {e}", "error")





#===============================================================================





# ===============================USD, SDO, LDI Sites count======================
log_print("📄 Starting USD, SDO, LDI Sites count trend processing...")
log_print("Grouping Total Unique Sites by Circle...")
df_usd_unique_count = combined_df.groupby('Circle')['Physical ID'].nunique().reset_index(name='Unique Sites Down')

log_print("Grouping Total SDO incidents by Circle...")
df_sdo = combined_df[combined_df[mttr_col].str.strip() == "SDO"]
df_sdo_unique_count = df_sdo.groupby('Circle')['Physical ID'].nunique().reset_index(name="Total SDI")


log_print("Grouping Total LDI incidents by Circle...")
df_ldi = combined_df[combined_df[mttr_col].str.strip() != "SDO"]
df_ldi_unique_count = df_ldi.groupby('Circle')['Physical ID'].nunique().reset_index(name="ISD Excluding SDI")


dashboard_usd_sdo_ldi = df_circles.merge(df_usd_unique_count, on='Circle', how='left').fillna(0)
dashboard_usd_sdo_ldi = dashboard_usd_sdo_ldi.merge(df_sdo_unique_count, on='Circle', how='left').fillna(0)
dashboard_usd_sdo_ldi = dashboard_usd_sdo_ldi.merge(df_ldi_unique_count, on='Circle', how='left').fillna(0)

#===============================Output Excel Files==========================




output_folder_ngor = r'D:\Automation\One_Outage\Output\Ngor_Dashboard'
output_folder_harmony = r'D:\Automation\One_Outage\Output\harmony_sites'
output_folder_ms1 = r'D:\Automation\One_Outage\Output\ms1_sites'
output_folder_daily_combined = r'D:\Automation\One_Outage\Output\Daily_Combined_Dump'
output_folder_priority_sites = r'D:\Automation\One_Outage\Output\Priority_Sites'

# Create output folder if it doesn't exist
os.makedirs(output_folder_ngor, exist_ok=True)
os.makedirs(output_folder_harmony, exist_ok=True)
os.makedirs(output_folder_ms1, exist_ok=True)
os.makedirs(output_folder_daily_combined, exist_ok=True)
os.makedirs(output_folder_priority_sites, exist_ok=True)

try:
    output_file_priority_sites = os.path.join(output_folder_priority_sites, f'Priority_Sites_{report_date_one_outage}.xlsx')
    with pd.ExcelWriter(output_file_priority_sites, engine='xlsxwriter') as writer: 
        df_cmo_inc_dashboard.to_excel(writer, sheet_name="Summary", index=False) 
        priority_sites_all.to_excel(writer, sheet_name="Priority Sites Dump", index=False)
        dashboard_usd_sdo_ldi.to_excel(writer, sheet_name="USD_SDO_LDI_Sites", index=False)
    log_print(f"💾 Excel file saved successfully at: {output_file_priority_sites}")
except Exception as e:
    log_print(f"❌ Failed to save Priority Sites Excel file: {e}", "error")

try:
    output_file = os.path.join(output_folder_ngor, f'NGOR_Dashboard_{report_date_one_outage}.xlsx')
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_dashboard.to_excel(writer, sheet_name="Dashboard", index=False)
        tt_mapped_bucket.to_excel(writer, sheet_name="TT_Mapped_Bucket", index=False)
        #df_cmo_inc_dashboard.to_excel(writer, sheet_name="Priority_Sites_CMO", index=False)
        df_daily_cmo_circle.to_excel(writer, sheet_name="Circlewise_CMO_Trend", index=False)
        df_circlewise_activity_count.to_excel(writer, sheet_name="Sitewise_incidents_Trend", index=False)
        df_rna.to_excel(writer, sheet_name="Circlewise_RNA_Trend", index=False)
        df_rna_ms2.to_excel(writer, sheet_name="Circlewise_RNA_Trend_Excluding", index=False)    
        df_daily_cmo.to_excel(writer, sheet_name="Sitewise_CMO_Trend", index=False) 
        df_cmo_activity.to_excel(writer, sheet_name="CMO_Incident_Count_Trend", index=False) 
    log_print(f"💾 Excel file saved successfully at: {output_file}")
except Exception as e:
    log_print(f"❌ Failed to save NGOR Dashboard Excel file: {e}", "error")

try:
    output_file_daily_combined = os.path.join(output_folder_daily_combined, f'NGOR_Dashboard_Dump_{report_date_one_outage}.xlsx')
    with pd.ExcelWriter(output_file_daily_combined, engine='xlsxwriter') as writer: 
        combined_df.to_excel(writer, sheet_name="One_Outage", index=False) 
        df_ericsson_tt.to_excel(writer, sheet_name= "Ericsson_TT", index=False)       
        combined_df_Locked_LODS.to_excel(writer, sheet_name= "Locked_LODS", index=False)
        combined_df_not_in_site_db.to_excel(writer, sheet_name= "Not_in_Site_DB", index=False)
    log_print(f"💾 Excel file saved successfully at: {output_file_daily_combined}")
except Exception as e:
    log_print(f"❌ Failed to save Daily Combined Dump Excel file: {e}", "error")

try:
    output_file_harmony = os.path.join(output_folder_harmony, f'Harmony_Sites_{report_date_one_outage}.xlsx')
    with pd.ExcelWriter(output_file_harmony, engine='xlsxwriter') as writer:
        df_indus_harmony_sites_matched.to_excel(writer, sheet_name="Harmony Indus Sites", index=False)
        df_airtel_owned_sites_matched.to_excel(writer, sheet_name="Airtel Owned Sites", index=False)
    log_print(f"💾 Excel file saved successfully at: {output_file_harmony}")
except Exception as e:
    log_print(f"❌ Failed to save Harmony Sites Excel file: {e}", "error")

try:
    output_file_ms1_sites = os.path.join(output_folder_ms1, f'MS1_Sites_{report_date_one_outage}.xlsx')
    with pd.ExcelWriter(output_file_ms1_sites, engine='xlsxwriter') as writer:
        df_ms1_sites.to_excel(writer, sheet_name="MS1 Sites", index=False)
        df_relocated_sites_matched.to_excel(writer, sheet_name="Relocated Sites Trend", index=False)
    log_print(f"💾 Excel file saved successfully at: {output_file_ms1_sites}")
except Exception as e:
    log_print(f"❌ Failed to save MS1 Sites Excel file: {e}", "error")



#==========================Sending Mails=========================

from_email = "IN_D_Engineering_NIfi_Reports@airtel.com"
today_date = datetime.now().strftime("%d%m%Y")
yesterday_date = (datetime.now() - timedelta(days=1)).strftime("%d%m%Y")

# Initialize Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
except Exception as e:
    log_print(f"❌ Error initializing Outlook: {e}", "error")
    sys.exit()

# ==================== Reusable Mail Function ====================
def send_email(recipient_file, subject, body, attachment_file):
    try:
        receipient_df = pd.read_excel(recipient_file)
        to_recipients = receipient_df["to"].dropna().tolist()
        cc_recipients = receipient_df["cc"].dropna().tolist()

        to_list = ";".join(to_recipients)
        cc_list = ";".join(cc_recipients)

        mail = outlook.CreateItem(0)
        mail.To = to_list
        mail.CC = cc_list
        mail.Subject = subject
        mail.Body = body
        mail.SentOnBehalfOfName = from_email

        matched_account = None
        for account in outlook.Session.Accounts:
            if account.SmtpAddress.lower() == from_email.lower():
                matched_account = account
                break

        if matched_account:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, matched_account))
        else:
            log_print(f"⚠️ Warning: Sending account '{from_email}' not found. Using default account.", "warning")

        if os.path.exists(attachment_file):
            mail.Attachments.Add(attachment_file)
        else:
            log_print(f"⚠️ Attachment not found: {attachment_file}", "warning")

        mail.Send()
        log_print(f"📧 Email sent successfully to: {to_list}")
    except Exception as e:
        log_print(f"❌ Error sending email: {e}", "error")


# ==================== Sending Harmony Sites Mail ====================
harmony_recipients = r"D:\Automation\One_Outage\Receipient_Harmony_Sites.xlsx"
harmony_subject = f"Harmony Sites CMO Trend Till- {report_date_one_outage}"
harmony_body = f"""
Hi Team,

Please find Harmony Sites CMO Trend from 26-05-2025 to {report_date_one_outage}.

Regards,  
Central Team FSO
"""

send_email(harmony_recipients, harmony_subject, harmony_body, output_file_harmony)
log_print("✅ Email process completed successfully for Harmony Sites")

# ==================== Sending MS1 Sites Mail ====================
ms1_recipients = r"D:\Automation\One_Outage\Receipient_MS1_Sites.xlsx"
ms1_subject = f"MS1_Sites_CMO_Contribution- {report_date_one_outage}"
ms1_body = f"""
Hi Team,

Please find MS1 Sites CMO Contribution of: {report_date_one_outage}.

Regards,  
Central Team FSO
"""

send_email(ms1_recipients, ms1_subject, ms1_body, output_file_ms1_sites)
log_print("✅ Email process completed successfully for MS1 Sites")


# ==================== Sending Priority Sites Mail ====================
daily_cmo_recipients = r"D:\Automation\One_Outage\Receipient_Priority_Sites.xlsx"
daily_cmo_subject = f"Daily CMO trend- {report_date_one_outage}"
daily_cmo_body = f"""
Hi Team,

Please find Priority Sites Summary of: {report_date_one_outage}.

Regards,  
Central Team FSO
"""

send_email(daily_cmo_recipients, daily_cmo_subject, daily_cmo_body, output_file_priority_sites)
log_print("✅ Email process completed successfully for Priority Sites")


# Release Outlook objects
del outlook, namespace