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


#=============================
# CONFIGURATION
#=============================
SAVE_PATH_TT = r"D:\Ericsson TT Daily Report"
SAVE_PATH_NGOR = r"D:\NGOR_Report"
OUTLOOK_FOLDER = "Required to Work"
SENDER_NGOR = "performance.management@ericsson.com"

#=============================
# SETUP
#=============================
os.makedirs(SAVE_PATH_TT, exist_ok=True)
os.makedirs(SAVE_PATH_NGOR, exist_ok=True)

try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
except Exception as e:
    print(f"❌ Error connecting to Outlook: {e}")
    sys.exit(1)

try:
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    target_folder = inbox.Folders(OUTLOOK_FOLDER)
except Exception as e:
    print(f"❌ Error accessing '{OUTLOOK_FOLDER}' folder: {e}")
    sys.exit(1)

messages = target_folder.Items
messages.Sort("[ReceivedTime]", True)

#=============================
# Set Filter Date — 19th July 2025
#=============================
filter_date = datetime(2025, 11, 8).date()

#=============================
# PROCESS ERICSSON DAILY REPORT
#=============================
print("🔍 Checking for 'Ericsson Daily Report' emails on 19th July...")

found_tt = False

for msg in messages:
    try:
        if msg.ReceivedTime.date() != filter_date:
            continue  # Skip other dates

        if "Ericsson Daily Report" in msg.Subject:
            print(f"✅ Found TT Email: {msg.Subject} | {msg.ReceivedTime.strftime('%H:%M')}")

            for att in msg.Attachments:
                if att.FileName.lower().endswith(".xlsx"):
                    base, ext = os.path.splitext(att.FileName)
                    date_str = msg.ReceivedTime.strftime('%Y%m%d_%H%M%S')
                    safe_name = re.sub(r'[\\/*?:"<>|]', "_", base)
                    final_name = f"{safe_name}_{date_str}{ext}"
                    att.SaveAsFile(os.path.join(SAVE_PATH_TT, final_name))
                    print(f"📂 Saved: {final_name}")
                    found_tt = True
    except Exception as e:
        print(f"⚠️ Error processing TT email: {e}")

if not found_tt:
    print("❌ No 'Ericsson Daily Report' emails found for 19th July.")

#=============================
# PROCESS NGOR REPORT
#=============================
print("\n🔍 Checking for 'NGOR Report' emails on 19th July...")

from collections import defaultdict

circle_emails = defaultdict(dict)

for msg in messages:
    try:
        if msg.ReceivedTime.date() != filter_date:
            continue

        sender_address = msg.SenderEmailAddress.lower().replace('smtp:', '')
        if sender_address == SENDER_NGOR.lower():
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
        print(f"⚠️ Error reading email: {e}")

found_ngor = False

for circle, records in circle_emails.items():
    msg_to_use = records.get("update") or records.get("plain")

    if msg_to_use:
        is_update = "update" in records
        status = "✅ Found Preferred Email" if is_update else "📂 Downloading Fallback Email"

        print(f"{status}: {msg_to_use.Subject} | {msg_to_use.ReceivedTime.strftime('%H:%M')}")

        try:
            for att in msg_to_use.Attachments:
                if att.FileName.lower().endswith(".xlsx"):
                    base, ext = os.path.splitext(att.FileName)
                    time_str = msg_to_use.ReceivedTime.strftime('%Y%m%d_%H%M%S')
                    safe_name = re.sub(r'[\\/*?:"<>|]', "_", base)
                    final_name = f"{safe_name}_{time_str}{ext}"

                    att.SaveAsFile(os.path.join(SAVE_PATH_NGOR, final_name))
                    print(f"📂 Saved: {final_name}")
                    found_ngor = True
        except Exception as e:
            print(f"⚠️ Error saving attachment for {circle}: {e}")

if not found_ngor:
    print("❌ No NGOR Report emails found for 19th July.")

print("\n🚀 Script completed.")

#===================================Defining Circle list=============================================

# List of Circle values
circle_list = [
    'AP', 'AS', 'BH', 'CH', 'DL', 'GJ', 'HP', 'HR', 'JK', 'KK', 'KL', 'KOL',
    'MH', 'MP', 'MU', 'NE', 'OR', 'PB', 'RJ', 'TN', 'UE', 'UW', 'WB'
]

# Create DataFrame
df_circles = pd.DataFrame({'Circle': circle_list})

#====================================Ericsson TT Daily Report========================================

# Folder path
folder_path = r'D:\Ericsson TT Daily Report'

# Get all .xlsx files in the folder
file_list = glob.glob(os.path.join(folder_path, '*.xlsx'))

# Sort files alphabetically (A-Z), then pick the last one
latest_file = sorted(file_list)[-1]

# Print the file name being extracted
print(f"📄 Reading file: {os.path.basename(latest_file)}")

# Read the latest file (starting from 4th row)
df_ericsson_tt = pd.read_excel(latest_file, header=3)

df_ericsson_tt['Create Date'] = df_ericsson_tt['Create Date Time of TT'].dt.date
df_ericsson_tt['Create Date'] = pd.to_datetime(df_ericsson_tt['Create Date'], errors='coerce')
report_date_ericsson_tt = df_ericsson_tt['Create Date'].max().strftime('%d-%m-%Y')

# df_ericsson_tt['Outage Type'] = df_ericsson_tt['Description'].str.contains(r'\b(Outage|Down)\b', case=False, na=False)
# df_ericsson_tt['Outage Type'] = df_ericsson_tt['Outage Type'].apply(lambda x: 'Outage/Down' if x else '')


# Condition 1: contains whole word "Outage" or "Down" (not "Downlink")
outage_down_condition = df_ericsson_tt['Description'].str.contains(r'(_OUTAGE|Outage|Down)', case=False, na=False)

# Condition 2: does NOT contain "L1LA member"
not_l1la_condition = ~df_ericsson_tt['Description'].str.contains(r'L1LA member', case=False, na=False)

# Condition 2: does NOT contain "Base Station Connectivity Down Minor"
base_stn_con = ~df_ericsson_tt['Description'].str.contains(r'Base Station Connectivity Down Minor', case=False, na=False)

# Condition 2: does NOT contain "QIA"
qia = ~df_ericsson_tt['Description'].str.contains(r'QIA', case=False, na=False)


# Condition 2: does NOT contain "AUTO ZTE 4G NVR-Carrier downlink data problem"
career_downlink = ~df_ericsson_tt['Description'].str.contains(r'AUTO ZTE 4G NVR-Carrier downlink data problem', case=False, na=False)

# Condition 2: does NOT contain "AUTO 5G OUTAGE-Service Degraded"
ser_degr = ~df_ericsson_tt['Description'].str.contains(r'AUTO 5G OUTAGE-Service Degraded', case=False, na=False)




# Apply both conditions
df_ericsson_tt['Outage Type'] = ((outage_down_condition) & (not_l1la_condition) & (base_stn_con) & (qia) & (career_downlink) & (ser_degr) ).apply(lambda x: 'Outage/Down' if x else '')
df_ericsson_tt.rename(columns={'Site Group/Circle': 'Circle'}, inplace=True)

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





#==============================================Reading Today received files (outage report)====================
# Input folder
input_folder_one_outage = r'D:\NGOR_Report'
# Target sheets and containers
sheet_to_df = {
    'Common Outage Format': [],
    'Locked_LODS': [],
    'Not in Site DB': [],
}

# Scan all Excel files in the folder
excel_files = [
    f for f in os.listdir(input_folder_one_outage)
    if f.lower().endswith(('.xlsx', '.xls'))
]

if not excel_files:
    print("🚫 No Excel files found in the folder.")
else:
    print(f"📁 Found {len(excel_files)} Excel files.\n")

    for file in excel_files:
        file_path = os.path.join(input_folder_one_outage, file)
        try:
            with pd.ExcelFile(file_path) as xls:
                for sheet_name in sheet_to_df.keys():
                    if sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        sheet_to_df[sheet_name].append(df)
                        print(f"✔️ Sheet '{sheet_name}' loaded from: {file}")
                    else:
                        print(f"⚠️ Sheet '{sheet_name}' not found in: {file}")
        except Exception as e:
            print(f"❌ Error reading {file}: {e}")

# Final combined DataFrames
combined_df = pd.concat(sheet_to_df['Common Outage Format'], ignore_index=True) if sheet_to_df['Common Outage Format'] else pd.DataFrame()
combined_df_Locked_LODS = pd.concat(sheet_to_df['Locked_LODS'], ignore_index=True) if sheet_to_df['Locked_LODS'] else pd.DataFrame()
combined_df_not_in_site_db = pd.concat(sheet_to_df['Not in Site DB'], ignore_index=True) if sheet_to_df['Not in Site DB'] else pd.DataFrame()




#print(f"\n✅ Combined file ready with {combined_df.shape[0]} rows")

# Clean column names
combined_df.columns = [col.strip() for col in combined_df.columns]

# Identify the MTTR column (handle inconsistent name)
mttr_col = [col for col in combined_df.columns if col.startswith("MTTR Category")][0]




common_outage_tt = combined_df[['Circle', 'TT ID']].copy()

# Step 1: Extract all TT Numbers as lists
common_outage_tt['TT Numbers'] = common_outage_tt['TT ID'].str.findall(r'INC\d+')

# Step 2: Explode the TT Numbers list into individual rows
common_outage_tt = common_outage_tt.explode('TT Numbers').dropna(subset=['TT Numbers']).reset_index(drop=True)
# Step 3: Remove duplicates
common_outage_tt = common_outage_tt.drop_duplicates(subset=['Circle', 'TT Numbers'])





common_outage_tt_lod = combined_df_Locked_LODS[['Circle', 'TT ID']].copy()

# Step 1: Extract all TT Numbers as lists
common_outage_tt_lod['TT Numbers'] = common_outage_tt_lod['TT ID'].str.findall(r'INC\d+')

# Step 2: Explode the TT Numbers list into individual rows
common_outage_tt_lod = common_outage_tt_lod.explode('TT Numbers').dropna(subset=['TT Numbers']).reset_index(drop=True)
# Step 3: Remove duplicates
common_outage_tt_lod = common_outage_tt_lod.drop_duplicates(subset=['Circle', 'TT Numbers'])




common_outage_tt_not_in_sdb = combined_df_not_in_site_db[['Circle', 'WorK Order ID']].copy()

# Step 1: Extract all TT Numbers as lists
common_outage_tt_not_in_sdb['TT Numbers'] = common_outage_tt_not_in_sdb['WorK Order ID'].str.findall(r'INC\d+')

# Step 2: Explode the TT Numbers list into individual rows
common_outage_tt_not_in_sdb = common_outage_tt_not_in_sdb.explode('TT Numbers').dropna(subset=['TT Numbers']).reset_index(drop=True)
# Step 3: Remove duplicates
common_outage_tt_not_in_sdb = common_outage_tt_not_in_sdb.drop_duplicates(subset=['Circle', 'TT Numbers'])







def assign_tt_status(df, column_to_check, tt_list_df, status_column_name):
    tt_set = set(tt_list_df['TT Numbers'].dropna().astype(str))
    df[status_column_name] = df[column_to_check].astype(str).apply(
        lambda x: 'Matched' if x in tt_set else 'Not matched'
    )

# Apply function for all TT comparisons
assign_tt_status(df_ericsson_tt, 'Ericsson Incident Number', common_outage_tt, 'TT status COR')
assign_tt_status(df_ericsson_tt, 'Ericsson Incident Number', common_outage_tt_lod, 'TT status LODS')
assign_tt_status(df_ericsson_tt, 'Ericsson Incident Number', common_outage_tt_not_in_sdb, 'TT status (Not in Site DB)')


def get_final_tt_status_combination(row):
    cor = row['TT status COR'] == 'Matched'
    lods = row['TT status LODS'] == 'Matched'
    not_sdb = row['TT status (Not in Site DB)'] == 'Matched'
    
    # All matched
    if cor and lods and not_sdb:
        return 'Matched in All'
    # Two-way combinations
    elif cor and lods:
        return 'Matched in COR & LODS'
    elif cor and not_sdb:
        return 'Matched in COR & Not in Site DB'
    elif lods and not_sdb:
        return 'Matched in LODS & Not in Site DB'
    # One-way matches
    elif cor:
        return 'Matched in COR'
    elif lods:
        return 'Matched in LODS'
    elif not_sdb:
        return 'Matched in Not in Site DB'
    else:
        return 'Not matched'

df_ericsson_tt['Final TT Status'] = df_ericsson_tt.apply(get_final_tt_status_combination, axis=1)

combined_df['TT Number'] = combined_df['TT ID'].str.extract(r'(INC\d+)')

# Ensure 'Date' column is datetime format
combined_df['Date'] = pd.to_datetime(combined_df['Date'], errors='coerce')

# Get latest date in 'dd-mm-yyyy' format
report_date_one_outage = combined_df['Date'].max().strftime('%d-%m-%Y')


print(f"Filling blank Physical ID from other columns...")


def fill_physical_id(row):
    if pd.isna(row['Physical ID']) or str(row['Physical ID']).strip() == '-':
        for col in ['Site ID - 2G', 'Site ID - 4G TDD', 'Site ID - 4G FDD', 'Site ID - 5G']:
            if pd.notna(row[col]) and str(row[col]).strip() != '-':
                return row[col]
    return row['Physical ID']

combined_df['Physical ID'] = combined_df.apply(fill_physical_id, axis=1)

print(f"Defining RCA detail from other columns...")

def classify_rca(oem, bb, bk):
    # Normalize text fields for case-insensitive matching
    bb = str(bb).lower() if pd.notnull(bb) else ""
    bk = str(bk).lower() if pd.notnull(bk) else ""

    if oem == "Ericsson":
        if "system restart alert requested restart type: power" in bb:
            return "Infra-BTS Booted due to power"
        elif "input power failure" in bb:
            return "Infra-Input Power Failure"
        elif "input voltage below configured threshold" in bb:
            return "Infra-Input Voltage Below Configured Threshold"
        elif "pmvoltage down" in bb:
            return "Infra-pmVoltage Down"
        elif any(f"voltage is {v}" in bb for v in range(36, 46)):
            return "Infra-voltage <46"
        elif "critical temperature taken out of service" in bb or "temperature high" in bb:
            return "Active(BSS)-High Temperature(Unit)"
        elif ("system restart alert requested restart type: cold" in bb or 
              "system restart alert requested restart type: data restore" in bb):
            return "Active(BSS)-Active(Cold Restart/Data Restore)"
        elif ":link failure" in bb or "link degraded" in bb:
            return "Active(BSS)-Hardware"
        else:
            return "Other/NORCA"

    elif oem == "NSN":
        if "bts booted" in bb and "due to power on" in bb:
            return "Infra-BTS Booted due to power"
        elif "unit power reset" in bb:
            return "Infra-Unit Power Reset"
        elif "low pdu input" in bb:
            return "Infra-Low PDU Input Voltage"
        elif any(f"voltage: {v}" in bb for v in range(36, 46)):
            return "Infra-voltage <46"
        elif "overheating" in bk:
            return "Active(BSS)-High Temperature(Unit)"
        elif (("bts booted" in bb and "due to software reset" in bb) or
              ("bts booted" in bb and "due to commissioning change" in bb) or
              ("bts booted" in bb and "due to recovery" in bb)):
            return "Active(BSS)-BTS Boot-Active(Sw reset/Cmm change/Recovery)"
        elif "bts booted" in bb and "due to user request" in bb:
            return "Active(BSS)-BTS Booted-Planned(User request)"
        elif "administrative state change" in bb:
            return "Active(BSS)-Active-Planned"
        else:
            return "Other/NORCA"

    else:
        return "OEM not added"
    
combined_df['RCA detail'] = combined_df.apply(
    lambda row: classify_rca(row['OEM Name'], row['Auto RCA - From Internal alarm'], row['Alarm Description']),
    axis=1
)

print(f"Defining RCA conclusive from other columns...")


def classify_rca(detail):
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

# Apply function to your DataFrame
combined_df["RCA conclusive"] = combined_df["RCA detail"].apply(classify_rca)

#==============================================MS1 Sites Dataframe=============================

df_ms1_sites = combined_df[combined_df['CBS Status (MS1 / MS2)'].str.contains('MS1|MS0|skipped', na=False, case=False)]


df_active_ms2 = combined_df[~combined_df['CBS Status (MS1 / MS2)'].str.contains('MS1', na=False, case=False)]
df_active_ms2 = df_active_ms2[
    ~df_active_ms2['Status'].str.contains('locked|theft|lock', na=False, case=False)
]
#====================================================Grouping===================================
# Total incidents by Circle
df_total_inc = combined_df.groupby('Circle').size().reset_index(name='Total Incidents')

# Total SDO incidents by Circle
df_sdo = combined_df[combined_df[mttr_col].str.strip() == "SDO"]
df_sdo_count = df_sdo.groupby('Circle')[mttr_col].count().reset_index(name="Total SDO")

# Auto RCA External filter
df_auto_rca_ext = combined_df[
    (combined_df["Auto RCA - From Ext alarm"].notna()) &
    (combined_df["Auto RCA - From Ext alarm"].str.strip() != "") &
    (combined_df["Auto RCA - From Ext alarm"].str.strip().str.lower() != "no alarm")
]
df_auto_rca_ext_count = df_auto_rca_ext.groupby('Circle').size().reset_index(name='RCA From External Alarm')

# Auto RCA External filter
df_auto_rca_int = combined_df[
    (combined_df["Auto RCA - From Internal alarm"].notna()) &
    (combined_df["Auto RCA - From Internal alarm"].str.strip() != "") &
    (combined_df["Auto RCA - From Internal alarm"].str.strip().str.lower() != "no alarm")
]
df_auto_rca_int_count = df_auto_rca_int.groupby('Circle').size().reset_index(name='RCA From Internal Alarm')

# Auto RCA External filter
df_auto_rca_txeq = combined_df[
    (combined_df["Auto RCA - From Transmission Equipment"].notna()) &
    (combined_df["Auto RCA - From Transmission Equipment"].str.strip() != "") &
    (combined_df["Auto RCA - From Transmission Equipment"].str.strip().str.lower() != "no alarm")
]
df_auto_rca_txeq_count = df_auto_rca_txeq.groupby('Circle').size().reset_index(name='RCA From Transmission Equipment')

# # Auto RCA External filter
# df_auto_rca_fiber = combined_df[
#     (combined_df["Fiber outage RCA"].notna())
#     #(combined_df["Fiber outage RCA"].str.strip() != "") &
#     #(combined_df["Fiber outage RCA"].str.strip().str.lower() != "no alarm")
# ]
# df_auto_rca_fiber_count = df_auto_rca_fiber.groupby('Circle').size().reset_index(name='RCA From Fiber')


total_filled_rca_df = combined_df[
    (combined_df["Auto RCA - From Ext alarm"].notna()) & (combined_df["Auto RCA - From Ext alarm"] != "No Alarm") |
    (combined_df["Auto RCA - From Internal alarm"].notna()) & (combined_df["Auto RCA - From Internal alarm"] != "No Alarm") |
    (combined_df["Auto RCA - From Transmission Equipment"].notna()) & (combined_df["Auto RCA - From Transmission Equipment"] != "No Alarm")
]

total_filled_rca_df_count = total_filled_rca_df.groupby('Circle').size().reset_index(name='Total Filled RCA')

tt_mapped_df = combined_df[combined_df["TT ID"].notna() & (combined_df["TT ID"].astype(str).str.strip() != "")]
#tt_mapped_df_count = tt_mapped_df.groupby('Circle').size().reset_index(name='TT Mapped')


combined_df["Duration of Cell Outage Minutes"] = pd.to_numeric(combined_df["Duration of Cell Outage Minutes"], errors='coerce')


df_less_than_15_all = combined_df[combined_df["Duration of Cell Outage Minutes"] < 15]
df_less_than_15_all_count = df_less_than_15_all.groupby('Circle').size().reset_index(name='Total Incidents Less Than 15 Minutes')



df_greater_than_15_all = combined_df[combined_df["Duration of Cell Outage Minutes"] >= 15]
df_greater_than_15_all_count = df_greater_than_15_all.groupby('Circle').size().reset_index(name='Total Incidents Greater Than equal to 15 Minutes')




df_less_than_15 = tt_mapped_df[tt_mapped_df["Duration of Cell Outage Minutes"] < 15]
df_less_than_15_count = df_less_than_15.groupby('Circle').size().reset_index(name='TT Mapping less than 15 minute')



df_greater_than_15 = tt_mapped_df[tt_mapped_df["Duration of Cell Outage Minutes"] >= 15]
df_greater_than_15_count = df_greater_than_15.groupby('Circle').size().reset_index(name='TT Mapping greater than equal to 15 minute')


df_infra = combined_df[combined_df["RCA conclusive"] == "Infra"]
df_infra_count = df_infra.groupby('Circle').size().reset_index(name='Infra')

df_active_bss = combined_df[combined_df["RCA conclusive"] == "Active(BSS)"]
df_active_bss_count = df_active_bss.groupby('Circle').size().reset_index(name='Active(BSS)')

df_no_rca = combined_df[combined_df["RCA conclusive"] == "NO RCA/Not conclusive"]
df_no_rca_count = df_no_rca.groupby('Circle').size().reset_index(name='NO RCA/Not conclusive')

df_oem_not_included = combined_df[combined_df["RCA conclusive"] == "OEM Not included yet"]
df_oem_not_included_count = df_oem_not_included.groupby('Circle').size().reset_index(name='OEM Not included yet')






formatted_date = combined_df['Date'].max().strftime('%d %b')
# Ensure the column is numeric
combined_df['Total Duration of Outage (Minutes) - All Tech'] = pd.to_numeric(
    combined_df['Total Duration of Outage (Minutes) - All Tech'], errors='coerce'
)

# Group by Circle and Physical ID, then sum the outage duration
df_all_tech_outage = combined_df.groupby(['Circle', 'Physical ID'], as_index=False)[
    'Total Duration of Outage (Minutes) - All Tech'
].sum()

# Rename column using formatted date
df_all_tech_outage.rename(
    columns={'Total Duration of Outage (Minutes) - All Tech': f'{formatted_date}'},
    inplace=True
)


# Group by Circle and Physical ID, then count the Circle
df_activity_count = combined_df.groupby(['Circle', 'Physical ID']).size().reset_index(name='Count')

# Rename column using formatted date
df_activity_count.rename(
    columns={'Count': f'{formatted_date}'},
    inplace=True
)








# Group by Circle and Physical ID, then sum the outage duration
df_all_tech_outage_circle = combined_df.groupby(['Circle'], as_index=False)[
    'Total Duration of Outage (Minutes) - All Tech'
].sum()

# Rename column using formatted date
df_all_tech_outage_circle.rename(
    columns={'Total Duration of Outage (Minutes) - All Tech': f'{formatted_date}'},
    inplace=True
)

# Calculate total sum of outage durations across all circles
total_sum = df_all_tech_outage_circle[f'{formatted_date}'].sum()

# Create a Total row as a DataFrame
total_row = pd.DataFrame({
    'Circle': ['Total'],
    f'{formatted_date}': [total_sum]
})

# Append the Total row to the DataFrame
df_all_tech_outage_circle = pd.concat([df_all_tech_outage_circle, total_row], ignore_index=True)


# Group by Circle and Physical ID, then sum the outage duration
df_all_tech_outage_circle_active = df_active_ms2.groupby(['Circle'], as_index=False)[
    'Total Duration of Outage (Minutes) - All Tech'
].sum()

# Rename column using formatted date
df_all_tech_outage_circle_active.rename(
    columns={'Total Duration of Outage (Minutes) - All Tech': f'{formatted_date}'},
    inplace=True
)

# Calculate total sum of outage durations across all circles
total_sum = df_all_tech_outage_circle_active[f'{formatted_date}'].sum()

# Create a Total row as a DataFrame
total_row = pd.DataFrame({
    'Circle': ['Total'],
    f'{formatted_date}': [total_sum]
})

# Append the Total row to the DataFrame
df_all_tech_outage_circle_active = pd.concat([df_all_tech_outage_circle_active, total_row], ignore_index=True)




#===============================================Matching TT Numbers========================================



df_matched_tt = df_ericsson_tt[
    (df_ericsson_tt['Final TT Status'] != 'Not matched') &
    (df_ericsson_tt['Outage Type'] == 'Outage/Down')
]

df_matched_tt = df_matched_tt.drop_duplicates(subset='Ericsson Incident Number', keep='first')
df_matched_tt_count = df_matched_tt.groupby('Circle').size().reset_index(name='TTs found in NGOR')




# Create Pivot Table
tt_mapped_bucket = df_matched_tt.pivot_table(
    index=['Circle'], 
    columns='Final TT Status',
    values='Ericsson Incident Number',  # Use any column that exists for counting
    aggfunc= 'count',  # Count occurrences
    fill_value=0,  # Fill missing values with 0
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

tt_mapped_bucket = df_outage_tt_count.merge(tt_mapped_bucket, on='Circle', how='left').fillna(0)

# df_matched_tt = common_outage_tt[common_outage_tt['TT Numbers'].isin(df_outage_tt['Ericsson Incident Number'])]
# df_matched_tt = df_matched_tt.drop_duplicates(subset='TT Numbers', keep='first')
# df_matched_tt_count = df_matched_tt.groupby('Circle').size().reset_index(name='TTs found in NGOR')



# Define folder paths
#======================Reading Yesterday Dashboard (Sitewise_CMO_Trend)========================================
folder_path_yesterday = r'D:\Automation\One_Outage\Output\Ngor_Dashboard'
# Get all .xlsx files in the folder
file_list_yesterday = glob.glob(os.path.join(folder_path_yesterday, '*.xlsx'))

# Filter out temporary/lock files like ~$<filename>.xlsx
valid_files = [f for f in file_list_yesterday if not os.path.basename(f).startswith('~$')]

# Sort by modified time (newest first)
valid_files = sorted(valid_files, key=os.path.getmtime, reverse=True)


# Read the first valid file
for file_path in valid_files:
    try:
        if os.path.getsize(file_path) < 1000:
            print(f"⚠️ Skipping small or empty file: {os.path.basename(file_path)}")
            continue

        print(f"📄 Reading file: {file_path}")
        df_daily_sum_yesterday = pd.read_excel(file_path, sheet_name='Sitewise_CMO_Trend', engine='openpyxl')
        print(f"✅ Successfully read: {os.path.basename(file_path)}")
        break
    except Exception as e:
        print(f"❌ Failed to read {os.path.basename(file_path)}: {e}")
        continue
else:
    raise FileNotFoundError("No valid Excel file found.")

# Drop 'Total' column if it exists
df_daily_sum_yesterday = df_daily_sum_yesterday.drop(df_daily_sum_yesterday.columns[2], axis=1)
df_daily_sum_yesterday = df_daily_sum_yesterday.drop(columns=['Total','Days ≥1000','Airtel ID'], errors='ignore')

df_daily_cmo = df_daily_sum_yesterday.merge(
    df_all_tech_outage,
    on=['Circle', 'Physical ID'],
    how='outer'  # Keeps all records from both DataFrames
).fillna(0) 

# Add Total column by summing across all date columns (ignoring first two columns: 'Circle', 'Physical ID')
df_daily_cmo['Total'] = df_daily_cmo.iloc[:, 2:].sum(axis=1)

# Get only date columns by excluding 'Circle', 'Physical ID', and 'Total'
date_columns = df_daily_cmo.columns.difference(['Circle', 'Physical ID', 'Total'])

# Count number of days with values >= 1000
df_daily_cmo['Days ≥1000'] = (df_daily_cmo[date_columns] >= 1000).sum(axis=1)


# Assuming df_daily_cmo is already defined

df_daily_cmo['Airtel ID'] = df_daily_cmo['Physical ID'].apply(
    lambda x: x[2:8] if isinstance(x, str) and len(x) == 15 else x
)


print("📊 Processing completed for Sitewise CMO Trend.")
#======================Reading Yesterday Dashboard (Circlewise_CMO_Trend)========================================

#df_daily_sum_yesterday_circle = pd.read_excel(latest_file_yesterday, sheet_name='Circlewise_CMO_Trend')

df_daily_sum_yesterday_circle = pd.read_excel(file_path, sheet_name='Circlewise_CMO_Trend', engine='openpyxl')

# Print the file name being extracted
print(f"📄 Trying to read file: {file_path}")

df_daily_cmo_circle = df_daily_sum_yesterday_circle.merge(
    df_all_tech_outage_circle,
    on=['Circle'],
    how='left'  # Keeps all records from both DataFrames
).fillna(0) 
print("📊 Processing completed for Circlewise_CMO_Trend.")


#======================Reading Yesterday Dashboard (Sitewise_Activity_Count)========================================

#df_circlewise_activity_count = pd.read_excel(latest_file_yesterday, sheet_name='Sitewise_incidents_Trend')

df_circlewise_activity_count = pd.read_excel(file_path, sheet_name='Sitewise_incidents_Trend', engine='openpyxl')
df_circlewise_activity_count = df_circlewise_activity_count.drop(df_circlewise_activity_count.columns[2], axis=1)

# Print the file name being extracted
print(f"📄 Trying to read file: {file_path}")

df_circlewise_activity_count = df_circlewise_activity_count.merge(
    df_activity_count,
    on=['Circle', 'Physical ID'],
    how='outer'  # Keeps all records from both DataFrames
).fillna(0) 
print("📊 Processing completed for Sitewise_Activity_Count.")

#======================Reading Yesterday Dashboard (Circlewise_RNA_Trend)========================================

#df_daily_rna_yesterday_circle = pd.read_excel(latest_file_yesterday, sheet_name='Circlewise_RNA_Trend')
df_daily_rna_yesterday_circle = pd.read_excel(file_path, sheet_name='Circlewise_RNA_Trend', engine='openpyxl')

# Print the file name being extracted
print(f"📄 Trying to read file: {file_path}")

df_cell_count = df_daily_rna_yesterday_circle[['Circle', 'Total Cells']]

df_all_tech_rna_circle = df_cell_count.merge(df_all_tech_outage_circle, on='Circle', how='left').fillna(0)


df_all_tech_rna_circle['RNA %'] = (
    ((df_all_tech_rna_circle['Total Cells'] * 1440) - df_all_tech_rna_circle[formatted_date])
    / (df_all_tech_rna_circle['Total Cells'] * 1440)
)*100  # Optional: Multiply by 100 to get percentage

#df_all_tech_rna_circle['RNA %'] = df_all_tech_rna_circle['RNA %'].round(2)

# Drop the formatted_date column
df_all_tech_rna_circle.drop(columns=[formatted_date,"Total Cells"], inplace=True)

# Rename 'RNA %' column to formatted_date
df_all_tech_rna_circle.rename(columns={'RNA %': formatted_date}, inplace=True)

df_rna = df_daily_rna_yesterday_circle.merge(df_all_tech_rna_circle, on='Circle', how='left').fillna(0)

 
print("📊 Processing completed for Circlewise_RNA_Trend.")

#======================Reading Yesterday Dashboard (Circlewise_RNA_Trend Excluding)========================================

#df_daily_rna_yesterday_circle = pd.read_excel(latest_file_yesterday, sheet_name='Circlewise_RNA_Trend')
df_daily_rna_yesterday_circle_ms2 = pd.read_excel(file_path, sheet_name='Circlewise_RNA_Trend_Excluding', engine='openpyxl')

# Print the file name being extracted
print(f"📄 Trying to read file: {file_path}")

df_cell_count_ms2 = df_daily_rna_yesterday_circle_ms2[['Circle', 'Total Cells']]

df_all_tech_rna_circle_ms2 = df_cell_count_ms2.merge(df_all_tech_outage_circle_active, on='Circle', how='left').fillna(0)


df_all_tech_rna_circle_ms2['RNA %'] = (
    ((df_all_tech_rna_circle_ms2['Total Cells'] * 1440) - df_all_tech_rna_circle_ms2[formatted_date])
    / (df_all_tech_rna_circle_ms2['Total Cells'] * 1440)
)*100  # Optional: Multiply by 100 to get percentage

#df_all_tech_rna_circle['RNA %'] = df_all_tech_rna_circle['RNA %'].round(2)

# Drop the formatted_date column
df_all_tech_rna_circle_ms2.drop(columns=[formatted_date,"Total Cells"], inplace=True)

# Rename 'RNA %' column to formatted_date
df_all_tech_rna_circle_ms2.rename(columns={'RNA %': formatted_date}, inplace=True)

df_rna_ms2 = df_daily_rna_yesterday_circle_ms2.merge(df_all_tech_rna_circle_ms2, on='Circle', how='left').fillna(0)

 
print("📊 Processing completed for Circlewise_RNA_Trend (excluding).")


#==========================================Read RBS_Type Excel file========================================
#Reading RBS type Excel file
rbs_type_file_path = r'D:\Automation\One_Outage\RBS_Type.xlsx'
df_rbs_type = pd.read_excel(rbs_type_file_path, engine='openpyxl')

#==========================================RNA Excluding Inactive Sites========================================



df_all_tech_rna_circle_active = df_cell_count.merge(df_all_tech_outage_circle_active, on='Circle', how='left').fillna(0)

df_all_tech_rna_circle_active['RNA %'] = (
    ((df_all_tech_rna_circle_active['Total Cells'] * 1440) - df_all_tech_rna_circle_active[formatted_date])
    / (df_all_tech_rna_circle_active['Total Cells'] * 1440)
)  # Optional: Multiply by 100 to get percentage

#df_all_tech_rna_circle['RNA %'] = df_all_tech_rna_circle['RNA %'].round(2)

# Drop the formatted_date column
df_all_tech_rna_circle_active.drop(columns=[formatted_date,"Total Cells"], inplace=True)

# Rename 'RNA %' column to formatted_date
df_all_tech_rna_circle_active.rename(columns={'RNA %': formatted_date}, inplace=True)

df_rna_active = df_daily_rna_yesterday_circle.merge(df_all_tech_rna_circle_active, on='Circle', how='left').fillna(0)

 
print("📊 Processing completed for Circlewise_RNA_Trend.")

#=========================================Dashboard DataFrame========================================
print("📊 Processing Dashboard")

# Merge SDO and Total incidents

df_dashboard = df_circles.merge(df_total_inc, on='Circle', how='left').fillna(0)
df_dashboard = df_dashboard.merge(df_sdo_count, on='Circle', how='left').fillna(0)
df_dashboard = df_dashboard.merge(df_auto_rca_ext_count, on='Circle', how='left').fillna(0)
df_dashboard = df_dashboard.merge(df_auto_rca_int_count, on='Circle', how='left').fillna(0)
df_dashboard = df_dashboard.merge(df_auto_rca_txeq_count, on='Circle', how='left').fillna(0)
#df_dashboard = df_dashboard.merge(df_auto_rca_fiber_count, on='Circle', how='left').fillna(0)
df_dashboard = df_dashboard.merge(total_filled_rca_df_count, on='Circle', how='left').fillna(0)
df_dashboard['Not Filled RCA'] = df_dashboard['Total Incidents'] - df_dashboard['Total Filled RCA']
#df_dashboard = df_dashboard.merge(tt_mapped_df_count, on='Circle', how='left').fillna(0)
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


# Calculate the sum for numeric columns
total_row = df_dashboard.select_dtypes(include="number").sum().to_frame().T

# Add a label to the first column (assuming the first column is categorical)
first_col = df_dashboard.columns[0]
total_row[first_col] = "Total"

# Ensure column order remains the same
total_row = total_row[df_dashboard.columns]

# Append the total row
df_dashboard = pd.concat([df_dashboard, total_row], ignore_index=True)

# Calculate percentage safely
df_dashboard["RCA External %"] = df_dashboard.apply(
    lambda row: f"{round((row['RCA From External Alarm'] / row['Total Incidents']) * 100, 2)}%"
    if row["Total Incidents"] and not pd.isna(row["Total Incidents"])
    else "0.0%",
    axis=1
)

df_dashboard["RCA Internal %"] = df_dashboard.apply(
    lambda row: f"{round((row['RCA From Internal Alarm'] / row['Total Incidents']) * 100, 2)}%"
    if row["Total Incidents"] and not pd.isna(row["Total Incidents"])
    else "0.0%",
    axis=1
)

df_dashboard["RCA Transmission %"] = df_dashboard.apply(
    lambda row: f"{round((row['RCA From Transmission Equipment'] / row['Total Incidents']) * 100, 2)}%"
    if row["Total Incidents"] and not pd.isna(row["Total Incidents"])
    else "0.0%",
    axis=1
)

df_dashboard["Total Filled RCA %"] = df_dashboard.apply(
    lambda row: f"{round((row['Total Filled RCA'] / row['Total Incidents']) * 100, 2)}%"
    if row["Total Incidents"] and not pd.isna(row["Total Incidents"])
    else "0.0%",
    axis=1
)


df_dashboard["TT Mapping less than 15 minute %"] = df_dashboard.apply(
    lambda row: f"{round((row['TT Mapping less than 15 minute'] / row['Total Incidents Less Than 15 Minutes']) * 100, 2)}%"
    if row["Total Incidents Less Than 15 Minutes"] and not pd.isna(row["Total Incidents Less Than 15 Minutes"])
    else "0.0%",
    axis=1
)

df_dashboard["TT Mapping greater than equal to 15 minute %"] = df_dashboard.apply(
    lambda row: f"{round((row['TT Mapping greater than equal to 15 minute'] / row['Total Incidents Greater Than equal to 15 Minutes']) * 100, 2)}%"
    if row["Total Incidents Greater Than equal to 15 Minutes"] and not pd.isna(row["Total Incidents Greater Than equal to 15 Minutes"])
    else "0.0%",
    axis=1
)

df_dashboard["TT created and Mapped %"] = df_dashboard.apply(
    lambda row: f"{round((row['TTs found in NGOR'] / row['Outage Category TTs raised']) * 100, 2)}%"
    if row["TTs found in NGOR"] and not pd.isna(row["Outage Category TTs raised"])
    else "0.0%",
    axis=1
)

df_dashboard["Infra %"] = df_dashboard.apply(
    lambda row: f"{round((row['Infra'] / row['Total Incidents']) * 100, 2)}%"
    if row["Infra"] and not pd.isna(row["Total Incidents"])
    else "0.0%",
    axis=1
)

df_dashboard["Active(BSS) %"] = df_dashboard.apply(
    lambda row: f"{round((row['Active(BSS)'] / row['Total Incidents']) * 100, 2)}%"
    if row["Active(BSS)"] and not pd.isna(row["Total Incidents"])
    else "0.0%",
    axis=1
)

df_dashboard["NO RCA/Not conclusive %"] = df_dashboard.apply(
    lambda row: f"{round((row['NO RCA/Not conclusive'] / row['Total Incidents']) * 100, 2)}%"
    if row["NO RCA/Not conclusive"] and not pd.isna(row["Total Incidents"])
    else "0.0%",
    axis=1
)

df_dashboard["OEM Not included yet %"] = df_dashboard.apply(
    lambda row: f"{round((row['OEM Not included yet'] / row['Total Incidents']) * 100, 2)}%"
    if row["OEM Not included yet"] and not pd.isna(row["Total Incidents"])
    else "0.0%",
    axis=1
)

print("📊 Processing Dashboard Completed Successfully.")


#=========================================daily CMO and Activity count trend=====================

# Drop 'Total' column if it exists
df_cmo_till_date = df_daily_cmo.drop(columns=['Total','Days ≥1000','Airtel ID'], errors='ignore')

#df_cmo_activity = df_cmo_till_date.merge(df_circlewise_activity_count, on=['Circle','Physical ID'], how='left')

df_cmo_activity = df_cmo_till_date.merge(
    df_circlewise_activity_count,
    on=['Circle', 'Physical ID'],
    how='left',
    suffixes=('_cmo', '_inc')  # Rename overlapping columns as _cmo and _act
)


# Select only columns that end with '_cmo'
cmo_cols = [col for col in df_cmo_activity.columns if col.endswith('_cmo')]

# Select the last 7 CMO columns dynamically
selected_cmo_cols = cmo_cols[-7:]

# Count how many of these columns have values greater than 1000 for each row
count_series = (df_cmo_activity[selected_cmo_cols] > 1000).sum(axis=1)

# Create 'CMO Priority' column based on condition
df_cmo_activity['CMO Priority'] = np.where(count_series > 2, count_series, 0)



# Get BR column by index
col_br = df_cmo_activity.columns[63]

# Ensure the column is numeric
df_cmo_activity[col_br] = pd.to_numeric(df_cmo_activity[col_br], errors='coerce')

# Apply multiple conditions using np.select
conditions = [
    (df_cmo_activity[col_br] > 10) & (df_cmo_activity[col_br] <= 20),
    (df_cmo_activity[col_br] > 20)
]

choices = [
    ">10 & <=20",
    ">20"
]

df_cmo_activity['Incident Priority'] = np.select(conditions, choices, default="")



# Correct incident condition for comparison
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


# Fallback-safe Merge Logic
try:
    df_cmo_activity = df_cmo_activity.merge(
        df_rbs_type[['Physical ID', 'Toco Name','RBS']],
        on='Physical ID',
        how='left'
    )
except KeyError:
    # If 'Physical ID' or same-name columns are missing, fallback to different column names
    df_cmo_activity = df_cmo_activity.merge(
        df_rbs_type[['Site_ID', 'Toco Name','RBS']],
        left_on='Physical ID',
        right_on='Site_ID',
        how='left'
    )

# Replace NaN or blank values with '#N/A' in 'Toco Name' and 'RBS' columns
df_cmo_activity[['Toco Name', 'RBS']] = df_cmo_activity[['Toco Name', 'RBS']].replace(['', None, np.nan], '#N/A')

#filter dataframe if CMO_Incident_Category is in ["Both-CMO&Incident", "High CMO Only", "High Incidents Only"]
priority_sites_cmo = df_cmo_activity[df_cmo_activity['CMO_Incident_Category'].isin(["Both-CMO&Incident", "High Incidents Only"])]
priority_sites_all = df_cmo_activity[df_cmo_activity['CMO_Incident_Category'].isin(["Both-CMO&Incident", "High Incidents Only","High CMO Only"])]


# Create Pivot Table
priority_sites_cmo_pivot = priority_sites_cmo.pivot_table(
    index='Circle', 
    columns='Incident Priority', 
    values='Physical ID',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()


# Create Pivot Table
priority_sites_all_pivot = priority_sites_cmo.pivot_table(
    index='Circle', 
    columns='RBS', 
    values='Physical ID',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()


# Create Pivot Table
priority_sites_cmo_inc = priority_sites_all.pivot_table(
    index='Circle', 
    columns='CMO_Incident_Category', 
    values='Physical ID',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()


df_cmo_inc_dashboard = priority_sites_cmo_pivot.merge(priority_sites_all_pivot, on='Circle', how='left')
df_cmo_inc_dashboard = df_cmo_inc_dashboard.merge(priority_sites_cmo_inc, on='Circle', how='left')










#==========================Project Harmony===================================
folder_path_harmony = r'D:\Automation\One_Outage\Harmony_Sites\Harmony Indus Sites & Airtel Owned _Physcial IDs.xlsx'

# Reading the "Harmony Indus Site detail" sheet
df_indus_harmony_sites = pd.read_excel(folder_path_harmony, sheet_name='Harmony Indus Site detail')
df_indus_harmony_sites_matched = df_daily_cmo[df_daily_cmo['Physical ID'].isin(df_indus_harmony_sites['Physical ID'])]
df_indus_harmony_sites_matched = df_indus_harmony_sites_matched.drop(columns=['Airtel ID'], errors='ignore')
df_indus_harmony_sites_matched = df_indus_harmony_sites_matched.merge(df_indus_harmony_sites, on='Physical ID', how='left')





# Reading the "Airtel Own_Esc Matrix" sheet
df_airtel_owned_sites = pd.read_excel(folder_path_harmony, sheet_name='Airtel Own_Esc Matrix')
df_airtel_owned_sites_matched = df_daily_cmo[df_daily_cmo['Physical ID'].isin(df_airtel_owned_sites['Physical ID'])]
df_airtel_owned_sites_matched = df_airtel_owned_sites_matched.drop(columns=['Airtel ID'], errors='ignore')
df_airtel_owned_sites_matched = df_airtel_owned_sites_matched.merge(df_airtel_owned_sites, on='Physical ID', how='left')





# ===============================USD, SDO, LDI Sites count======================

df_usd_unique_count = combined_df.groupby('Circle')['Physical ID'].nunique().reset_index(name='Unique Sites Down')


df_sdo = combined_df[combined_df[mttr_col].str.strip() == "SDO"]
df_sdo_unique_count = df_sdo.groupby('Circle')['Physical ID'].nunique().reset_index(name="Total SDI")


df_ldi = combined_df[combined_df[mttr_col].str.strip() != "SDO"]
df_ldi_unique_count = df_ldi.groupby('Circle')['Physical ID'].nunique().reset_index(name="ISD Excluding SDI")


dashboard_usd_sdo_ldi = df_circles.merge(df_usd_unique_count, on='Circle', how='left').fillna(0)
dashboard_usd_sdo_ldi = dashboard_usd_sdo_ldi.merge(df_sdo_unique_count, on='Circle', how='left').fillna(0)
dashboard_usd_sdo_ldi = dashboard_usd_sdo_ldi.merge(df_ldi_unique_count, on='Circle', how='left').fillna(0)











#==========================Project Harmony end===============================
#print(df_ericsson_tt.head())
print("📊 Writing Required Data frames.")
#========================================Printing required Dataframes==============================================
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


#yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y%m%d")
output_file = os.path.join(output_folder_ngor, f'NGOR_Dashboard_{report_date_one_outage}.xlsx')
output_file_harmony = os.path.join(output_folder_harmony, f'Harmony_Sites_{report_date_one_outage}.xlsx')
output_file_ms1_sites = os.path.join(output_folder_ms1, f'MS1_Sites_{report_date_one_outage}.xlsx')
output_file_daily_combined = os.path.join(output_folder_daily_combined, f'NGOR_Dashboard_Dump_{report_date_one_outage}.xlsx')
output_file_priority_sites = os.path.join(output_folder_priority_sites, f'Priority_Sites_{report_date_one_outage}.xlsx')



with pd.ExcelWriter(output_file_priority_sites, engine='xlsxwriter') as writer: 
    df_cmo_inc_dashboard.to_excel(writer, sheet_name="Summary", index=False) 
    priority_sites_all.to_excel(writer, sheet_name="Priority Sites Dump", index=False)
    dashboard_usd_sdo_ldi.to_excel(writer, sheet_name="USD_SDO_LDI_Sites", index=False)

print(f"\n💾 Excel file saved successfully at: {output_file_priority_sites}")

# Write to Excel
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

print(f"\n💾 Excel file saved successfully at: {output_file}")

with pd.ExcelWriter(output_file_daily_combined, engine='xlsxwriter') as writer: 
    combined_df.to_excel(writer, sheet_name="One_Outage", index=False) 
    df_ericsson_tt.to_excel(writer, sheet_name= "Ericsson_TT", index=False)       
    combined_df_Locked_LODS.to_excel(writer, sheet_name= "Locked_LODS", index=False)
    combined_df_not_in_site_db.to_excel(writer, sheet_name= "Not_in_Site_DB", index=False)

print(f"\n💾 Excel file saved successfully at: {output_file_daily_combined}")

  
with pd.ExcelWriter(output_file_harmony, engine='xlsxwriter') as writer:
    df_indus_harmony_sites_matched.to_excel(writer, sheet_name="Harmony Indus Sites", index=False)
    df_airtel_owned_sites_matched.to_excel(writer, sheet_name="Airtel Owned Sites", index=False)

print(f"\n💾 Excel file saved successfully at: {output_file_harmony}")

with pd.ExcelWriter(output_file_ms1_sites, engine='xlsxwriter') as writer:
    df_ms1_sites.to_excel(writer, sheet_name="MS1 Sites", index=False)

print(f"\n💾 Excel file saved successfully at: {output_file_ms1_sites}")

# # #==========================Sending Mails=========================

# # Common variables
# from_email = "IN_D_Engineering_NIfi_Reports@airtel.com"
# today_date = datetime.now().strftime("%d%m%Y")
# yesterday_date = (datetime.now() - timedelta(days=1)).strftime("%d%m%Y")

# # Initialize Outlook
# try:
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     namespace = outlook.GetNamespace("MAPI")
# except Exception as e:
#     print(f"❌ Error initializing Outlook: {e}")
#     sys.exit()

# # ==================== Reusable Mail Function ====================
# def send_email(recipient_file, subject, body, attachment_file):
#     try:
#         receipient_df = pd.read_excel(recipient_file)
#         to_recipients = receipient_df["to"].dropna().tolist()
#         cc_recipients = receipient_df["cc"].dropna().tolist()

#         # ========================= ⬇️ CHANGE: Store recipients before sending mail
#         to_list = ";".join(to_recipients)
#         cc_list = ";".join(cc_recipients)
#         # =========================

#         mail = outlook.CreateItem(0)
#         mail.To = to_list  # Using pre-stored recipient string
#         mail.CC = cc_list
#         mail.Subject = subject
#         mail.Body = body
#         mail.SentOnBehalfOfName = from_email

#         matched_account = None
#         for account in outlook.Session.Accounts:
#             if account.SmtpAddress.lower() == from_email.lower():
#                 matched_account = account
#                 break

#         if matched_account:
#             mail._oleobj_.Invoke(*(64209, 0, 8, 0, matched_account))
#         else:
#             print(f"⚠️ Warning: Sending account '{from_email}' not found. Using default account.")

#         if os.path.exists(attachment_file):
#             mail.Attachments.Add(attachment_file)
#         else:
#             print(f"⚠️ Attachment not found: {attachment_file}")

#         mail.Send()

#         # ========================= ⬇️ CHANGE: Printing using pre-stored recipient string
#         print(f"📧 Email sent successfully to: {to_list}")
#         # =========================

#     except Exception as e:
#         print(f"❌ Error sending email: {e}")






# # ==================== Sending Harmony Sites Mail ====================
# harmony_recipients = r"D:\Automation\One_Outage\Receipient_Harmony_Sites.xlsx"
# harmony_subject = f"Harmony Sites CMO Trend Till- {report_date_one_outage}"
# harmony_body = f"""
# Hi Team,

# Please find Harmony Sites CMO Trend from 26-05-2025 to {report_date_one_outage}.

# Regards,  
# Central Team FSO
# """

# send_email(harmony_recipients, harmony_subject, harmony_body, output_file_harmony)
# print("✅ Email process completed successfully for Harmony Sites")

# # ==================== Sending MS1 Sites Mail ====================
# ms1_recipients = r"D:\Automation\One_Outage\Receipient_MS1_Sites.xlsx"
# ms1_subject = f"MS1_Sites_CMO_Contribution- {report_date_one_outage}"
# ms1_body = f"""
# Hi Team,

# Please find Ms1 Sites CMO Contribution of: {report_date_one_outage}.

# Regards,  
# Central Team FSO
# """

# send_email(ms1_recipients, ms1_subject, ms1_body, output_file_ms1_sites)
# print("✅ Email process completed successfully for MS1 Sites")


# # ==================== Sending Priority Sites Mail ====================
# daily_cmo_recipients = r"D:\Automation\One_Outage\Receipient_Priority_Sites.xlsx"
# daily_cmo_subject = f"Daily CMO trend- {report_date_one_outage}"
# daily_cmo_body = f"""
# Hi Team,

# Please find Prioity Sites Summary of: {report_date_one_outage}.

# Regards,  
# Central Team FSO
# """

# send_email(daily_cmo_recipients, daily_cmo_subject, daily_cmo_body, output_file_priority_sites)
# print("✅ Email process completed successfully for Priority Sites")


# # Release Outlook objects
# del outlook, namespace


# # #=================================================================================================End of Code=========================================================================================================?.jcx