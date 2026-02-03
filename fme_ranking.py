from turtle import circle
import pandas as pd
import numpy as np
import glob
import os
import win32com.client
import re
import requests
from datetime import datetime, timedelta
from getpass import getpass  # Secure password input
from urllib.parse import unquote, urlparse, parse_qs
import logging
import traceback
from urllib.parse import urlparse, parse_qs, unquote
from pathlib import Path
from openpyxl import load_workbook
from win32com.client import GetActiveObject, Dispatch
from win32com.client.dynamic import Dispatch as DynamicDispatch
import sys
import win32timezone
#===================================================================Downloading WO Dump=====================================================================================
#Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")



# 1️⃣ Define Download Paths
save_paths = {
    "ClosedTT15-30": Path(r"D:\Daily_Work_Order_Dump"),
    "ClosedTT15": Path(r"D:\Daily_Work_Order_Dump"),
    "OpenTT30": Path(r"D:\Daily_Work_Order_Dump"),
    "Active_Users_Report_FME_OPS": Path(r"D:\Active_User_FME"),
}

for path in save_paths.values():
    path.mkdir(parents=True, exist_ok=True)

# 2️⃣ Connect to Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Inbox
    target_folder = inbox.Folders("Required to Work")  # Subfolder
    logging.info("✅ Connected to Outlook successfully.")
except Exception as e:
    logging.error(f"❌ Outlook connection error: {e}")
    exit()

# 3️⃣ Get Emails (Last 7 Days, Sorted by Newest)
messages = target_folder.Items
messages.Sort("[ReceivedTime]", True)
date_filter = datetime.now() - timedelta(days=7)
messages = messages.Restrict(f"[ReceivedTime] >= '{date_filter.strftime('%m/%d/%Y %H:%M %p')}'")

# 4️⃣ Subject Keywords to Match
subject_keywords = {
    "MobilityActiveClosedTT15-30DayDump V2": "ClosedTT15-30",
    "MobilityActiveClosedTT15DayDump V2": "ClosedTT15",
    "MobilityActiveOpenTT30DayDump V3": "OpenTT30",
    "Mobility_Active_Users_Report_FME_OPS": "Active_Users_Report_FME_OPS",
}

latest_emails = {}
today = datetime.today().date()

# Process Emails in a Single Loop
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
required_files = {"ClosedTT15-30", "ClosedTT15", "OpenTT30", "Active_Users_Report_FME_OPS"}
# #**5️⃣ Check if All Required Emails are Found**
# required_files = {"Mobility_Active_Users_Report_FME_OPS"}





if not required_files.issubset(set(latest_emails.keys())):
    logging.error("❌ Not all required emails were found. Exiting...")
    exit()

logging.info("✅ All required emails found. Proceeding with download...")

#**6️⃣ Ask for Credentials**
# USERNAME = input("Enter your username: ").strip()
# PASSWORD = getpass("Enter your password: ").strip()

credentials = pd.read_excel(r"D:\Automation\FME_Ranking\credentials.xlsx")
USERNAME = credentials.iloc[0, 0]  # First row, first column (username)
PASSWORD = credentials.iloc[0, 1]  # First row, second column (password)
print("Credentials loaded successfully.")

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
                    file_path = save_paths[filename_part] / file_name

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

#======================================================================Processing Active FME Data===========================================================================================


circles = [
    "AP","BR","CN","DL","GJ","HP","HR","JK","KK","KL","KO",
    "MH","MP","MU","NESA","OR","PB","RJ","TN","UE","UW","WB","Total"
]

df_circles = pd.DataFrame(circles, columns=['circle'])


# Folder path
folder_path = r"D:\Active_User_FME"

today_date = datetime.now().strftime('%Y%m%d')

# Search for today's file
file_pattern = os.path.join(folder_path, f"*{today_date}*.csv")
files = glob.glob(file_pattern)
#columns_to_load = ["circle", "name","olm_id", "msisdn", "manager", "manager_msisdn","site"]
if files:
    latest_file = files[0]  # Assuming there's only one file for today
    df_fme = pd.read_csv(latest_file)
    print(f"Loaded file: {latest_file}")
    # Remove rows where circle == "DYMMYTNG"
    df_fme = df_fme[df_fme['circle'] != "DYMMYTNG"]
    df_fme = df_fme[~df_fme['olm_id'].isin(["A1KLLK3D", "B0313581", "B0201025"])]
    df_fme = df_fme[~((df_fme['olm_id'] == "A1V3UAL0") & (df_fme['circle'] == "HP"))]
    # Replace 'AS' and 'NE' with 'NESA' in 'circle' column
    df_fme['circle'] = df_fme['circle'].replace({'AS': 'NESA', 'NE': 'NESA'})
    # Remove rows where site is blank
    df_fme = df_fme[df_fme['site'].notna()]        
    df_fme['unique_id'] = df_fme['circle'] + "_" + df_fme['site']
    df_fme['FME Status'] = "Active"
    
    
    #print(df_fme.head())  # Display first few rows
else:
    print("No file found for today.")

df_fme_grouped = (
                df_fme.groupby(["circle", "name", "olm_id", "msisdn", "manager_name", "manager_msisdn","FME Status"])
                .agg(Total_Sites=('unique_id', 'nunique'))  # Count distinct 'unique_id'
                .reset_index()
            )
#==============================================DF_FME====================================



# #==========================================================================Processing WO Data====================================================================================================

today_date = datetime.now().strftime('%Y%m%d')
# Specify the folder path, using either double backslashes or a raw string
folder_path_wo = "D:\\Daily_Work_Order_Dump"
# Get all .csv files in the folder
file_paths = glob.glob(os.path.join(folder_path_wo, f"*{today_date}*.csv"))

#print(df_fme_data.info())
# List to hold data from each file
all_data = []

# Read each file and append the data to the list

# Read each file and append the data to the list
for file_path in file_paths:  
        try:
            data = pd.read_csv(file_path, low_memory=False, on_bad_lines = 'skip')
            # Get the file name from the file path
            file_name = os.path.basename(file_path)
            # Add a new column 'Source_File' with the file name
            data['Source_File'] = file_name
            # Append the data to the list
            all_data.append(data)
            print(f"Read successfully: {file_path} with shape {data.shape}")
        except Exception as e:
            print(f"Error reading {file_path}: {e}")

# Combine all dataframes into one (optional)
combined_data_all = pd.concat(all_data, ignore_index=True)
combined_data_all['WO Date/Time'] = pd.to_datetime(combined_data_all['WO Date/Time'], errors='coerce')
combined_data_all['WO Date'] = pd.to_datetime(combined_data_all['WO Date/Time']).dt.date
combined_data_all['WO Date'] = pd.to_datetime(combined_data_all['WO Date'], errors='coerce')
combined_data_all['In Progress on site Date/Time'] = pd.to_datetime(combined_data_all['In Progress on site Date/Time'], errors='coerce')
combined_data_all['WO Completion Date/time'] = pd.to_datetime(combined_data_all['WO Completion Date/time'], errors='coerce')
combined_data_all['RAN PM form submitted date / time'] = pd.to_datetime(combined_data_all['RAN PM form submitted date / time'], errors='coerce')
combined_data_all['In Progress on site Date'] = pd.to_datetime(combined_data_all['In Progress on site Date/Time']).dt.date
combined_data_all['In Progress on site Date'] = pd.to_datetime(combined_data_all['In Progress on site Date'], errors='coerce')
combined_data_all['In Progress On Site Time'] = combined_data_all['In Progress on site Date/Time'].dt.strftime('%H:%M')

combined_data_all = combined_data_all.rename(columns={"WO Status": "WO_Status", "Circle" : "circle", "WO Assignee Mobile No": "msisdn"})

# Convert msisdn in main DF
combined_data_all['msisdn'] = pd.to_numeric(combined_data_all['msisdn'], errors='coerce').astype('Int64')



# Read exclude list from Excel
exclude_df = pd.read_excel(r"D:\Automation\FME_Ranking\exclusion.xlsx")

# Assuming the column name is 'msisdn'
exclude = exclude_df['msisdn'].astype('Int64').tolist()
# Filter
combined_data_all = combined_data_all[~combined_data_all['msisdn'].isin(exclude)]



combined_data_all = combined_data_all[combined_data_all['circle'] != 'DYMMYTNG' ]
combined_data_all['circle'] = combined_data_all['circle'].replace({'AS': 'NESA', 'NE': 'NESA','North East': 'NESA'})


combined_data_all = combined_data_all[combined_data_all['Airtel Site ID'].str.strip() != ""]
combined_data_all = combined_data_all.dropna(subset=['Airtel Site ID'])
combined_data_all['Airtel Site ID'] = combined_data_all['Airtel Site ID'].astype(str).str.upper()
combined_data_all['Unique_Site_ID'] = combined_data_all['circle'] + "_" + combined_data_all['Airtel Site ID']

combined_data_all["trip_done_by"] = np.where(
    combined_data_all["trip_order_id(2w/4w)"].str.contains("M_TRIP", na=False),
    "4 Wheeler",
    np.where(
        combined_data_all["trip_order_id(2w/4w)"].str.contains("2W_TRIP", na=False),
        "2 Wheeler",
        "No Vehicle Used"   # fallback value
    )
)

# Function to classify alarm types based on Additional Info
def work_order_status(Source_File):
    if isinstance(Source_File, str):
        if 'Closed' in Source_File:
            return 'Closed'
        elif 'Open' in Source_File:
            return 'Open'
    return np.nan
combined_data_all['Open/Closed'] = combined_data_all['Source_File'].apply(work_order_status)
# Sort the DataFrame
combined_data_all = combined_data_all.sort_values(by=["WO Number","Open/Closed"])
combined_data_all = combined_data_all.drop_duplicates(subset="WO Number", keep="first")


# Get today's date with '00:00:00' time component
today_date_only = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
# # Filter out rows where 'In Progress on site Date' is today's date
combined_data_all = combined_data_all[combined_data_all['In Progress on site Date'] != today_date_only]

combined_data_all['Completed Within SLA'] = combined_data_all['Completed Within SLA'].mask(
    combined_data_all['Completed Within SLA'].isna() | (combined_data_all['Completed Within SLA'].astype(str).str.strip() == ''),
    combined_data_all['Within SLA']
)


# Define conditions
conditions = [
    (combined_data_all['Travel Time'].isna()) | (combined_data_all['Travel Time'] == 0),
    (combined_data_all['Travel Time'] > 0) & (combined_data_all['Travel Time'] < 30),
    (combined_data_all['Travel Time'] >= 30) & (combined_data_all['Travel Time'] <= 180),
    (combined_data_all['Travel Time'] > 180)
]

# Define corresponding categories
choices = ['0 MIN', '>0 & <30 Min', '30 min-3 hrs', '>3 Hours']

# Apply conditions to create new column
combined_data_all['Travel Time Category'] = np.select(conditions, choices, default='Unknown')

#==============================================================Calculating Activity Time======================================

# Calculate Activity Time only for WO Types: CORRECTIVE, OPTIMIZATION, PLANNED
mask = combined_data_all['WO Type'].isin(['CORRECTIVE', 'OPTIMIZATION', 'PLANNED'])

combined_data_all.loc[mask, 'Activity Time'] = (
    combined_data_all['WO Completion Date/time'] - combined_data_all['In Progress on site Date/Time']
)

# Calculate duration in minutes and round
combined_data_all.loc[mask, 'Activity Time (Minutes)'] = (
    combined_data_all.loc[mask, 'Activity Time'].dt.total_seconds() / 60
).round(2)


# Calculate Activity Time only for WO Type "CORRECTIVE"
combined_data_all.loc[combined_data_all['WO Type'] == 'PREVENTIVE', 'Activity Time'] = \
    combined_data_all['RAN PM form submitted date / time'] - combined_data_all['In Progress on site Date/Time']

# Calculate duration in minutes and round, again only for "CORRECTIVE"
combined_data_all.loc[combined_data_all['WO Type'] == 'PREVENTIVE', 'Activity Time (Minutes)'] = \
    combined_data_all['Activity Time'].dt.total_seconds() / 60

combined_data_all['Activity Time (Minutes)'] = combined_data_all['Activity Time (Minutes)'].round(2)
#combined_data.drop(columns=['Activity Time'], inplace=True)
combined_data_all['Activity Time'] = combined_data_all['Activity Time'].fillna(0)
combined_data_all['Activity Time (Minutes)'] = combined_data_all['Activity Time (Minutes)'].fillna(0)

# Define conditions
conditions = [
    (combined_data_all['Activity Time (Minutes)'].isna()) | (combined_data_all['Activity Time (Minutes)'] == 0),
    (combined_data_all['Activity Time (Minutes)'] > 0) & (combined_data_all['Activity Time (Minutes)'] < 30),
    (combined_data_all['Activity Time (Minutes)'] >= 30) & (combined_data_all['Activity Time (Minutes)'] <= 180),
    (combined_data_all['Activity Time (Minutes)'] > 180) & (combined_data_all['Activity Time (Minutes)'] <= 300),
    (combined_data_all['Activity Time (Minutes)'] > 300)
]

# Define corresponding categories
choices = ['0 MIN', '>0 & <30 Min', '30 min-3 hrs', '3-5 hr','>5 Hr']

# Apply conditions to create new column
combined_data_all['Activity Time Category'] = np.select(conditions, choices, default='Unknown')
combined_data_all['Category'] = combined_data_all['Description'].str.split(':').str[0]

#=================================Calculating QIA, Outage and Non QIA=============================

# Define conditions
conditions = [
    (combined_data_all['WO Type'].eq('CORRECTIVE')) & (combined_data_all['Description'].str.contains('QIA', case=False, na=False)),
    (combined_data_all['WO Type'].eq('CORRECTIVE')) & (combined_data_all['Description'].str.contains('Outage|Down', case=False, na=False)),
    (combined_data_all['WO Type'].eq('CORRECTIVE'))   # if corrective but doesn’t match QIA/Outage
]

# Define choices
choices = ['QIA', 'Outage', 'Non QIA']

# Create new column
combined_data_all['WO_Category'] = np.select(conditions, choices, default=combined_data_all['WO Type'])
df_outage_dump = combined_data_all[combined_data_all['WO_Category'] == 'Outage']

#================================================SVD Work Orders=================================

df_svd = combined_data_all[combined_data_all['Site Visit Status'] == 'SVD']
df_corrective_wo = combined_data_all[combined_data_all['WO Type'] == 'CORRECTIVE']

#=======================================Sitewise WO generation Type Corrective====================
df_repeat_wo_corrective = (
    df_corrective_wo
    .groupby(['Unique_Site_ID'])['WO Number']
    .count()
    .reset_index(name="wo_generated")
)

df_repeat_wo_corrective = df_repeat_wo_corrective[df_repeat_wo_corrective['wo_generated'] > 10]

# Create Pivot Table
df_wo_type_corrective = df_corrective_wo.pivot_table(
    index=['Unique_Site_ID'], 
    columns='WO_Category', 
    values='WO Number',  
    aggfunc='count',
    fill_value=np.nan
).reset_index()


df_repeat_wo_corrective = df_repeat_wo_corrective.merge(df_wo_type_corrective, on='Unique_Site_ID', how='left')





#=====================================Repeat WO Trend=============================================

df_repeat_wo = (
    df_corrective_wo
    .groupby(['circle', 'WO Date', 'Unique_Site_ID', 'WO_Category'])['WO Number']
    .count()
    .reset_index(name="wo_generated")
)

df_repeat_wo = df_repeat_wo[df_repeat_wo['wo_generated'] > 10]

df_qia_repeat = df_repeat_wo[df_repeat_wo['WO_Category'] == 'QIA']
df_outage_repeat = df_repeat_wo[df_repeat_wo['WO_Category'] == 'Outage']
df_non_qia_repeat = df_repeat_wo[df_repeat_wo['WO_Category'] == 'Non QIA']
#===========================================Repeat WO Daily Summary===============================

# Create Pivot Table
df_qia_repeat_daywise = df_qia_repeat.pivot_table(
    index=['circle'], 
    columns='WO Date', 
    values='Unique_Site_ID',  
    aggfunc=lambda x: x.nunique(),
    fill_value=np.nan,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()


# Format Date Columns Properly
df_qia_repeat_daywise.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in df_qia_repeat_daywise.columns
]


df_qia_repeat_daywise = df_circles.merge(df_qia_repeat_daywise, on='circle', how='left')


# Create Pivot Table
df_outage_repeat_daywise = df_outage_repeat.pivot_table(
    index=['circle'], 
    columns='WO Date', 
    values='Unique_Site_ID',  
    aggfunc=lambda x: x.nunique(),
    fill_value=np.nan,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()


# Format Date Columns Properly
df_outage_repeat_daywise.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in df_outage_repeat_daywise.columns
]

df_outage_repeat_daywise = df_circles.merge(df_outage_repeat_daywise, on='circle', how='left')

# Create Pivot Table
df_non_qia_repeat_daywise = df_non_qia_repeat.pivot_table(
    index=['circle'], 
    columns='WO Date', 
    values='Unique_Site_ID',  
    aggfunc=lambda x: x.nunique(),
    fill_value=np.nan,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()


# Format Date Columns Properly
df_non_qia_repeat_daywise.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in df_non_qia_repeat_daywise.columns
]

df_non_qia_repeat_daywise = df_circles.merge(df_non_qia_repeat_daywise, on='circle', how='left')

#====================================Transmission Work Order Trend=================================

# Filter rows where 'Description' contains 'txn' (case-insensitive)
df_txn_wo = combined_data_all[
    combined_data_all['Description'].str.contains('txn', case=False, na=False) &
    ~combined_data_all['Description'].str.contains('outage|down', case=False, na=False)
]


# Calculate cutoff date (last 15 days)
cutoff_date = pd.Timestamp.today() - pd.Timedelta(days=16)

# Filter df_txn_wo for only the last 15 days
df_txn_wo_last15 = df_txn_wo[df_txn_wo['WO Date'] >= cutoff_date].copy()

#=====================================working on Transmission Work Order Trend=================================================

# Create Pivot Table
df_txn_wo_summary = df_txn_wo_last15.pivot_table(
    index=['circle'], 
    columns='Open/Closed', 
    values='WO Number',  
    aggfunc='count',
    fill_value=0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

df_txn_wo_summary = df_txn_wo_summary.rename(
    columns={
        "Open": "Pending",
        "Closed": "Total Closure",
        "Total": "Total Creation"
    }
)


# Reorder columns (keeping circle + Total row intact)
df_txn_wo_summary = df_txn_wo_summary[["circle", "Total Creation", "Total Closure", "Pending"]]


# Create Pivot Table
df_txn_wo_summary_daywise = df_txn_wo_last15.pivot_table(
    index=['circle'], 
    columns='WO Date', 
    values='WO Number',  
    aggfunc='count',
    fill_value=0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()


# Format Date Columns Properly
df_txn_wo_summary_daywise.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in df_txn_wo_summary_daywise.columns
]

df_txn_wo_category = df_txn_wo_last15.pivot_table(
    index=['Category'], 
    columns='WO Date', 
    values='WO Number',  
    aggfunc='count',
    fill_value=0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

# Format Date Columns Properly
df_txn_wo_category.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in df_txn_wo_category.columns
]



#==============================Sitewise SVD Count and WO Category=========================================================
# Create Pivot Table
sitewise_svd_count = df_svd.pivot_table(
    index=['circle','Unique_Site_ID'], 
    columns='In Progress on site Date', 
    values='WO Number',  
    aggfunc='count',
    fill_value=np.nan,
).reset_index()

# Format Date Columns Properly
sitewise_svd_count.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in sitewise_svd_count.columns
]

sitewise_svd_count['Total Site Visit MTD'] = sitewise_svd_count.iloc[:, 2:].count(axis=1)


# Assuming 'Total Site Visit MTD' column corresponds to AG2 in Excel
conditions = [
    sitewise_svd_count['Total Site Visit MTD'] == 1,
    sitewise_svd_count['Total Site Visit MTD'] == 2,
    sitewise_svd_count['Total Site Visit MTD'] == 3,
    (sitewise_svd_count['Total Site Visit MTD'] < 8) & (sitewise_svd_count['Total Site Visit MTD'] > 3),
    sitewise_svd_count['Total Site Visit MTD'] > 7
]

choices = [
    '1 Visit',
    '2 Visits',
    '3 Visits',
    '4-7 Visits',
    '>7 Visits'
]

# Apply the conditions
sitewise_svd_count['SVD MTD Category'] = np.select(conditions, choices, default='No Visit')



sitewise_svd_cat_count = df_svd.pivot_table(
    index=['circle','Unique_Site_ID'], 
    columns='WO Type', 
    values='WO Number',  
    aggfunc='count',
    fill_value=np.nan,
).reset_index()


sitewise_svd_count = sitewise_svd_count.merge(sitewise_svd_cat_count, on=['circle','Unique_Site_ID'], how='left')





# Create Pivot Table
circlewise_svd_cat_count = sitewise_svd_count.pivot_table(
    index=['circle'], 
    columns='SVD MTD Category', 
    values='Unique_Site_ID',  
    aggfunc='count',
    fill_value=0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()





#==============================Sitewise SVD Count and WO Category=========================================================

df_svd_non_cat = df_svd[~df_svd['Description'].str.contains("Spare handling\(CAT\)", na=False)]
# Create Pivot Table
sitewise_svd_count_non_cat = df_svd_non_cat.pivot_table(
    index=['circle','Unique_Site_ID'], 
    columns='In Progress on site Date', 
    values='WO Number',  
    aggfunc='count',
    fill_value=np.nan,
).reset_index()

# Format Date Columns Properly
sitewise_svd_count_non_cat.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in sitewise_svd_count_non_cat.columns
]

sitewise_svd_count_non_cat['Total Site Visit MTD'] = sitewise_svd_count_non_cat.iloc[:, 2:].count(axis=1)


# Assuming 'Total Site Visit MTD' column corresponds to AG2 in Excel
conditions = [
    sitewise_svd_count_non_cat['Total Site Visit MTD'] == 1,
    sitewise_svd_count_non_cat['Total Site Visit MTD'] == 2,
    sitewise_svd_count_non_cat['Total Site Visit MTD'] == 3,
    (sitewise_svd_count_non_cat['Total Site Visit MTD'] < 8) & (sitewise_svd_count_non_cat['Total Site Visit MTD'] > 3),
    sitewise_svd_count_non_cat['Total Site Visit MTD'] > 7
]

choices = [
    '1 Visit',
    '2 Visits',
    '3 Visits',
    '4-7 Visits',
    '>7 Visits'
]

# Apply the conditions
sitewise_svd_count_non_cat['SVD MTD Category'] = np.select(conditions, choices, default='No Visit')


sitewise_svd_non_cat_wo_type = df_svd_non_cat.pivot_table(
    index=['circle','Unique_Site_ID'], 
    columns='WO Type', 
    values='WO Number',  
    aggfunc='count',
    fill_value=np.nan,
).reset_index()


sitewise_svd_non_cat_count = sitewise_svd_count_non_cat.merge(sitewise_svd_non_cat_wo_type, on=['circle','Unique_Site_ID'], how='left')






sitewise_svd_non_cat_count = sitewise_svd_non_cat_count.merge(
    df_fme[['unique_id', 'name', 'msisdn', 'manager_name', 'manager_msisdn']],
    left_on='Unique_Site_ID',
    right_on='unique_id',
    how='left'  # keep all rows from sitewise_svd_non_cat_count
)

# Drop duplicate join key from df_fme if not needed
sitewise_svd_non_cat_count.drop(columns=['unique_id'], inplace=True)












# Create Pivot Table
circlewise_svd_cat_count_non_cat = sitewise_svd_non_cat_count.pivot_table(
    index=['circle'], 
    columns='SVD MTD Category', 
    values='Unique_Site_ID',  
    aggfunc='count',
    fill_value=0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

#========Creating Unique User Database (Active FME + FME Found in WO Data Except CEM in Active FME)=============================
df_users_wo = df_svd[["circle", "WO Assignee Name", "msisdn"]].drop_duplicates()
df_users_wo = df_users_wo.rename(columns={"WO Assignee Name": "name"})
df_users_wo = df_users_wo[df_users_wo['name'] != "Ram Kushal"]
df_users_wo = df_users_wo[~df_users_wo['msisdn'].isin(df_fme['msisdn']) & 
                          ~df_users_wo['msisdn'].isin(df_fme['manager_msisdn'])]
df_users_wo['FME Status'] = "WO_Dump"

df_fmepluswo = pd.concat([df_fme_grouped, df_users_wo], ignore_index=True)  

df_fmepluswo.update(df_fmepluswo[['circle', 'name', 'olm_id', 'manager_name', 'FME Status']].fillna("#N/A"))
df_fmepluswo['Total_Sites'] = df_fmepluswo['Total_Sites'].fillna(0)
df_fmepluswo['manager_msisdn'] = df_fmepluswo['manager_msisdn'].fillna(999999999)

#=======================================FME wise Productivity Module=============================

# Group by 'msisdn' and count unique 'WO Number' occurrences
df_svd_count = df_svd.groupby('msisdn')['WO Number'].count().reset_index(name="MTD WO")
df_svd_count = df_svd.groupby('msisdn')['WO Number'].count().reset_index(name="MTD WO")

df_working_day_count = df_svd.groupby('msisdn')['In Progress on site Date'].nunique().reset_index(name="Total Working Days")
df_working_day_count['Zero WO Days'] = df_svd['In Progress on site Date'].nunique() - df_working_day_count['Total Working Days'] 
latest_date = df_svd['In Progress on site Date'].max()
latest5daysdump = df_svd[df_svd['In Progress on site Date'] >= (latest_date - pd.Timedelta(days=4))]


# Create Pivot Table
daily_wo_count = df_svd.pivot_table(
    index='msisdn', 
    columns='In Progress on site Date', 
    values='WO Number',  
    aggfunc='count',
    fill_value=0,
).reset_index()

# Format Date Columns Properly
daily_wo_count.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in daily_wo_count.columns
]



# Create Pivot Table
df_trip_done_by = df_svd.pivot_table(
    index='circle', 
    columns='trip_done_by', 
    values='WO Number',  
    aggfunc='count',
    fill_value=0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()









# Create Pivot Table
daily_wo_count_unique_sites = df_svd.pivot_table(
    index='msisdn', 
    columns='In Progress on site Date', 
    values='Unique_Site_ID',  
    aggfunc=lambda x: x.nunique(),  # Aggregate function (count occurrences)
    fill_value=0,
).reset_index()

# Format Date Columns Properly
daily_wo_count_unique_sites.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in daily_wo_count_unique_sites.columns
]




df_unique_site_visit = df_fmepluswo.merge(df_working_day_count, on='msisdn', how='left').fillna(0)
df_unique_site_visit = df_unique_site_visit.merge(daily_wo_count_unique_sites, on='msisdn', how='left').fillna(0)
# Convert 'msisdn' to int64




#daily_wo_count['msisdn'] = daily_wo_count['msisdn'].astype('Int64') 

# # Convert msisdn safely
# daily_wo_count['msisdn'] = pd.to_numeric(daily_wo_count['msisdn'], errors='coerce')
# daily_wo_count['msisdn'] = daily_wo_count['msisdn'].astype('Int64')



last_col = df_unique_site_visit.columns[-1]

df_unique_site_visit["SVD Category"] = df_unique_site_visit[last_col].apply(
    lambda x: "0 SVD" if x == 0 else
              "1 SVD" if x == 1 else
              "2-3 SVD" if x <= 3 else
              "4 SVD" if x == 4 else
              ">=5 SVD"
)


# Step 1: Identify all date columns (e.g., '11 Jul', '01 Jun', etc.)
date_columns = [col for col in df_unique_site_visit.columns if re.match(r'\d{2} \w{3}', str(col))]

# Step 2: Calculate the sum of all date columns row-wise
df_unique_site_visit['Total SVD Count'] = df_unique_site_visit[date_columns].sum(axis=1)

# Step 3: Compute Avg SVD safely (avoid division by zero)
df_unique_site_visit['Avg SVD'] = np.where(
    df_unique_site_visit['Total Working Days'] == 0,
    0,
    df_unique_site_visit['Total SVD Count'] / df_unique_site_visit['Total Working Days']
)

# Optional: Round to 2 decimal places
df_unique_site_visit['Avg SVD'] = df_unique_site_visit['Avg SVD'].round(2)


def classify_avg_svd(avg_svd):
    if avg_svd < 1:
        return "0-1 SVD"
    elif avg_svd < 2:
        return "1-2 SVD"
    elif avg_svd < 3:
        return "2-3 SVD"
    elif avg_svd < 4:
        return "3-4 SVD"
    elif avg_svd < 5:
        return "4-5 SVD"
    elif avg_svd < 6:
        return "5-6 SVD"
    elif avg_svd < 7:
        return "6-7 SVD"
    else:
        return "7+ SVD"

# Apply to your DataFrame (e.g., df_unique_site_visit)
df_unique_site_visit['Average Category'] = df_unique_site_visit['Avg SVD'].apply(classify_avg_svd)




fme_svd_bucket = df_unique_site_visit.pivot_table(
    index='circle', 
    columns='SVD Category', 
    values='olm_id',  
    aggfunc='count',
    fill_value=0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

fme_unique_site_prod = df_unique_site_visit.pivot_table(
    index='circle', 
    columns='Average Category', 
    values='olm_id',  
    aggfunc='count',
    fill_value=0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

 

#===============================================Merging on All FMES ================================
# Merge with df_fme_prod
df_fme_prod = df_fmepluswo.merge(df_svd_count, on='msisdn', how='left').fillna(0)
df_fme_prod = df_fme_prod.merge(df_working_day_count, on='msisdn', how='left').fillna(0)

# Avoid division by zero
df_fme_prod['Avg WO/Day'] = df_fme_prod['MTD WO'] / df_fme_prod['Total Working Days']
df_fme_prod['Avg WO/Day'] = df_fme_prod['Avg WO/Day'].replace([float('inf'), -float('inf')], 0).fillna(0).round(1)

# Check WO existence in the last 5 days
df_fme_prod['FME with No WO in last 5 days(Yes/No)'] = df_fme_prod['msisdn'].apply(
    lambda x: 'Yes' if x not in latest5daysdump['msisdn'].values else 'No'
)

# Final merge with daily_wo_count
df_fme_prod = df_fme_prod.merge(daily_wo_count, on='msisdn', how='left').fillna(0)



#================================================Circle wise creation===============================
df_corrective = df_svd[df_svd["WO Type"] == "CORRECTIVE"]
circle_site_count = df_fme.groupby('circle')['unique_id'].count().reset_index(name="Number of Sites")
wo_created = df_corrective.groupby('circle')['WO Number'].count().reset_index(name="Total WO Created")
# Create Pivot Table
daily_wo_count_corrective = df_corrective.pivot_table(
    index='circle', 
    columns='In Progress on site Date', 
    values='WO Number',  
    aggfunc='count',
    fill_value=0,
).reset_index()
# Format Date Columns Properly
daily_wo_count_corrective.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in daily_wo_count_corrective.columns
]


#====================================merging with Circle wise creation==============================

circle_site_count = circle_site_count.merge(wo_created, on='circle', how='left').fillna(0)
circle_site_count["Number of WO creation/Site"] = (
    circle_site_count["Total WO Created"] / circle_site_count["Number of Sites"].replace(0, np.nan)
).round(1)

circle_site_count["Daily Average creation"] = (
    circle_site_count["Total WO Created"] / df_corrective['In Progress on site Date'].nunique()
).round(1)

circle_site_count = circle_site_count.merge(daily_wo_count_corrective, on='circle', how='left').fillna(0)



# Calculate the sum for numeric columns
total_row = circle_site_count.select_dtypes(include="number").sum().to_frame().T

# Add a label to the first column (assuming the first column is categorical)
first_col = circle_site_count.columns[0]
total_row[first_col] = "Total"

# Ensure column order remains the same
total_row = total_row[circle_site_count.columns]

# Append the total row
circle_site_count = pd.concat([circle_site_count, total_row], ignore_index=True)


#============================================11 AM Dashboard==========================================
svd_matched_fme = df_svd[df_svd['msisdn'].isin(df_fmepluswo['msisdn'])]

grth11AM = svd_matched_fme.groupby(['circle', 'msisdn', 'In Progress on site Date'])['In Progress On Site Time'].min().reset_index()
# Filter for 'In Progress On Site Time' greater than 11:00 AM
grth11AM = grth11AM[grth11AM['In Progress On Site Time'] > '11:00']

# Count unique dates where work started after 11 AM
grth11AM_count = grth11AM.groupby('msisdn')['In Progress on site Date'].nunique().reset_index(name="MTD Greater than 11 AM count")


# Create Pivot Table
daily_11_AM_FME = df_svd.pivot_table(
    index='msisdn', 
    columns='In Progress on site Date', 
    values='In Progress On Site Time',  
    aggfunc='min',
    fill_value= "NOSVD"
).reset_index()
# Format Date Columns Properly
daily_11_AM_FME.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in daily_11_AM_FME.columns
]
#====================================merging with Circle wise creation (df_fme_11am_backup)==============================
# # Merge with df_fme_prod
df_fme_11am_backup = df_fmepluswo.merge(df_working_day_count, on='msisdn', how='left').fillna(0)
df_fme_11am_backup = df_fme_11am_backup.merge(grth11AM_count, on='msisdn', how='left').fillna(0)
df_fme_11am_backup = df_fme_11am_backup.merge(df_svd_count, on='msisdn', how='left').fillna(0)

# Define conditions based on 'Total Working Days' column
conditions = [
    df_fme_11am_backup['Total Working Days'] > 25,
    df_fme_11am_backup['Total Working Days'] < 10,
    df_fme_11am_backup['Total Working Days'].between(10, 19),
    df_fme_11am_backup['Total Working Days'].between(20, 25)
]

# Define corresponding outputs
choices = [
    ">25 working days",
    "<10 working days",
    ">=10 & <20 Working days",
    ">=20 & <=25 working days"
]

# Apply conditions
df_fme_11am_backup['Working Days Category'] = np.select(conditions, choices, default="Other")



# Avoid division by zero
df_fme_11am_backup['Avg WO/Day'] = df_fme_11am_backup['MTD WO'] / df_fme_11am_backup['Total Working Days']
df_fme_11am_backup['Avg WO/Day'] = df_fme_11am_backup['Avg WO/Day'].replace([float('inf'), -float('inf')], 0).fillna(0).round(1)



# Define conditions based on 'Avg WO/Day' column
conditions = [
    df_fme_11am_backup['Avg WO/Day'] < 1,
    df_fme_11am_backup['Avg WO/Day'] <= 2,
    df_fme_11am_backup['Avg WO/Day'] <= 4,
    df_fme_11am_backup['Avg WO/Day'] > 4
]

# Define corresponding output labels
choices = [
    "<=1 productivity",
    ">1 & <=2 productivity",
    ">2 & <=4 productivity",
    ">4 productivity"
]

# Apply the conditions
df_fme_11am_backup['Productivity Category'] = np.select(conditions, choices, default="Other")

df_fme_11am_backup.rename(columns={'Zero WO Days': 'NOSVD Days'}, inplace=True)
df_fme_11am_backup = df_fme_11am_backup.merge(daily_11_AM_FME, on='msisdn', how='left').fillna("NOSVD")

#===================================11AM Summary Productivity======================================

# Create Pivot Table
working_day_summary = df_fme_11am_backup.pivot_table(
    index='circle', 
    columns='Working Days Category', 
    values='name',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

working_day_summary.rename(columns={'Total': 'Total Active FME'}, inplace=True)


# Create Pivot Table
productivity_summary = df_fme_11am_backup.pivot_table(
    index='circle', 
    columns='Productivity Category', 
    values='name',  
    aggfunc='count',
    fill_value= 0,
    margins=True,  # Add totals for both rows
    margins_name='Total'  # Name for the row and column totals
).reset_index()

productivity_summary.rename(columns={'Total': 'Productivity Total'}, inplace=True)

working_day_summary = working_day_summary.merge(productivity_summary, on='circle', how='left')

# Define the desired column order
column_order = [
    'circle', 
    '<10 working days', 
    '>=10 & <20 Working days', 
    '>=20 & <=25 working days', 
    '>25 working days', 
    'Total Active FME',  
    '<=1 productivity', 
    '>1 & <=2 productivity', 
    '>2 & <=4 productivity', 
    '>4 productivity', 
    'Productivity Total'
]

# Reorder the DataFrame
working_day_summary = working_day_summary.reindex(columns=column_order)


#==============================================Circle Productivity======================
latest1daysdump = df_svd[df_svd['In Progress on site Date'] > (latest_date - pd.Timedelta(days=1))]
latest1daysdump_fme = latest1daysdump[latest1daysdump['msisdn'].isin(df_fmepluswo['msisdn'])]
latest1daysdump_cem = latest1daysdump[latest1daysdump['msisdn'].isin(df_fme['manager_msisdn'])]



df_svd_count_circle_fme = latest1daysdump_fme.groupby('circle')['msisdn'].nunique().reset_index(name="Handeled FME")
total_fme_count = df_fmepluswo.groupby('circle')['msisdn'].count().reset_index(name="Total FME")
circle_fme_count = total_fme_count.merge(df_svd_count_circle_fme, on='circle', how='left')


circle_fme_count["Not Handled FME"] = circle_fme_count["Total FME"] - circle_fme_count["Handeled FME"]
circle_ztm_count = df_fme.groupby('circle')['manager_msisdn'].nunique().reset_index(name="Total CEM")
df_svd_count_fme = latest1daysdump_fme.groupby('circle')['WO Number'].count().reset_index(name="SVD WOS FME")
df_svd_count_cem = latest1daysdump_cem.groupby('circle')['WO Number'].count().reset_index(name="SVD WOS CEM")


circle_fme_count = circle_fme_count.merge(circle_ztm_count, on='circle', how='left').fillna(0)
circle_fme_count = circle_fme_count.merge(df_svd_count_fme, on='circle', how='left').fillna(0)
circle_fme_count = circle_fme_count.merge(df_svd_count_cem, on='circle', how='left').fillna(0)

# Calculate the sum for numeric columns
total_row = circle_fme_count.select_dtypes(include="number").sum().to_frame().T

# Add a label to the first column (assuming the first column is categorical)
first_col = circle_fme_count.columns[0]
total_row[first_col] = "Total"

# Ensure column order remains the same
total_row = total_row[circle_fme_count.columns]

# Append the total row
circle_fme_count = pd.concat([circle_fme_count, total_row], ignore_index=True)

circle_fme_count["Productivity"] = (circle_fme_count["SVD WOS FME"] / circle_fme_count["Total FME"]).round(1)

# Identify numeric columns except 'Productivity'
cols_to_convert = circle_fme_count.select_dtypes(include="number").columns.difference(["Productivity"])

# Convert to integer
circle_fme_count[cols_to_convert] = circle_fme_count[cols_to_convert].astype(int)

#==============================================SVD travel and (activity time)===============================
# Create Pivot Table
SVD_Travel_summary = latest1daysdump.pivot_table(
    index='circle', 
    columns='Travel Time Category', 
    values='WO Number',  
    aggfunc='count',
    fill_value= 0,
).reset_index()



SVD_activity_time_summary = latest1daysdump.pivot_table(
    index='circle', 
    columns='Activity Time Category', 
    values='WO Number',  
    aggfunc='count',
    fill_value= 0,
).reset_index()



SVD_Travel_summary = SVD_Travel_summary.merge(SVD_activity_time_summary, on='circle', how='left')


SVD_Travel_summary = SVD_Travel_summary.merge(df_svd_count_fme, on='circle', how='left')

SVD_Travel_summary = SVD_Travel_summary.merge(df_svd_count_cem, on='circle', how='left')
SVD_Travel_summary = SVD_Travel_summary.fillna(0)
SVD_Travel_summary["Total SVD WOS"] = (SVD_Travel_summary["SVD WOS FME"] + SVD_Travel_summary["SVD WOS CEM"])


# Calculate the sum for numeric columns
total_row = SVD_Travel_summary.select_dtypes(include="number").sum().to_frame().T

# Add a label to the first column (assuming the first column is categorical)
first_col = SVD_Travel_summary.columns[0]
total_row[first_col] = "Total"

# Ensure column order remains the same
total_row = total_row[SVD_Travel_summary.columns]

# Append the total row
SVD_Travel_summary = pd.concat([SVD_Travel_summary, total_row], ignore_index=True)

#=========================================================MTTR===============================

# Step 1: Filter for the latest 1 day from combined_data_all
latest1daysdump_all = combined_data_all[combined_data_all['WO Date'] > (latest_date - pd.Timedelta(days=1))]

# Step 2: Create a helper function to generate the pivot
def create_priority_pivot(df):
    # Convert to timedelta (safe conversion)
    df['Activity Time'] = pd.to_timedelta(df['Activity Time'], errors='coerce')
    
    # Drop rows where conversion failed (optional: or handle it as per logic)
    df = df.dropna(subset=['Activity Time'])

    pivot = df.pivot_table(
        index='circle',
        columns='Priority',
        values='Activity Time',
        aggfunc='mean',
        fill_value=np.nan        
    ).reset_index()
    

   # Step 4: Format timedelta to HH:MM:SS properly
    def format_hhmmss(td):
        if pd.isna(td):
            return ''
        total_seconds = int(td.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}:{seconds:02}"

    time_columns = pivot.columns.drop('circle')
    pivot[time_columns] = pivot[time_columns].applymap(format_hhmmss)
    
    return pivot
  

# Step 3: Create all required pivot tables
mttr_overall = create_priority_pivot(combined_data_all)
mttr_svd_all = create_priority_pivot(df_svd)
mttr_latestoneday_all = create_priority_pivot(latest1daysdump_all)
mttr_latestoneday_svd = create_priority_pivot(latest1daysdump)

# Step 4: Rename columns in each pivot to prepare for merge
mttr_overall = mttr_overall.add_suffix('_all')
mttr_overall = mttr_overall.rename(columns={'circle_all': 'circle'})

mttr_svd_all = mttr_svd_all.add_suffix('_svd')
mttr_latestoneday_all = mttr_latestoneday_all.add_suffix('_1day_all')
mttr_latestoneday_svd = mttr_latestoneday_svd.add_suffix('_1day_svd')

# Step 5: Merge all pivots on 'circle'
merged = mttr_overall.merge(mttr_svd_all, left_on='circle', right_on='circle_svd', how='left')
merged = merged.merge(mttr_latestoneday_all, left_on='circle', right_on='circle_1day_all', how='left')
merged = merged.merge(mttr_latestoneday_svd, left_on='circle', right_on='circle_1day_svd', how='left')

# Step 6: Drop extra circle columns from merged DataFrame
merged.drop(columns=['circle_svd', 'circle_1day_all', 'circle_1day_svd'], inplace=True)

#============================================= 11 AM Daily Summary=====================

latest5daysdump_11AM = grth11AM[grth11AM['In Progress on site Date'] >= (latest_date - pd.Timedelta(days=4))]
latest5daysdump_11AM_summary = latest5daysdump_11AM.pivot_table(
    index='circle', 
    columns='In Progress on site Date', 
    values='msisdn',  
    aggfunc='count',
    fill_value=0,
).reset_index()
# Format Date Columns Properly
latest5daysdump_11AM_summary.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in latest5daysdump_11AM_summary.columns
]
lat5days11am = total_fme_count.merge(latest5daysdump_11AM_summary, on='circle', how='left')

latest5daysdump_nosvd = svd_matched_fme[svd_matched_fme['In Progress on site Date'] >= (latest_date - pd.Timedelta(days=4))]
latest5daysdump_nosvd_summary = latest5daysdump_nosvd.pivot_table(
    index='circle', 
    columns='In Progress on site Date', 
    values='msisdn',  
    aggfunc=lambda x: x.nunique(),
    fill_value=0,
).reset_index()
# Format Date Columns Properly
latest5daysdump_nosvd_summary.columns = [
    col.strftime('%d %b') if isinstance(col, pd.Timestamp) else col 
    for col in latest5daysdump_nosvd_summary.columns
]

lat5daysnosvd = total_fme_count.merge(latest5daysdump_nosvd_summary, on='circle', how='left')

# Get all columns that are in date format (e.g., '18 Apr', '19 Apr', etc.)
date_columns = [col for col in lat5daysnosvd.columns if col not in ['circle', 'Total FME']]

# Subtract each date column from 'Total FME'
for col in date_columns:
    lat5daysnosvd[f'{col}_Nosvd'] = lat5daysnosvd['Total FME'] - lat5daysnosvd[col]

lat5daysnosvd.drop(columns=date_columns + ['Total FME'], inplace=True)
lat5days11am = lat5days11am.merge(lat5daysnosvd, on='circle', how='left')

# Calculate the sum for numeric columns
total_row = lat5days11am.select_dtypes(include="number").sum().to_frame().T

# Add a label to the first column (assuming the first column is categorical)
first_col = lat5days11am.columns[0]
total_row[first_col] = "Total"

# Ensure column order remains the same
total_row = total_row[lat5days11am.columns]

# # Append the total row
lat5days11am = pd.concat([lat5days11am, total_row], ignore_index=True)

# for date in date_columns:
#     nosvd_col = f'{date}_Nosvd'
#     calc_col = f'{date}_11AM_%'
    
#     # Avoid division by zero
#     denominator = lat5days11am['Total FME'] - lat5days11am[nosvd_col]
#     lat5days11am[calc_col] = lat5days11am[date] / denominator.replace(0, np.nan)

#     # Optional: Format as percent and round
#     lat5days11am[calc_col] = (lat5days11am[calc_col] * 100).round(0).astype('Int64').astype(str) + '%'


# for date in date_columns:
#     nosvd_col = f'{date}_Nosvd'
#     calc_col = f'{date}_No_SVD_%'
    
#     # Avoid division by zero
    
#     lat5days11am[calc_col] = lat5days11am[nosvd_col] / lat5days11am['Total FME']

#     # Optional: Format as percent and round
#     lat5days11am[calc_col] = (lat5days11am[calc_col] * 100).round(0).astype('Int64').astype(str) + '%'

# ---------------------- 1️⃣  Calculate {date}_11AM_% ----------------------
for date in date_columns:
    nosvd_col = f'{date}_Nosvd'
    calc_col = f'{date}_11AM_%'
    
    # Convert safely to numeric
    lat5days11am['Total FME'] = pd.to_numeric(lat5days11am['Total FME'], errors='coerce').fillna(0)
    lat5days11am[nosvd_col] = pd.to_numeric(lat5days11am[nosvd_col], errors='coerce').fillna(0)
    lat5days11am[date] = pd.to_numeric(lat5days11am[date], errors='coerce').fillna(0)
    
    # Avoid division by zero
    denominator = lat5days11am['Total FME'] - lat5days11am[nosvd_col]
    denominator = denominator.replace(0, np.nan)
    
    lat5days11am[calc_col] = (lat5days11am[date] / denominator) * 100
    
    # Replace NaN or invalid values with 0 and format as %
    lat5days11am[calc_col] = (
        lat5days11am[calc_col]
        .fillna(0)
        .round(0)
        .astype(int)
        .astype(str) + '%'
    )

# ---------------------- 2️⃣  Calculate {date}_No_SVD_% ----------------------
for date in date_columns:
    nosvd_col = f'{date}_Nosvd'
    calc_col = f'{date}_No_SVD_%'
    
    # Safe numeric conversion
    lat5days11am['Total FME'] = pd.to_numeric(lat5days11am['Total FME'], errors='coerce').fillna(0)
    lat5days11am[nosvd_col] = pd.to_numeric(lat5days11am[nosvd_col], errors='coerce').fillna(0)
    
    # Avoid division by zero
    lat5days11am[calc_col] = (lat5days11am[nosvd_col] / lat5days11am['Total FME'].replace(0, np.nan)) * 100
    
    # Replace NaN or invalid values with 0 and format as %
    lat5days11am[calc_col] = (
        lat5days11am[calc_col]
        .fillna(0)
        .round(0)
        .astype(int)
        .astype(str) + '%'
    )












#==============================================SLA %===============================

#==============================================WO Acceptence==============================


#==============================================Backlog===============================

#==============================================Repeat Site Visit===========================

#==============================================SFN===============================

#==============================================PRODUCTIVITY (Zone wise)===================

#==============================================SLA Zone wise===============================


#===============================================Creating Dashboard==========================
#max_rows_per_sheet = 1048570
# Get the current date in YYYYMMDD format
#current_date = datetime.now().strftime("%Y%m%d")
# Define output file path with the current date as a suffix

report_date_wo_dump = combined_data_all['WO Date'].max().strftime('%d-%m-%Y')

report_date_wo_dump_max = combined_data_all['WO Date'].max().strftime('%d %b')
report_date_wo_dump_min = combined_data_all['WO Date'].min().strftime('%d %b')
report_date_txn_wo_min = df_txn_wo_last15['WO Date'].min().strftime('%d %b')


output_file_path = rf"D:\Automation\FME_Ranking\output\FME_Ranking_Dashboard_{report_date_wo_dump}.xlsx"


with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer: 
    fme_svd_bucket.to_excel(writer, sheet_name="Summary", index=False, startcol=1, startrow=2)
    circlewise_svd_cat_count_non_cat.to_excel(writer, sheet_name="Summary", index=False, startcol=9, startrow=2)
    fme_unique_site_prod.to_excel(writer, sheet_name="Summary", index=False, startcol=17, startrow=2)
    lat5days11am.to_excel(writer, sheet_name="11 AM Summary", index=False, startrow=1)
    df_qia_repeat_daywise.to_excel(writer, sheet_name=">10 WO Summary", index=False,startcol=0, startrow=1)
    df_outage_repeat_daywise.to_excel(writer, sheet_name=">10 WO Summary", index=False,startcol=0, startrow=29)
    df_non_qia_repeat_daywise.to_excel(writer, sheet_name=">10 WO Summary", index=False,startcol=0, startrow=56)


    df_fme_prod.to_excel(writer, sheet_name="FME wise Productivity", index=False)
    df_unique_site_visit.to_excel(writer, sheet_name="FME Unique Site Visit", index=False)
    #sitewise_svd_count.to_excel(writer, sheet_name="Site wise SVD Count", index=False) 
    sitewise_svd_non_cat_count.to_excel(writer, sheet_name="Site wise SVD Count", index=False) 
      
    df_fme_11am_backup.to_excel(writer, sheet_name="11 AM FME Backup", index=False)
    
    working_day_summary.to_excel(writer, sheet_name="FME_Prod_Summary", index=False, startrow=1)
    circle_site_count.to_excel(writer, sheet_name="Circle wise Creation", index=False)
    circle_fme_count.to_excel(writer, sheet_name="Circle Productivity", index=False)
    df_trip_done_by.to_excel(writer, sheet_name="Trip Done By", index=False)
    
    df_repeat_wo_corrective.to_excel(writer, sheet_name="Repeat WO", index=False)
    #df_outage_dump.to_excel(writer, sheet_name="Outage Dump", index=False)
    #SVD_Travel_summary.to_excel(writer, sheet_name="SVD Travel_Act_Time", index=False, startrow=1)
    #merged.to_excel(writer, sheet_name="MTTR", index=False, startrow=2)
    #latest5daysdump_nosvd_summary.to_excel(writer, sheet_name="Latest 5 Days No SVD", index=False, startrow=1)
    #df_svd.to_excel(writer, sheet_name="combined_data", index=False)
    #df_svd_non_cat.to_excel(writer, sheet_name="SVD Non Cat", index=False)
    #sitewise_svd_count_non_cat.to_excel(writer, sheet_name="SVD Count Non Cat", index=False)
    #df_fme.to_excel(writer, sheet_name="FME WO Dump", index=False)





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
    worksheet.merge_range(1, 1, 1, 7, 'Sites Visited -'+ report_date_wo_dump_max, merge_format)
    worksheet.merge_range(1, 9, 1, 15, 'Repeat Site visit (' + report_date_wo_dump_min + " to " + report_date_wo_dump_max +")", merge_format)
    worksheet.merge_range(1, 17, 1, 25, 'FME Productivity in month (' + report_date_wo_dump_min + " to " + report_date_wo_dump_max +")", merge_format)

     # Apply formatting
    workbook  = writer.book
    worksheet = writer.sheets["11 AM Summary"]

    # Merge cells with labels
    worksheet.merge_range(0, 0, 0, 1, '11 AM Dashboard', merge_format)
    worksheet.merge_range(0, 2, 0, 6, 'No of FMEs reaching first site after 11AM', merge_format)
    worksheet.merge_range(0, 7, 0, 11, 'FMEs with Zero SVD', merge_format)
    worksheet.merge_range(0, 12, 0, 16, "%age of Total FME after 11AM", merge_format)
    worksheet.merge_range(0, 17, 0, 21, "%age FME with Zero SVD", merge_format)

    # Apply formatting
    workbook  = writer.book
    worksheet = writer.sheets[">10 WO Summary"]

    # Merge cells with labels
    worksheet.merge_range(0, 0, 0, 31, '> 10 Work Orders generated Site count in QIA', merge_format)
    worksheet.merge_range(28, 0, 28, 31, '> 10 Work Orders generated Site count in Outage', merge_format)
    worksheet.merge_range(55, 0, 55, 31, '> 10 Work Orders generated Site count in Non QIA', merge_format)
  
    



writer.close()

print("Data saved successfully to multiple sheets in:",output_file_path)


#===============================================Creating Dashboard for TXN Work Order Trend==========================
output_file_path_txn = rf"D:\Automation\FME_Ranking\output\TXN_Work_Order_Trend_{report_date_wo_dump}.xlsx"


with pd.ExcelWriter(output_file_path_txn, engine='xlsxwriter') as writer: 
    df_txn_wo_summary.to_excel(writer, sheet_name="TXN WO Summary", index=False,startcol=1, startrow=2)
    df_txn_wo_summary_daywise.to_excel(writer, sheet_name="TXN WO Summary", index=False,startcol=6, startrow=2)
    df_txn_wo_category.to_excel(writer, sheet_name="Category Wise Trend", index=False)
    df_txn_wo_last15.to_excel(writer, sheet_name="Dump", index=False)
    
    

    # Apply formatting
    workbook  = writer.book
    worksheet = writer.sheets["TXN WO Summary"]

    merge_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    # Merge cells with labels
    worksheet.merge_range(1, 1, 1, 4, '15/30Days', merge_format)
    worksheet.merge_range(1, 6, 1, 22, 'Daywise Trend', merge_format)




writer.close()

print("Data saved successfully to multiple sheets in:",output_file_path_txn)

#===============================================Mail Body Calculations====================

high_repeat_visits = circlewise_svd_cat_count_non_cat[
    (circlewise_svd_cat_count_non_cat['circle'] != 'Total') &
    (circlewise_svd_cat_count_non_cat[["3 Visits", "4-7 Visits", ">7 Visits"]].sum(axis=1) > 150)
]
high_repeat_circles = high_repeat_visits['circle'].tolist()


# high_productivity_fmes = fme_unique_site_prod[
#     (fme_unique_site_prod['circle'] != 'Total') &
#     (fme_unique_site_prod[["4-5 SVD", "5-6 SVD"]].sum(axis=1) > 5)
# ]


# Define the columns you want to consider
svd_columns = ["4-5 SVD", "5-6 SVD"]

# Filter only the columns that exist in the DataFrame
existing_svd_columns = [col for col in svd_columns if col in fme_unique_site_prod.columns]

# Apply filter only if there is at least one valid column
if existing_svd_columns:
    high_productivity_fmes = fme_unique_site_prod[
        (fme_unique_site_prod['circle'] != 'Total') &
        (fme_unique_site_prod[existing_svd_columns].sum(axis=1) > 5)
    ]
else:
    high_productivity_fmes = pd.DataFrame()  # or handle it differently



high_prod_circles = high_productivity_fmes['circle'].tolist()


low_productivity_fmes = fme_unique_site_prod[
    (fme_unique_site_prod['circle'] != 'Total') &
    (fme_unique_site_prod["0-1 SVD"] > 5)
]
low_prod_circles = low_productivity_fmes['circle'].tolist()


# Exclude 'Total' row
lat5days_filtered = lat5days11am[lat5days11am['circle'] != 'Total']

# Get the last column (assuming it's the percentage of FMEs with zero SVD)
last_col = lat5days_filtered.columns[-1]

# Convert % values to numeric (strip %, handle strings)
lat5days_filtered[last_col] = lat5days_filtered[last_col].str.rstrip('%').astype(float)

# Filter circles where percentage is >= 15%
zero_svd_circles = lat5days_filtered[lat5days_filtered[last_col] >= 15]['circle'].tolist()

# Get the last column (assuming it's the percentage of FMEs with zero SVD)
last_6th_col = lat5days_filtered.columns[-6]

# Convert % values to numeric (strip %, handle strings)
lat5days_filtered[last_6th_col] = lat5days_filtered[last_6th_col].str.rstrip('%').astype(float)

# Filter circles where percentage is >= 15%
svd_circles = lat5days_filtered[lat5days_filtered[last_6th_col] >= 30]['circle'].tolist()



# Prepare comma-separated strings for email
high_repeat_circles_str = ", ".join(high_repeat_circles)
high_prod_circles_str = ", ".join(high_prod_circles)
low_prod_circles_str = ", ".join(low_prod_circles)
zero_svd_circles_str = ", ".join(zero_svd_circles)
svd_circles_str = ", ".join(svd_circles)

#===================================Printing Tables in mail body===================
def generate_html_table_with_header(df, title_text):
    # Convert DataFrame to HTML without index

    df.columns.name = None
    df.rename(columns={'circle': 'Circle'}, inplace=True)
    df_html = df.to_html(index=False, border=1)

    # Create merged header row with dynamic colspan
    merged_header_row = f"""
    <tr>
      <th colspan="{len(df.columns)}" style="text-align:center; background-color:#D9D9D9; font-weight:bold;">
        {title_text}
      </th>
    </tr>
    """

    # Clean pandas-generated HTML table
    df_html_clean = df_html.replace(
        '<table border="1" class="dataframe">', ''
    ).replace('</table>', '')

    # Final HTML Table with header row
    final_html_table = f"""
    <table border="1" cellspacing="0" cellpadding="5" style="border-collapse:collapse; font-family: Arial, sans-serif; font-size: 12px;">
      {merged_header_row}
      {df_html_clean}
    </table>
    """
    return final_html_table


# Example for fme_svd_bucket
fme_svd_table_html = generate_html_table_with_header(
    fme_svd_bucket,
    "Sites Visited Summary (Last 1 Day)"
)

# Example for circlewise_svd_cat_count
fme_unique_site_prod_html = generate_html_table_with_header(
    fme_unique_site_prod,
    "FME Productivity MTD"
)

# Example for circlewise_svd_cat_count
circlewise_svd_table_html = generate_html_table_with_header(
    circlewise_svd_cat_count_non_cat,
    "Repeat Site visit MTD"
)


# Example for circlewise_svd_cat_count
lat5days11am_html = generate_html_table_with_header(
    lat5days11am,
    "11 AM Summary"
)

# Example for circlewise_svd_cat_count
circle_fme_count_html = generate_html_table_with_header(
    circle_fme_count,
    "Circle Productivity (Last 1 Day)"
)

# #=================================sending Mail=====================================

# Read recipients from Excel
file_path_receipient = r"D:\Automation\FME_Ranking\Receipient.xlsx"
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

# Get today's date for email subject
today_date = datetime.now().strftime("%d%m%Y")
yesterday_date = (datetime.now() - timedelta(days=1)).strftime("%d%m%Y")


email_subject = f"FME Productivity, 11 AM Summary, Daily Repeat WOs >10WO & Repeat Site Visits - {yesterday_date}"
email_body = f"""
<html>
  <body>
    <p>Dear All,</p>

    <p>
      Please find attached the FME Productivity Report, 11 AM Summary, and Repeat Site Visits from 
      <b>{report_date_wo_dump_min}</b> to <b>{report_date_wo_dump_max}</b>.
    </p>

    <p>
      Kindly check for:
      <ul>
        <li>Repeat site visits</li>
        <li>Daily Repeat WOs &gt; 10 WO</li>
        <li>11 AM Summary for <b>{report_date_wo_dump_max}</b></li>
        <li>
          FMEs with 
          <span style="color:red;">Very Low (&lt;1)</span> and 
          <span style="color:red;">Exceptionally High (&gt;4)</span> productivity
        </li>
      </ul>
    </p>    

    <p><b>Highlights:</b></p>
    <ul>
      <li><b>Repeat site visits (≥3 times/rolling month &gt;150 sites)</b> in: <i><span style="color:red;">{high_repeat_circles_str}</span></i>.</li>
      <li><b>Very High Productivity FMEs (&gt;5 FMEs with 4–6 SVD):</b> <i><span style="color:red;">{high_prod_circles_str}</span></i>.</li>
      <li><b>Very Low Productivity FMEs (&gt;5 FMEs with &lt;1 SVD):</b> <i><span style="color:red;">{low_prod_circles_str}</span></i>.</li>
      <li><b>≥30% FMEs reached their first site after 11 AM on {report_date_wo_dump_max}:</b> <i><span style="color:red;">{svd_circles_str}</span></i>.</li>
      <li><b>Circles with ≥15% FMEs having Zero SVD on {report_date_wo_dump_max}:</b> <i><span style="color:red;">{zero_svd_circles_str}</span></i>.</li>
    </ul>

    <p>
      This report will be shared daily for the previous day's site visits and productivity to help take corrective actions.
    </p>

    <table cellspacing="10" cellpadding="5" style="border-collapse:collapse; font-family: Arial, sans-serif; font-size: 12px;">
      <tr>
        <td valign="top">
          {fme_svd_table_html}
        </td>

        <td valign="top">
          {fme_unique_site_prod_html}
        </td>

        <td valign="top">
          {circlewise_svd_table_html}
        </td>

        <td valign="top">
          {circle_fme_count_html}
        </td>
      </tr>

    </table>

    <table cellspacing="10" cellpadding="5" style="border-collapse:collapse; font-family: Arial, sans-serif; font-size: 12px;">
    <tr>
        <td valign="top">
          {lat5days11am_html}
        </td>
      </tr>
    </table>
    
    <p>Regards,<br>Central Team FSO</p>

  </body>
</html>
"""

try:
    # Create a new email
    mail = outlook.CreateItem(0)
    mail.To = ";".join(to_recipients) if to_recipients else ""  # Join multiple recipients with ";"
    mail.CC = ";".join(cc_recipients) if cc_recipients else ""  # Join multiple CC recipients
    mail.Subject = email_subject
    mail.HTMLBody = email_body
    #mail.Body = email_body

    # Specify the sender email
    mail.SentOnBehalfOfName = from_email  # Option 1 (if account has permission)

    # Alternative approach (set the sender account explicitly)
    for account in outlook.Session.Accounts:
        if account.SmtpAddress == from_email:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # Option 2 (SendUsingAccount)

    # Attach the file
    mail.Attachments.Add(output_file_path)
    mail.Send()  # Send the email

    print(f"📧 Email sent successfully from {from_email}")

except Exception as e:
    print(f"❌ Error sending email: {e}")

# Release Outlook objects
del outlook, namespace

print("✅ Email process completed successfully!")
# #=============================================End of Code=================================================


# #=================================sending Mail=====================================

# Read recipients from Excel
file_path_receipient = r"D:\Automation\FME_Ranking\Receipient_txn.xlsx"
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

# Get today's date for email subject
today_date = datetime.now().strftime("%d%m%Y")
yesterday_date = (datetime.now() - timedelta(days=1)).strftime("%d%m%Y")


email_subject = f"TXN Work Order Trend - {yesterday_date}"
email_body = f"""
<html>
  <body>
    <p>Dear All,</p>

    <p>
      Please find attached the TXN Work Order Trend from 
      <b>{report_date_txn_wo_min}</b> to <b>{report_date_wo_dump_max}</b>.
    </p>    
    
    <p>Regards,<br>Central Team FSO</p>

  </body>
</html>
"""

try:
    # Create a new email
    mail = outlook.CreateItem(0)
    mail.To = ";".join(to_recipients) if to_recipients else ""  # Join multiple recipients with ";"
    mail.CC = ";".join(cc_recipients) if cc_recipients else ""  # Join multiple CC recipients
    mail.Subject = email_subject
    mail.HTMLBody = email_body
    #mail.Body = email_body

    # Specify the sender email
    mail.SentOnBehalfOfName = from_email  # Option 1 (if account has permission)
 
    # Alternative approach (set the sender account explicitly)
    for account in outlook.Session.Accounts:
        if account.SmtpAddress == from_email:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # Option 2 (SendUsingAccount)

    # Attach the file
    mail.Attachments.Add(output_file_path_txn)
    mail.Send()  # Send the email

    print(f"📧 Email sent successfully from {from_email}")

except Exception as e:
    print(f"❌ Error sending email: {e}")

# Release Outlook objects
del outlook, namespace

print("✅ Email process completed successfully!")
#=============================================End of Code=================================================


