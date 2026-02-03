import pandas as pd
import numpy as np
import os
import warnings
from datetime import timedelta

# Define the text
text = """


                           ##############           #################                  ##################
                          ###############           ##################                ####################
                        ###                         ###             ###               ###              ###
                        ###                         ###             ###               ###              ###
                        ###                         ###             ###                 ###
                        ###                         #################                       ###   
                        ###                         #################                            ###
                        ###                         ###             ###                               ###
                        ###                         ###             ###               ###               ###
                        ###                         ###             ###               ###               ###
                          ###############           ##################                #####################
                           ##############           #################                  ###################

# ----------------------------------------------------------------------------------------------------------------------------------
# Copyright (c) 2024 Bharti Airtel Limited. All Rights Reserved.
# 
# This script is intended for authorized users only. Unauthorized access or 
# use of this script is prohibited. By logging in, you agree to the terms and 
# conditions set forth by Bharti Airtel Limited. 
#
# For support or inquiries, contact Bharti Airtel Limited at Mob No. +91-9931888288. Email ID: kaushal3.kumar@airtel.com
# ----------------------------------------------------------------------------------------------------------------------------------

"""

# Print the text
print(text)

# Suppress UserWarnings from openpyxl
warnings.simplefilter("ignore", category=UserWarning)

#=================================================Creating Sample Dataset===========================================
# Define the data
data = {
    "Current Tech Id": ["Test"] * 10,
    "Date On Air": ["05/08/1991  12:00:00 AM"] * 10,
    "MS1-Date": ["05/08/1991  12:00:00 AM"] * 10,
    "MS2-Date": ["05/08/1991  12:00:00 AM"] * 10,
    "Active /Locked/Dismantled": ["Active"] * 10,
    "Existing/New/Relocation": ["New", "New", "New", "Existing", "Existing", "New", "Relocation", "Relocation", "Relocation","Relocation"],
    "MS Status": ["MS1", "MS1", "MS1", "MS1", "MS1", "MS2", "MS2", "MS2", "MS2", "MS2"],
    "Frequency Band(L900, L1800, L2100, TD2300)": ["900", "1800", "900+1800", "L850", "L900","L1800","L2100","L2300", "TD2300","L2100"],
    "Source File": ["NSN_2G_CBS_A$GJ@20240625.xlsx"] *10,
    "Sheet Name": ["CBS_A"] *10,
    "Date Since Locked": ["8-May-91"] * 10,    
    "Relocation OLD SiteId": ["Test1"] * 10    
}

# Create the DataFrame
test_df = pd.DataFrame(data)
test_df['Date On Air'] = pd.to_datetime(test_df['Date On Air'], errors='coerce')
test_df['MS1-Date'] = pd.to_datetime(test_df['MS1-Date'], errors='coerce')
test_df['MS2-Date'] = pd.to_datetime(test_df['MS2-Date'], errors='coerce')

#==============================================Reading 2G CBS Files=================================================
# Define the folder containing the Excel files
folder_path = 'D:\\Automation\\CBS_Auto\\Dump\\2G'  # Replace with your folder path

# List of columns to extract
required_columns = [
    'Current Tech Id.', 
    'Date On Air', 
    'MS1-Date', 
    'MS2-Date', 
    'Active /Locked/Dismantled', 
    'Date Since Locked', 
    'Existing/New/Relocation', 
    'Relocation OLD SiteId', 
    'Site MAPA Status', 
    'Frequency BAND(900,1800,900+1800)'
]

# Initialize an empty list to store DataFrames
dataframes_2g = []

# Loop through all files in the folder
for file in os.listdir(folder_path):
    # Check if the file is an Excel file
    if file.endswith('.xlsx') or file.endswith('.xls'):
        file_path = os.path.join(folder_path, file)
        
        try:
            
            # Read the required sheets into DataFrames
            df_cbs_a = pd.read_excel(file_path, sheet_name='CBS_A', skiprows = 1)
            df_cbs_l = pd.read_excel(file_path, sheet_name='CBS_L', skiprows = 1)

              # Print the file name, sheet name, and size
            print(f"Reading File: {file}")
            print(f"Sheet Name: CBS_A, Size: {df_cbs_a.shape}")
            print(f"Sheet Name: CBS_L, Size: {df_cbs_l.shape}")
            
            # Add a source column to identify data origin
            df_cbs_a['Source File'] = file
            df_cbs_a['Sheet Name'] = 'CBS_A'
            
            df_cbs_l['Source File'] = file
            df_cbs_l['Sheet Name'] = 'CBS_L' 

            # Extract only the required columns from both DataFrames
            df_cbs_a_filtered = df_cbs_a[required_columns + ['Source File', 'Sheet Name']]
            df_cbs_l_filtered = df_cbs_l[required_columns + ['Source File', 'Sheet Name']]

            # Append the filtered DataFrames to the list
            dataframes_2g.append(df_cbs_a_filtered)
            dataframes_2g.append(df_cbs_l_filtered)
        
        except Exception as e:
            print(f"Error reading {file}: {e}")
            
# Filter out empty or all-NA columns from each DataFrame before concatenation
dataframes_2g = [df.dropna(axis=1, how='all') for df in dataframes_2g]

# Combine all DataFrames into one
if dataframes_2g:
    combined_df_2g = pd.concat(dataframes_2g, ignore_index=True)
    print("2G Data combined successfully!")
else:
    combined_df_2g = pd.DataFrame()
    print("No valid 2G data found.")

# Rename columns to match the required format
combined_df_2g.rename(columns={
    'Frequency BAND(900,1800,900+1800)' : 'Frequency Band(L900, L1800, L2100, TD2300)',
    'Current Tech Id.': 'Current Tech Id', 
    'Site MAPA Status' : 'MS Status',
}, inplace=True)

#==============================================Reading 4G CBS Files=================================================
# Define the folder containing the Excel files
folder_path = 'D:\\Automation\\CBS_Auto\\Dump\\4G'  # Replace with your folder path


# List of columns to extract
required_columns = [
    'Current Tech Id',  
    'Date On Air', 
    'MS1-Date', 
    'MS2-Date', 
    'Active /Locked/Dismantled', 
    'Date Since Locked', 
    'Existing/New/Relocation', 
    'Relocation OLD SiteId', 
    'Site MAPA status', 
    'Frequency Band(L900, L1800, L2100, TD2300)'
    
]

# Initialize an empty list to store DataFrames
dataframes_4g = []

# Loop through all files in the folder
for file in os.listdir(folder_path):
    # Check if the file is an Excel file
    if file.endswith('.xlsx') or file.endswith('.xls'):
        file_path = os.path.join(folder_path, file)
        
        try:
            
            # Read the required sheets into DataFrames
            df_cbs_a = pd.read_excel(file_path, sheet_name='CBS_A', skiprows = 1)
            df_cbs_l = pd.read_excel(file_path, sheet_name='CBS_L', skiprows = 1)

            # Print the file name, sheet name, and size
            print(f"Reading File: {file}")
            print(f"Sheet Name: CBS_A, Size: {df_cbs_a.shape}")
            print(f"Sheet Name: CBS_L, Size: {df_cbs_l.shape}")
            
            # Add a source column to identify data origin
            df_cbs_a['Source File'] = file
            df_cbs_a['Sheet Name'] = 'CBS_A'
            
            df_cbs_l['Source File'] = file
            df_cbs_l['Sheet Name'] = 'CBS_L' 

            # Extract only the required columns from both DataFrames
            df_cbs_a_filtered = df_cbs_a[required_columns + ['Source File', 'Sheet Name']]
            df_cbs_l_filtered = df_cbs_l[required_columns + ['Source File', 'Sheet Name']]

            # Append the filtered DataFrames to the list
            dataframes_4g.append(df_cbs_a_filtered)
            dataframes_4g.append(df_cbs_l_filtered)
        
        except Exception as e:
            print(f"Error reading {file}: {e}")

# Filter out empty or all-NA columns from each DataFrame before concatenation
dataframes_4g = [df.dropna(axis=1, how='all') for df in dataframes_4g]

# Combine all DataFrames into one
if dataframes_4g:
    combined_df_4g = pd.concat(dataframes_4g, ignore_index=True)
    print("4G Data combined successfully!")
else:
    combined_df_4g = pd.DataFrame()
    print("No valid 4G data found.")

# Rename columns to match the required format
combined_df_4g.rename(columns={ 
    'Site MAPA status' :'MS Status'
}, inplace=True)

#=============================================Concatinating 2G and 4G Dump==========================================

combined_df = pd.concat([combined_df_2g,combined_df_4g,test_df], ignore_index = True)
# Assuming 'Current Tech Id' is the column name
combined_df = combined_df[combined_df['Current Tech Id'].notna() & (combined_df['Current Tech Id'] != '')]


print(f"2G (CBS_A, CBS_L) and 4G (CBS_A, CBS_L) Sheets Combined Sussessfully Size: {combined_df.shape}")
print(combined_df.shape)
print(f"Defining Frequency Band")

#Defining Frequnecy Band
def assign_frequency_band(row):
    frequency_band = str(row['Frequency Band(L900, L1800, L2100, TD2300)'])  # Ensure the value is a string
    
    # Check for 2G frequency bands first
    if '900' in frequency_band and 'L900' not in frequency_band:
        return '2G'  # 900 as 2G
    elif '900+1800' in frequency_band or '1800' in frequency_band and 'L1800' not in frequency_band:
        return '2G'  # 1800 as 2G
    
      # Check for L900 separately
    elif 'L850' in frequency_band:
        return 'L850'
    
    # Check for L900 separately
    elif 'L900' in frequency_band:
        return 'L900'
    
    # Check for L1800 separately
    elif 'L1800' in frequency_band:
        return 'L1800'
    
    # Check for other frequency bands
    elif 'L2100' in frequency_band:
        return 'L2100'
    
    # Check for other frequency bands
    elif 'L2300' in frequency_band:
        return 'L2300'
    
    # Check for other frequency bands
    elif 'TD2300' in frequency_band:
        return 'L2300'
    
    return None
# # Apply the function to create the new column
combined_df['Frequency Band'] = combined_df.apply(assign_frequency_band, axis=1)

print(combined_df.shape)

print(f"Defining Circles from Source Excel Files")

patterns = [
    r'\$GJ@', r'\$KL@', r'\$MH@', r'\$MU@', r'\$OR@',
    r'\$PB@', r'\$UE@', r'\$WB@', r'BH', r'MP' ]
circles = ['GJ', 'KL', 'MH', 'MU', 'OR', 'PB', 'UE', 'WB', 'BH', 'MP']

# Create a condition for each pattern
conditions = [combined_df['Source File'].str.contains(pattern, regex=True) for pattern in patterns]

# Use np.select to assign values based on conditions
combined_df['Circle'] = np.select(conditions, circles, default=None)

print(combined_df.shape)

print(f"Defining Unique Current Tech ID")

#combined_df = combined_df[combined_df['Current Tech Id'].notna()]

# Convert 'Current Tech Id' to uppercase only if it's a string
combined_df['Current Tech Id'] = combined_df['Current Tech Id'].apply(
    lambda x: x.upper() if isinstance(x, str) else x
)

print(combined_df.shape)

print(f"Removing potential float issue like Site ID 9063.0 = 9063")
# Convert 'Relocation OLD SiteId' to string, removing potential float issue
combined_df['Current Tech Id'] = combined_df['Current Tech Id'].apply(lambda x: str(x).replace('.0', '') if x != '#N/A' else x)




#Creating Unique Current Tech ID
combined_df['Unique Current Tech ID'] = combined_df['Circle'].astype(str) + "_" + combined_df['Current Tech Id'].astype(str)

print(combined_df.shape)

print(f"Setting Minimum Dates against Unique Current Tech ID as per frequency band")
# Minimize merging by precomputing grouped values
min_dates = combined_df.groupby(['Unique Current Tech ID', 'Frequency Band'])['Date On Air'].transform('min')
combined_df['Date On Air'] = min_dates


# Replace manual date parsing with a direct assignment
report_date_input = input("Enter the Report Date (DD-MM-YYYY): ")
try:
    report_date = pd.to_datetime(report_date_input, format='%d-%m-%Y')
except ValueError:
    print("Invalid date format. Setting report date to None.")
    report_date = pd.NaT

combined_df['Report Date'] = report_date

combined_df['Date Since Locked'] = pd.to_datetime(combined_df['Date Since Locked'], errors='coerce')

print(f"Setting Blank Values to 08-05-1991 in columns 'MS1-Date', 'MS2-Date', 'Date Since Locked' ")
# Fill missing dates in one step
default_date = '1991-05-08'
combined_df[['MS1-Date', 'MS2-Date', 'Date Since Locked']] = combined_df[['MS1-Date', 'MS2-Date', 'Date Since Locked']].fillna(default_date)


print(combined_df.shape)

print(f"Setting Mapa Status to MS2 where any sector of that site in particular frequency Band is MS2")
combined_df['MS Status'] = combined_df.groupby(['Unique Current Tech ID','Frequency Band'])['MS Status'].transform(lambda x: 'MS2' if 'MS2' in x.values else 'MS1')

#Deleting records where MS Status' is Blank
#combined_df = combined_df[combined_df['MS Status'].notna()]

print(f"Calculating Days Difference.....")
combined_df['Days Difference'] = (combined_df['Report Date'] - combined_df['Date On Air']).dt.days

print(combined_df.shape)

print(f"Calculating A/L Column.....")
combined_df['A/L'] = combined_df.groupby('Unique Current Tech ID')['Active /Locked/Dismantled'].transform(
    lambda group: 'L' if 'Locked' in group.values else '#N/A'
)
#============================================Working on Relocation Site ID==========================================

print(combined_df.shape)

print(f"Converting Relocation OLD SiteId to uppercase only if it's a string ")
# Convert 'Relocation OLD SiteId' to uppercase only if it's a string
combined_df['Relocation OLD SiteId'] = combined_df['Relocation OLD SiteId'].apply(
    lambda x: x.upper() if isinstance(x, str) else x
)

print("Filling Blank values to #N/A in Relocation OLD SiteId Column")
combined_df['Relocation OLD SiteId'] = combined_df['Relocation OLD SiteId'].fillna('#N/A')


print(f"Removing potential float issue like Site ID 9063.0 = 9063")
# Convert 'Relocation OLD SiteId' to string, removing potential float issue
combined_df['Relocation OLD SiteId'] = combined_df['Relocation OLD SiteId'].apply(lambda x: str(x).replace('.0', '') if x != '#N/A' else x)



print("Filter out rows where 'Relocation OLD SiteId' is '#N/A' and concatenate")
# Filter out rows where 'Relocation OLD SiteId' is '#N/A' and concatenate
combined_df['Unique Relocation OLD SiteId'] = combined_df.apply(lambda row: f"{row['Circle']}_{row['Relocation OLD SiteId']}" if row['Relocation OLD SiteId'] != '#N/A' else '#N/A', axis=1)

print("Updating Relocation Site ID as per Frequency Band")

# Function to update 'Relocation OLD SiteId' based on groupings
def update_relocation_siteid(combined_df):
    # Group by 'Unique Current Tech ID' and 'Frequency Band'
    for current_tech_id, group in combined_df.groupby('Unique Current Tech ID'):
        for freq_band, freq_group in group.groupby('Frequency Band'):
            # Get the first 'Relocation OLD SiteId' for this 'Current Tech Id' and 'Frequency Band'
            first_site_id = freq_group['Relocation OLD SiteId'].iloc[0]
            # Update all 'Relocation OLD SiteId' in this frequency band to the first site ID
            combined_df.loc[freq_group.index, 'Relocation OLD SiteId'] = first_site_id
    return combined_df

# Apply the function to 'combined_df'
combined_df = update_relocation_siteid(combined_df)


print(combined_df.shape)

print ("Creating Column matched relocation in main site")
#Precompute the unique set of 'Unique Relocation OLD SiteId' for efficiency
relocation_ids = set(combined_df['Unique Relocation OLD SiteId'])

def check_relocation_status(row):
    # Check if the Unique Current Tech ID matches in the same row
    if row['Unique Current Tech ID'] == row['Unique Relocation OLD SiteId']:
        return "OK"
    # Check if the Unique Current Tech ID matches in any other row
    elif row['Unique Current Tech ID'] in relocation_ids and row['Active /Locked/Dismantled'] == 'Locked':
        return "NOT OK"
    # Default case: no match
    else:
        return "OK"

# Apply the function to each row
combined_df['matched relocation in main site'] = combined_df.apply(check_relocation_status, axis=1)




#==============================================Defining AMC Status================================================
print(f"Calculating AMC Status.........")
# Faster `AMC` assignment using a single function
def assign_amc_status(row):
    valid_bands = ['2G', 'L850', 'L900', 'L1800', 'L2100', 'L2300']
    mapa_status = ['MS1', 'MS2']
    is_active_non_relocation = (
        pd.notna(row["Unique Current Tech ID"]) and
        row["Active /Locked/Dismantled"] == "Active" and
        row["MS Status"] == "MS2" and
        row["Days Difference"] > 365 and
        row["Existing/New/Relocation"] != "Relocation" and
        row["Frequency Band"] in valid_bands and
        row["matched relocation in main site"] == "OK"
    )
    is_active_relocation = (
        pd.notna(row["Unique Current Tech ID"]) and
        row["Active /Locked/Dismantled"] == "Active" and
        row["MS Status"] in mapa_status and
        row["Existing/New/Relocation"] == "Relocation" and
        row["Frequency Band"] in valid_bands and
        row["matched relocation in main site"] == "OK"
    )
    is_locked_non_relocation = (
        pd.notna(row["Unique Current Tech ID"]) and
        row["Active /Locked/Dismantled"] == "Locked" and
        row["MS Status"] in mapa_status and
        row["Existing/New/Relocation"] != "Relocation" and
        report_date - timedelta(days=365) < row['Date Since Locked'] <= report_date and
        row["matched relocation in main site"] == "OK"
    )
    is_locked_relocation = (
        pd.notna(row["Unique Current Tech ID"]) and
        row["Active /Locked/Dismantled"] == "Locked" and
        row["MS Status"] in mapa_status and
        row["Existing/New/Relocation"] == "Relocation" and
        report_date - timedelta(days=365) < row['Date Since Locked'] <= report_date and
        row["matched relocation in main site"] == "OK"
    )
    return 'Y' if any([is_active_non_relocation, is_active_relocation, is_locked_non_relocation, is_locked_relocation]) else np.nan

combined_df["AMC"] = combined_df.apply(assign_amc_status, axis=1)

#================================================ Dashboard=========================================================


df_dashboard = combined_df[['Unique Current Tech ID', 'Circle', 'Current Tech Id', 'A/L', 'matched relocation in main site', 'Relocation OLD SiteId', 'Unique Relocation OLD SiteId','Date On Air', 'Date Since Locked', 'MS Status', 'Frequency Band', 'AMC']]
df_dashboard.loc[:,'matched relocation in main site'] = combined_df['matched relocation in main site'].fillna ('#N/A')

df_dashboard = combined_df[combined_df['matched relocation in main site'] != 'NOT OK' ]

df_dashboard.loc[:,'AMC'] = combined_df['AMC'].fillna ('#N/A')

print(f"Prepairing Dashboard with relocation site id.........")
# Applying Pivot Table
df_dashboard_pivot = pd.pivot_table(
    data=df_dashboard,  # Specify the DataFrame to pivot
    index=['Unique Current Tech ID', 'Circle','Current Tech Id', 'A/L', 'Relocation OLD SiteId'],  # Column(s) to use as the index
    columns=['Frequency Band'],  # Column(s) to use as the new columns
    values=['Date On Air', 'MS Status','AMC'],  # Columns to pivot
    aggfunc='first'  # Aggregation function to apply
).reset_index()  # Reset index to flatten the pivot table

# Flatten the MultiIndex columns
df_dashboard_pivot.columns = ['_'.join(map(str, col)).strip() if isinstance(col, tuple) else col for col in df_dashboard_pivot.columns]


# # Create 'matched_relocation' column with TRUE if they match, else FALSE
# df_dashboard_pivot['matched relocation in main site'] = df_dashboard_pivot.apply(lambda row: False if row['Current Tech Id_'] == row['Relocation OLD SiteId_'] and row['Relocation OLD SiteId_'] != '#N/A' else np.nan, axis=1)


print(f"Renaming Columns for Dashboard preparation")
#Rename columns to match the required format
df_dashboard_pivot.rename(columns={
    'Unique Current Tech ID_': 'Unique Current Tech ID', 'Circle_': 'Circle', 'Current Tech Id_': 'Current Tech Id', 'A/L_' :'A/L', 'Relocation OLD SiteId_' :'Relocation OLD SiteId', 'AMC_2G' : '2G', 'AMC_L1800' : 'L1800', 'AMC_L900' : 'L900', 'AMC_L2300' : 'L2300', 'AMC_L2100' :'L2100', 'AMC_L850' : 'L850'
}, inplace=True)

print(f"Removing time from Date Columns")

df_dashboard_pivot['Date On Air_2G'] = df_dashboard_pivot['Date On Air_2G'].dt.date
df_dashboard_pivot['Date On Air_L850'] = df_dashboard_pivot['Date On Air_L850'].dt.date
df_dashboard_pivot['Date On Air_L2100'] = df_dashboard_pivot['Date On Air_L2100'].dt.date
df_dashboard_pivot['Date On Air_L1800'] = df_dashboard_pivot['Date On Air_L1800'].dt.date
df_dashboard_pivot['Date On Air_L900'] = df_dashboard_pivot['Date On Air_L900'].dt.date
df_dashboard_pivot['Date On Air_L2300'] = df_dashboard_pivot['Date On Air_L2300'].dt.date
df_dashboard_pivot['Locked FD-AMC removal'] = ""
df_dashboard_pivot['Locked TD-Amc removal'] = ""

rearange_cols = ['Unique Current Tech ID', 'Circle', 'Current Tech Id', 'A/L', 'Relocation OLD SiteId', 'Locked FD-AMC removal', 'Locked TD-Amc removal' , 'Date On Air_2G','MS Status_2G','Date On Air_L850','MS Status_L850', 'Date On Air_L900', 'MS Status_L900', 'Date On Air_L1800', 'MS Status_L1800','Date On Air_L2100','MS Status_L2100', 'Date On Air_L2300', 'MS Status_L2300', '2G','L850', 'L900', 'L1800','L2100','L2300' ]



df_dashboard_pivot_1 = df_dashboard_pivot [rearange_cols]
pd.set_option('future.no_silent_downcasting', True)
df_dashboard_pivot_1.replace('#N/A', np.nan, inplace=True)

print(f"Setting Band Combination........")

def get_band_combination(row):
    bands = ['2G', 'L850', 'L900', 'L1800', 'L2100']
    active_bands = [band for band in bands if row[band] == 'Y']
    return '+'.join(active_bands)

# Apply the function to each row
df_dashboard_pivot_1['Band Combination'] = df_dashboard_pivot_1.apply(get_band_combination, axis=1)

print(f"Setting Final Combination........")
# Function to determine final combination
def get_final_combination(row):
    combination = []
    if row['2G'] == 'Y': combination.append('2G')
    if row['L850'] == 'Y': combination.append('4G')
    if row['L900'] == 'Y': combination.append('4G')
    if row['L1800'] == 'Y': combination.append('4G')
    if row['L2100'] == 'Y': combination.append('4G')
    
    # Combine the elements in the list into a single string
    return '+'.join(combination) if combination else ''

# Apply the function to each row
df_dashboard_pivot_1['Final Combination'] = df_dashboard_pivot_1.apply(get_final_combination, axis=1)

print(f"Setting TD Column ........")

def get_band_combination(row):
    bands = ['L2300']
    active_bands = [band for band in bands if row[band] == 'Y']
    return '+'.join(active_bands)

# Apply the function to each row
df_dashboard_pivot_1['TD'] = df_dashboard_pivot_1.apply(get_band_combination, axis=1)


#=========================================Dashboard 11==========================================================

print(f"Prepairing Dashboard without relocation site id.........")
# Applying Pivot Table
df_dashboard_11_pivot = pd.pivot_table(
    data=df_dashboard,  # Specify the DataFrame to pivot
    index=['Unique Current Tech ID', 'Circle','Current Tech Id', 'A/L'],  # Column(s) to use as the index
    columns=['Frequency Band'],  # Column(s) to use as the new columns
    values=['Date On Air', 'MS Status','AMC'],  # Columns to pivot
    aggfunc='first'  # Aggregation function to apply
).reset_index()  # Reset index to flatten the pivot table

# Flatten the MultiIndex columns
df_dashboard_11_pivot.columns = ['_'.join(map(str, col)).strip() if isinstance(col, tuple) else col for col in df_dashboard_11_pivot.columns]

print(f"Renaming Columns for Dashboard preparation")
#Rename columns to match the required format
df_dashboard_11_pivot.rename(columns={
    'Unique Current Tech ID_': 'Unique Current Tech ID', 'Circle_': 'Circle', 'Current Tech Id_': 'Current Tech Id', 'A/L_' :'A/L', 'AMC_2G' : '2G', 'AMC_L1800' : 'L1800', 'AMC_L900' : 'L900', 'AMC_L2300' : 'L2300', 'AMC_L2100' :'L2100', 'AMC_L850' : 'L850'
}, inplace=True)

print(f"Removing time from Date Columns")

df_dashboard_11_pivot['Date On Air_2G'] = df_dashboard_11_pivot['Date On Air_2G'].dt.date
df_dashboard_11_pivot['Date On Air_L850'] = df_dashboard_11_pivot['Date On Air_L850'].dt.date
df_dashboard_11_pivot['Date On Air_L2100'] = df_dashboard_11_pivot['Date On Air_L2100'].dt.date
df_dashboard_11_pivot['Date On Air_L1800'] = df_dashboard_11_pivot['Date On Air_L1800'].dt.date
df_dashboard_11_pivot['Date On Air_L900'] = df_dashboard_11_pivot['Date On Air_L900'].dt.date
df_dashboard_11_pivot['Date On Air_L2300'] = df_dashboard_11_pivot['Date On Air_L2300'].dt.date
df_dashboard_11_pivot['Locked FD-AMC removal'] = ""
df_dashboard_11_pivot['Locked TD-Amc removal'] = ""

rearange_cols = ['Unique Current Tech ID', 'Circle', 'Current Tech Id', 'A/L', 'Locked FD-AMC removal', 'Locked TD-Amc removal' , 'Date On Air_2G','MS Status_2G','Date On Air_L850','MS Status_L850', 'Date On Air_L900', 'MS Status_L900', 'Date On Air_L1800', 'MS Status_L1800','Date On Air_L2100','MS Status_L2100', 'Date On Air_L2300', 'MS Status_L2300', '2G','L850', 'L900', 'L1800','L2100','L2300' ]



df_dashboard_pivot_2 = df_dashboard_11_pivot [rearange_cols]
pd.set_option('future.no_silent_downcasting', True)
df_dashboard_pivot_2.replace('#N/A', np.nan, inplace=True)

print(f"Setting Band Combination........")

def get_band_combination(row):
    bands = ['2G', 'L850', 'L900', 'L1800', 'L2100']
    active_bands = [band for band in bands if row[band] == 'Y']
    return '+'.join(active_bands)

# Apply the function to each row
df_dashboard_pivot_2['Band Combination'] = df_dashboard_pivot_2.apply(get_band_combination, axis=1)

print(f"Setting Final Combination........")
# Function to determine final combination
def get_final_combination(row):
    combination = []
    if row['2G'] == 'Y': combination.append('2G')
    if row['L850'] == 'Y': combination.append('4G')
    if row['L900'] == 'Y': combination.append('4G')
    if row['L1800'] == 'Y': combination.append('4G')
    if row['L2100'] == 'Y': combination.append('4G')
    
    # Combine the elements in the list into a single string
    return '+'.join(combination) if combination else ''

# Apply the function to each row
df_dashboard_pivot_2['Final Combination'] = df_dashboard_pivot_2.apply(get_final_combination, axis=1)

print(f"Setting TD Column ........")

def get_band_combination(row):
    bands = ['L2300']
    active_bands = [band for band in bands if row[band] == 'Y']
    return '+'.join(active_bands)

# Apply the function to each row
df_dashboard_pivot_2['TD'] = df_dashboard_pivot_2.apply(get_band_combination, axis=1)

#=================================================Output File=======================================================
output_file_path = "D:\\Automation\\CBS_Auto\\Output\\CBS_Report.xlsx"
print(f"Printing Content to:", output_file_path)
# Set the maximum number of rows per sheet
max_rows_per_sheet = 1048570
# Save the filtered data and pivot table to Excel

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    # Save the dump DataFrame in chunks
    for i, start_row in enumerate(range(0, len(combined_df), max_rows_per_sheet)):
        end_row = start_row + max_rows_per_sheet
        chunk = combined_df.iloc[start_row:end_row]
        chunk.to_excel(writer, sheet_name=f"Dump_{i+1}", index=False)
    #df_relocation_sites.to_excel(writer, sheet_name="relocation sites", index=False)  
    df_dashboard_pivot_2.to_excel(writer, sheet_name="Dashboard", index=False) 
    df_dashboard_pivot_1.to_excel(writer, sheet_name="Sub_Dashboard", index=False)  
    #test_df.to_excel(writer, sheet_name="test df", index=False) 

   
print("Data saved successfully to multiple sheets in:", output_file_path)
