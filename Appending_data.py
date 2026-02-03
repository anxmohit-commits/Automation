import pandas as pd
import numpy as np
import glob
import os
import openpyxl
import xlsxwriter
from datetime import datetime
from datetime import timedelta
#==============================================Reading 2G CBS Files=================================================
# Define the folder containing the Excel files
folder_path = 'D:\\Automation\\CBS_Auto\\Dump\\Current\\2G'  # Replace with your folder path


# List of columns to extract
required_columns = [
    'Current Tech Id.', 
    'Cell Name', 
    'Physical Site Id', 
    'Towns', 
    'Date On Air', 
    'MS1-Date', 
    'MS2-Date', 
    'Active /Locked/Dismantled', 
    'Date Since Locked', 
    'Locked SubType', 
    'Detailed Remarks against Locked', 
    'Date on Dismantled', 
    'Site TYPE (IM-Indoor Macro,OM-Outdoor Macro,TT-Tower Top Macro,MI-Micro,IB-IBS,COW,LM-Lamp Cell,MM-Mimo Cell,Q Cell,PICO,Small Cell)', 
    'Existing/New/Relocation', 
    'Relocation OLD SiteId', 
    'MS Status', 
    'Frequency BAND(900,1800,900+1800)'
]

# Initialize an empty list to store DataFrames
dataframes = []

# Loop through all files in the folder
for file in os.listdir(folder_path):
    # Check if the file is an Excel file
    if file.endswith('.xlsx') or file.endswith('.xls'):
        file_path = os.path.join(folder_path, file)
        
        try:
            
            # Read the required sheets into DataFrames
            df_cbs_a = pd.read_excel(file_path, sheet_name='CBS_A', header = 1)
            df_cbs_l = pd.read_excel(file_path, sheet_name='CBS_L', header = 1)
            
            # Add a source column to identify data origin
            df_cbs_a['Source File'] = file
            df_cbs_a['Sheet Name'] = 'CBS_A'
            
            df_cbs_l['Source File'] = file
            df_cbs_l['Sheet Name'] = 'CBS_L' 

            # Extract only the required columns from both DataFrames
            df_cbs_a_filtered = df_cbs_a[required_columns + ['Source File', 'Sheet Name']]
            df_cbs_l_filtered = df_cbs_l[required_columns + ['Source File', 'Sheet Name']]

            # Append the filtered DataFrames to the list
            dataframes.append(df_cbs_a_filtered)
            dataframes.append(df_cbs_l_filtered)
        
        except Exception as e:
            print(f"Error reading {file}: {e}")
            
# Filter out empty or all-NA columns from each DataFrame before concatenation
dataframes = [df.dropna(axis=1, how='all') for df in dataframes]

# Combine all DataFrames into one
if dataframes:
    combined_df_2g = pd.concat(dataframes, ignore_index=True)
    print("2G Data combined successfully!")
else:
    combined_df_2g = pd.DataFrame()
    print("No valid 2G data found.")