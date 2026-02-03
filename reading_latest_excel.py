import os
import pandas as pd

# Define the folder path
folder_path = r"D:\Automation\WO Dump\Dashboard"

# Get a list of all Excel files in the folder
excel_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) 
               if f.endswith(('.xls', '.xlsx')) and os.path.isfile(os.path.join(folder_path, f))]

# Find the latest modified Excel file
latest_file = max(excel_files, key=os.path.getmtime) if excel_files else None

df = pd.read_excel(latest_file, sheet_name="backup", header=0)  # Load 'backup' sheet
        
print(df)