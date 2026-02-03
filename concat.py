import os
import pandas as pd
import numpy as np

# Define the folder path where Excel files are stored
folder_path = r"C:\Users\b0321755\Downloads\OneDrive_1_4-25-2025"

# Get a list of all .xlsx files in the folder
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# Dictionary to store DataFrames
dataframes = {}

# Loop through each file
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    try:
        df = pd.read_excel(file_path, engine='openpyxl')  # Read Excel file
        col_count = df.shape[1]
        print(f"✅ Read: {file} | Columns: {col_count}")

        # Check for Excel sheet column limit
        if col_count > 16384:
            print(f"⚠️ Skipped '{file}' - too many columns: {col_count}")
        else:
            dataframes[file] = df

    except Exception as e:
        print(f"❌ Error reading {file}: {e}")

# Combine all safe DataFrames
if dataframes:
    filtered_df = pd.concat(dataframes.values(), ignore_index=True)
    print(f"\n✅ Total combined rows: {filtered_df.shape[0]}, columns: {filtered_df.shape[1]}")

    # Save to Excel in chunks (limit 1,048,576 rows per sheet)
    output_folder = r"C:\Users\b0321755\Downloads"
    output_file_path = os.path.join(output_folder, "new_joinee.xlsx")
    max_rows_per_sheet = 1048570

    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        for i, start_row in enumerate(range(0, len(filtered_df), max_rows_per_sheet)):
            end_row = start_row + max_rows_per_sheet
            chunk = filtered_df.iloc[start_row:end_row]
            chunk.to_excel(writer, sheet_name=f"New_Joinee_{i+1}", index=False)

    print(f"\n✅ Export completed successfully: {output_file_path}")
else:
    print("⚠️ No valid files to combine or export.")
