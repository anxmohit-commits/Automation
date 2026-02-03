import os
import zipfile
import pandas as pd
from tqdm import tqdm

# Define paths
input_folder = r"D:\Automation\PM_Reports\Input"
output_folder = r"D:\Automation\PM_Reports\Output"
output_file = os.path.join(output_folder, "combined_output.xlsx")

def extract_zip_files(input_folder):
    extracted_files = []
    for file in os.listdir(input_folder):
        if file.endswith(".zip"):
            zip_path = os.path.join(input_folder, file)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                extracted_dir = os.path.join(input_folder, os.path.splitext(file)[0])
                zip_ref.extractall(extracted_dir)
                extracted_files.extend([(os.path.join(extracted_dir, f), file) for f in os.listdir(extracted_dir) if f.endswith(".xlsx") or f.endswith(".xls")])
    return extracted_files

def combine_excel_files(excel_files, output_file):
    combined_df = pd.DataFrame()
    for file, source_zip in tqdm(excel_files, desc="Reading Excel files", unit="file"):
        try:
            df = pd.read_excel(file, header=0)  # Change header if needed
            df['Source File'] = os.path.basename(file)  # Add a column to track extracted file
            df['Source ZIP'] = source_zip  # Add a column to track original ZIP file
            combined_df = pd.concat([combined_df, df], ignore_index=True)
        except Exception as e:
            print(f"Error reading {file}: {e}")
    combined_df.to_excel(output_file, index=False)
    print(f"Combined Excel file saved at: {output_file}")

# Run the process
extracted_files = extract_zip_files(input_folder)
if extracted_files:
    combine_excel_files(extracted_files, output_file)
else:
    print("No Excel files found in ZIP archives.")
