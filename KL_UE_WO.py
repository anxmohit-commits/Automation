import os
import pandas as pd

# Define the folder path and row limit
folder_path = r"D:\Automation\Regular Scripts\WO 6"
output_folder = r"D:\Automation\Regular Scripts\WO 6\Output"
row_limit = 1048570  # Set your desired row limit for each output file


columns = [

'WO Number', 
'WO Date/Time', 
'WO Type', 
'WO Status', 
'WO Assignee Name', 
'WO Assignee Mobile No', 
'Airtel Site ID', 
'Priority', 
'Circle', 
'SLA End Date', 
'TOCO Name', 
'TOCO ID', 
'Description', 
'Classic Notes', 
'NOC Reference ID - Airtel', 
'NOC Reference - Ericsson', 
'Ericsson Site ID', 
'Completed Within SLA', 
'Travel Time', 
'Zone', 
'op_categorization_tier1', 
'op_categorization_tier2', 
'op_categorization_tier3', 
'Site Visit Status', 
'WO Completion Date/time', 
'Time to resolve in hrs', 
'Airtel RFO 1', 
'Airtel RFO 2', 
'Airtel RFO 3'


]






# Define the list of Circles you want to filter
filter_circles = ["UE", "UPE", "KL", "Kerala"]

# Ensure the output folder exists
os.makedirs(output_folder, exist_ok=True)

# Create a DataFrame to store skipped rows
skipped_data = pd.DataFrame()

# Read and concatenate all CSV files
all_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
combined_data = pd.DataFrame()

for file in all_files:
    file_path = os.path.join(folder_path, file)
    print(f"Reading file: {file}")
    try:
        # Read CSV with error handling, skipping bad lines
        temp_df = pd.read_csv(file_path, usecols= columns, header=0)  # Skip rows with bad column counts
        
        # Filter the rows based on Circle column values
        temp_df_filtered = temp_df[temp_df['Circle'].isin(filter_circles)]
        combined_data = pd.concat([combined_data, temp_df_filtered], ignore_index=True)
    except Exception as e:
        print(f"Error reading {file}: {e}")
        # If an error occurs, save the skipped lines to a separate file (if available)
        skipped_data = skipped_data.append(temp_df, ignore_index=True)

# Save skipped lines to a CSV file
if not skipped_data.empty:
    skipped_lines_file = os.path.join(output_folder, "skipped_lines.csv")
    skipped_data.to_csv(skipped_lines_file, index=False)
    print(f"Skipped lines saved to: {skipped_lines_file}")

# Split the data if it exceeds the row limit
file_count = 1
for i in range(0, len(combined_data), row_limit):
    chunk = combined_data.iloc[i:i + row_limit]
    output_file = os.path.join(output_folder, f"output_part_{file_count}.csv")
    chunk.to_csv(output_file, index=False)
    print(f"Saved file: {output_file}")
    file_count += 1

print("Processing complete.")
