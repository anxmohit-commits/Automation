import os
import pandas as pd

# Define the folder path
folder_path = r"D:\Automation\Regular Scripts\Upload"

# Get a list of all CSV files in the folder
csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]

# Dictionary to store DataFrames
dataframes = {}

# Read each CSV file and store it in the dictionary
for file in csv_files:
    file_path = os.path.join(folder_path, file)
    try:
        df = pd.read_csv(file_path)
        dataframes[file] = df
        print(f"Successfully read: {file}")
    except Exception as e:
        print(f"Error reading {file}: {e}")


df_town = pd.read_excel("D://Automation//Regular Scripts//input//Report.xlsx")
# Concatenate all DataFrames
if dataframes:
    final_df = pd.concat(dataframes, ignore_index=False)

    # **Fix column name spacing**
    final_df.columns = final_df.columns.str.strip()
  
    # Convert createdate to datetime
    final_df['createdate'] = pd.to_datetime(final_df['createdate'], errors='coerce')
    final_df = final_df.dropna(subset=['createdate'])  # Drop NaT values

    # Extract Month and Week Number
    final_df['Month'] = final_df['createdate'].dt.strftime('%b')
    final_df['Week_Number'] = final_df['createdate'].dt.isocalendar().week

    # Fill missing values
    final_df["city_name"].fillna("Other", inplace=True)
    final_df["rsu"].fillna("#N/A", inplace=True)
    final_df["circle"].fillna("#N/A", inplace=True)

    # Create new columns
    final_df["city_rsu"] = final_df["city_name"] + "_" + final_df["rsu"]
    final_df["circle_rsu"] = final_df["circle"] + "_" + final_df["rsu"]

    # Merge with df_town
    #df_town["circle_rsu"] = df_town["circle"] + "_" + df_town["rsu"]
    final_df = pd.merge(final_df, df_town, on='circle_rsu', how='left')

    # Convert sr_num to numeric
    final_df['sr_num'] = final_df['sr_num'].astype(str)

    print(final_df.info())
    print(final_df.head())

    print(f"\n✅ Successfully processed {len(dataframes)} files.")
    print(f"📉 Removed Duplicates: Final DataFrame Shape = {final_df.shape}")

else:
    print("\n❌ No data extracted, check file structure.")







# Create Pivot Table
monthwise = final_df.pivot_table(
index=['Month'], 
columns='sourcegroup', 
values='sr_num',  # Use any column that exists for counting
aggfunc= 'count',
fill_value= 0,  # Fill missing values with 0
margins=True,  # Add totals for both rows
margins_name='Total'  # Name for the row and column totals
).reset_index()

# Create Pivot Table
monthwise_circle = final_df.pivot_table(
index=['Month',"circle"], 
columns='sourcegroup', 
values='sr_num',  # Use any column that exists for counting
aggfunc= 'count',
fill_value= 0,  # Fill missing values with 0
margins=True,  # Add totals for both rows
margins_name='Total'  # Name for the row and column totals
).reset_index()

# Create Pivot Table
monthwise_circle_rsu = final_df.pivot_table(
index=['Month',"circle","city_rsu"], 
columns='sourcegroup', 
values='sr_num',  # Use any column that exists for counting
aggfunc= 'count',
fill_value= 0,  # Fill missing values with 0
margins=True,  # Add totals for both rows
margins_name='Total'  # Name for the row and column totals
).reset_index()





# Create Pivot Table
weekwise = final_df.pivot_table(
index=['Week_Number'], 
columns='sourcegroup', 
values='sr_num',  # Use any column that exists for counting
aggfunc= 'count',
fill_value= 0,  # Fill missing values with 0
margins=True,  # Add totals for both rows
margins_name='Total'  # Name for the row and column totals
).reset_index()

# Create Pivot Table
weekwise_circle = final_df.pivot_table(
index=['Week_Number',"circle"], 
columns='sourcegroup', 
values='sr_num',  # Use any column that exists for counting
aggfunc= 'count',
fill_value= 0,  # Fill missing values with 0
margins=True,  # Add totals for both rows
margins_name='Total'  # Name for the row and column totals
).reset_index()



# Create Pivot Table
weekwise_circle_rsu = final_df.pivot_table(
index=['Week_Number',"circle","city_rsu"], 
columns='sourcegroup', 
values='sr_num',  # Use any column that exists for counting
aggfunc= 'count',
fill_value= 0,  # Fill missing values with 0
margins=True,  # Add totals for both rows
margins_name='Total'  # Name for the row and column totals
).reset_index()




# Create Pivot Table
daywise = final_df.pivot_table(
index=['createdate'], 
columns='sourcegroup', 
values='sr_num',  # Use any column that exists for counting
aggfunc= 'count',
fill_value= 0,  # Fill missing values with 0
margins=True,  # Add totals for both rows
margins_name='Total'  # Name for the row and column totals
).reset_index()

# Create Pivot Table
daywise_circle = final_df.pivot_table(
index=['createdate',"circle"], 
columns='sourcegroup', 
values='sr_num',  # Use any column that exists for counting
aggfunc= 'count',
fill_value= 0,  # Fill missing values with 0
margins=True,  # Add totals for both rows
margins_name='Total'  # Name for the row and column totals
).reset_index()



# Create Pivot Table
daywise_circle_rsu = final_df.pivot_table(
index=['createdate',"circle","city_rsu"], 
columns='sourcegroup', 
values='sr_num',  # Use any column that exists for counting
aggfunc= 'count',
fill_value= 0,  # Fill missing values with 0
margins=True,  # Add totals for both rows
margins_name='Total'  # Name for the row and column totals
).reset_index()

max_rows_per_sheet = 1048570
# Save the filtered data and pivot table to Excel
output_folder = r"D:\Automation\Regular Scripts\output"
output_file_path = os.path.join(output_folder, "Report.xlsx")
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    # Save the `dump` DataFrame in chunks
    for i, start_row in enumerate(range(0, len(daywise), max_rows_per_sheet)):
        end_row = start_row + max_rows_per_sheet
        chunk = daywise.iloc[start_row:end_row]
        chunk.to_excel(writer, sheet_name=f"daywise_{i+1}", index=False)
    # Save the `dump` DataFrame in chunks
    for i, start_row in enumerate(range(0, len(daywise_circle), max_rows_per_sheet)):
        end_row = start_row + max_rows_per_sheet
        chunk = daywise_circle.iloc[start_row:end_row]
        chunk.to_excel(writer, sheet_name=f"daywise_circle{i+1}", index=False)
    # Save the `dump` DataFrame in chunks
    for i, start_row in enumerate(range(0, len(daywise_circle_rsu), max_rows_per_sheet)):
        end_row = start_row + max_rows_per_sheet
        chunk = daywise_circle_rsu.iloc[start_row:end_row]
        chunk.to_excel(writer, sheet_name=f"daywise_circle_rsu{i+1}", index=False)


    weekwise.to_excel(writer, sheet_name="weekwise", index=False)
    #Sitewise_backup_1.to_excel(writer, sheet_name="Sitewise_Backup_Daily_1", index=False)
    weekwise_circle.to_excel(writer, sheet_name="weekwise_circle", index=False) 
    #latest2daysdump.to_excel(writer, sheet_name="Dump_Latest_2_Days", index=False) 
    weekwise_circle_rsu.to_excel(writer, sheet_name="weekwise_circle_rsu", index=False)


    #final_df.to_excel(writer, sheet_name="dump", index=False)
    monthwise.to_excel(writer, sheet_name="monthwise", index=False)
    #Sitewise_backup_1.to_excel(writer, sheet_name="Sitewise_Backup_Daily_1", index=False)
    monthwise_circle.to_excel(writer, sheet_name="monthwise_circle", index=False) 
    #latest2daysdump.to_excel(writer, sheet_name="Dump_Latest_2_Days", index=False) 
    monthwise_circle_rsu.to_excel(writer, sheet_name="monthwise_circle_rsu", index=False)

    