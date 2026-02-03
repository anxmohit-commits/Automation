import pandas as pd
import numpy as np

# ============================================================Ericsson================================================================
df_dashboard_eri  = pd.DataFrame({
    "Circle": ["AP", "CHN", "DL", "HP", "HR", "J&K", "KL", "KK", "NESA","PB", "RAJ", "TN", "UPE", "UPW"],
    "OEM": ["Eric"] * 14,    
})

# Specify the folder path
folder_path = "D:\\Automation\\Spare Management\\Upload"

# Read Ericsson file and clean the data
file_path_ericsson = f"{folder_path}\\BHARTI HWS Summary_Ericsson.xlsb"

columns_to_load = ['Customer Name','Circle','Serial No.','Part Code','M5 SHIP DATE','Return Reason Meaning']

# Use pyxlsb engine to read .xlsb file
df_ericsson_spare = pd.read_excel(file_path_ericsson, sheet_name="HWS Data", usecols=columns_to_load, engine="pyxlsb")

# Convert 'M5 SHIP DATE' from serial to date format
if pd.api.types.is_numeric_dtype(df_ericsson_spare['M5 SHIP DATE']):
    df_ericsson_spare['M5 SHIP DATE'] = pd.to_timedelta(df_ericsson_spare['M5 SHIP DATE'], unit='D') + pd.Timestamp('1899-12-30')
else:
    df_ericsson_spare['M5 SHIP DATE'] = pd.NaT  # Handle non-numeric cases safely

# Drop rows with invalid dates
df_ericsson_spare = df_ericsson_spare.dropna(subset=['M5 SHIP DATE'])

# Extract 'Month-Year' format and convert to datetime
df_ericsson_spare['M5 SHIP Month-Year'] = df_ericsson_spare['M5 SHIP DATE'].dt.to_period('M').dt.to_timestamp()

# Drop rows where 'Customer Name' contains 'PROJECT'
df_ericsson_spare = df_ericsson_spare[~df_ericsson_spare['Customer Name'].str.contains('PROJECT', case=False, na=False)]

# Get the max Month-Year
max_month_year = df_ericsson_spare['M5 SHIP Month-Year'].max()

# Filter data for the max Month-Year
if pd.notna(max_month_year):
    df_ericsson_spare = df_ericsson_spare[df_ericsson_spare['M5 SHIP Month-Year'] == max_month_year]
    print("Ericsson: Auto-selected maximum Month-Year:", max_month_year.strftime('%b-%Y'))
else:
    print("No valid Month-Year found.")





# # Extract unique Month-Year values
# unique_month_years = df_ericsson_spare['M5 SHIP Month-Year'].unique()
# print("Unique M5 SHIP Month-Year values:")
# for month_year in unique_month_years:
#     print(month_year)

# # Ask user for desired Month-Year values
# desired_month_years = input("Enter desired Month-Year in Mmm-YYYY format for Ericsson: ").split(',')

# # Strip any extra whitespace from user input
# desired_month_years = [month_year.strip() for month_year in desired_month_years]


# Filter the DataFrame based on user input

# Remove "BHARTI " and " 5G" from the "Circle" column
df_ericsson_spare['Circle'] = df_ericsson_spare['Circle'].str.replace('BHARTI ', '', regex=False)
df_ericsson_spare['Circle'] = df_ericsson_spare['Circle'].str.replace(' 5G', '', regex=False)
df_ericsson_spare['Circle'] = df_ericsson_spare['Circle'].str.replace('DEL', 'DL', regex=False)
df_ericsson_spare['Circle'] = df_ericsson_spare['Circle'].str.replace('RJ', 'RAJ', regex=False)
df_ericsson_spare['Circle'] = df_ericsson_spare['Circle'].str.replace('JK', 'J&K', regex=False)
df_ericsson_spare['Circle'] = df_ericsson_spare['Circle'].str.replace('KER', 'KL', regex=False)
df_ericsson_spare['Circle'] = df_ericsson_spare['Circle'].str.replace('PUN', 'PB', regex=False)

# Update the Circle column using np.where
df_ericsson_spare['Circle'] = np.where(
    df_ericsson_spare['Customer Name'].str.contains('CHENNAI', case=False, na=False),
    'CHN',
    np.where(
        df_ericsson_spare['Customer Name'].str.contains('COIMBATORE', case=False, na=False),
        'TN',
        df_ericsson_spare['Circle']  # Keep the original value if no condition is met
    )
)
df_scrap_eri = df_ericsson_spare[
    df_ericsson_spare['Return Reason Meaning'].str.contains('Scrap|Scrap Adj', case=False, na=False)
]
# Group by the "Circle" column and count the occurrences
circle_counts_ericsson = df_ericsson_spare.groupby('Circle').size().reset_index(name='Cards Received')
circle_counts_scarp_ericsson = df_scrap_eri.groupby('Circle').size().reset_index(name='Scrap')

# Perform a left join on the 'Circle' column
df_dashboard_eri = pd.merge(df_dashboard_eri, circle_counts_ericsson, on='Circle', how='left')
df_dashboard_eri['Cards Received'] = df_dashboard_eri['Cards Received'].fillna(0)
df_dashboard_eri = pd.merge(df_dashboard_eri, circle_counts_scarp_ericsson, on='Circle', how='left')
df_dashboard_eri['Scrap'] = df_dashboard_eri['Scrap'].fillna(0)
df_dashboard_eri['%'] = (df_dashboard_eri['Scrap'] / df_dashboard_eri['Cards Received']) * 100
# Format the % column as whole numbers followed by a percent sign, handling NaN values
df_dashboard_eri['%'] = df_dashboard_eri['%'].apply(lambda x: f"{int(round(x))}%" if pd.notna(x) else '0%')
total_cards_received_eri = df_dashboard_eri['Cards Received'].sum()
total_scrap_eri = df_dashboard_eri['Scrap'].sum()
# Calculate the total percentage
total_percentage_ericsson = (total_scrap_eri / total_cards_received_eri) * 100 if total_cards_received_eri != 0 else 0


# Append the total row
total_row_eri = pd.DataFrame({
    'Circle': ['Total'],
    'OEM': [''],
    'Cards Received': [total_cards_received_eri],
    'Scrap': [total_scrap_eri],
    '%': [f"{total_percentage_ericsson:.2f}%"]
})

df_dashboard_eri = pd.concat([df_dashboard_eri, total_row_eri], ignore_index=True)


#=================================================================Nokia==================================================================================================
df_dashboard_nokia  = pd.DataFrame({
    "Circle": ["BH", "DL", "CHN", "GUJ", "J&K", "KL", "KOL","MHG","MPCG","MUM","OR","RAJ","PB","UPE","UPW","WB"],
    "OEM": ["Nokia"] * 16,    
})



# Specify the folder path, using either double backslashes or a raw string
folder_path = "D:\\Automation\\Spare Management\\Upload"
# Read Ericsson file and clean the data
file_path_nokia = f"{folder_path}\\Bharti Monthly Report_Nokia.xlsb"

columns_to_load = ['Circle','Dispatch date Physical','Remarks']
# Use pyxlsb engine to read .xlsb file
df_nokia_spare = pd.read_excel(file_path_nokia, sheet_name="Despatch",usecols=columns_to_load, engine="pyxlsb",header = 1)
# Convert 'M5 SHIP DATE' from serial to date format
# df_nokia_spare['Dispatch date Physical'] = pd.TimedeltaIndex(df_nokia_spare['Dispatch date Physical'], unit='D') + pd.Timestamp('1899-12-30')
df_nokia_spare['Dispatch date Physical'] = pd.to_timedelta(df_nokia_spare['Dispatch date Physical'], unit='D') + pd.Timestamp('1899-12-30')
df_nokia_spare = df_nokia_spare.dropna(subset=['Dispatch date Physical'])  
df_nokia_spare['Dispatch Month-Year Physical'] = df_nokia_spare['Dispatch date Physical'].dt.strftime('%b-%Y')


df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Bihar', 'BH', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Maharashtra', 'MHG', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Mumbai', 'MUM', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('West Bengal', 'WB', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Kerala', 'KL', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Gujrat', 'GUJ', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Odisha', 'OR', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('UP East', 'UPE', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Madhya Pradesh', 'MPCG', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Punjab', 'PB', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Chennai', 'CHN', regex=False)
df_nokia_spare['Circle'] = df_nokia_spare['Circle'].str.replace('Delhi', 'DL', regex=False)


df_scrap_nokia = df_nokia_spare[
    df_nokia_spare['Remarks'].str.contains('RNP', case=False, na=False)
]
# Group by the "Circle" column and count the occurrences
circle_counts_nokia = df_nokia_spare.groupby('Circle').size().reset_index(name='Cards Received')
circle_counts_scarp_nokia = df_scrap_nokia.groupby('Circle').size().reset_index(name='Scrap')

# Perform a left join on the 'Circle' column
df_dashboard_nokia = pd.merge(df_dashboard_nokia, circle_counts_nokia, on='Circle', how='left')
df_dashboard_nokia['Cards Received'] = df_dashboard_nokia['Cards Received'].fillna(0)
df_dashboard_nokia = pd.merge(df_dashboard_nokia, circle_counts_scarp_nokia, on='Circle', how='left')
df_dashboard_nokia['Scrap'] = df_dashboard_nokia['Scrap'].fillna(0)
df_dashboard_nokia['%'] = (df_dashboard_nokia['Scrap'] / df_dashboard_nokia['Cards Received']) * 100
# Format the % column as whole numbers followed by a percent sign, handling NaN values
df_dashboard_nokia['%'] = df_dashboard_nokia['%'].apply(lambda x: f"{int(round(x))}%" if pd.notna(x) else '0%')
total_cards_received_nokia = df_dashboard_nokia['Cards Received'].sum()
total_scrap_nokia = df_dashboard_nokia['Scrap'].sum()
# Calculate the total percentage
total_percentage_nokia = (total_scrap_nokia / total_cards_received_nokia) * 100 if total_cards_received_nokia != 0 else 0


# Append the total row
total_row_nokia = pd.DataFrame({
    'Circle': ['Total'],
    'OEM': [''],
    'Cards Received': [total_cards_received_nokia],
    'Scrap': [total_scrap_nokia],
    '%': [f"{total_percentage_nokia:.2f}%"]
})

df_dashboard_nokia = pd.concat([df_dashboard_nokia, total_row_nokia], ignore_index=True)


#============================================================================Ceragon===============================================================================
df_dashboard_ceragon  = pd.DataFrame({
    "Circle": ["NESA" , "BH" , "MPCG" , "JH" , "KL" , "KOL" , "MHG" , "MUM" , "OR" , "PB" , "RAJ" , "TN" , "AP" , "UPE" , "UPW" , "WB" , "GUJ"],
    "OEM": ["Ceragon"] * 17,    
})



# Specify the folder path, using either double backslashes or a raw string
folder_path = "D:\\Automation\\Spare Management\\Upload"
# Read Ericsson file and clean the data
file_path_ceragon = f"{folder_path}\\Bharti RR Report Latest_Ceragon.xlsx"

columns_to_load = ['Shipment Date','Repair Status','Circle','Category','Month']
# Use pyxlsb engine to read .xlsb file
df_ceragon_spare = pd.read_excel(file_path_ceragon, sheet_name="OUT",usecols=columns_to_load, )
# Convert 'M5 SHIP DATE' from serial to date format

# df_ceragon_spare['Shipment Date'] = pd.to_timedelta(df_ceragon_spare['Shipment Date'], unit='D') + pd.Timestamp('1899-12-30')
# df_ceragon_spare = df_ceragon_spare.dropna(subset=['Shipment Date'])  
# df_ceragon_spare['Shipment Month-Year'] = df_ceragon_spare['Shipment Date'].dt.strftime('%b-%Y')

# # Drop rows where 'Customer Name' contains the keyword 'PROJECT'
df_ceragon_spare = df_ceragon_spare[~df_ceragon_spare['Category'].str.contains('Project', case=False, na=False)]

# Extract unique Month values
unique_months = df_ceragon_spare['Month'].unique()
print("Unique Dispatch Month values for Ceragon:")
for month in unique_months:
    print(month)


# Ask user for desired Month-Year values
desired_month = input("Enter desired Month(s) in 'Mmm' format (comma-separated) for Ceragon: ")

# Strip whitespace, handle case sensitivity, and convert to a list
desired_months = [month.strip().capitalize() for month in desired_month.split(",")]

# Filter the DataFrame based on user input
df_ceragon_spare = df_ceragon_spare[df_ceragon_spare['Month'].isin(desired_months)]

circle_mapping = {
    'Bihar': 'BH',
    'Maharashtra': 'MHG',
    'M&G': 'MHG',
    'UP-West': 'UPW',
    'Kolkata': 'KOL',
    'kolkata': 'KOL',
    'Mumbai': 'MUM',
    'Rajasthan': 'RAJ',
    'Jharkhand': 'JH',
    'Tamil Nadu': 'TN',
    'Chhattisgarh': 'MPCG',
    'West Bengal': 'WB',
    'Kerala': 'KL',
    'Gujarat': 'GUJ',
    'Odisha': 'OR',
    'UP-East': 'UPE',
    'Punjab': 'PB',
    'Madhya Pradesh': 'MPCG',
    'Chennai': 'CHN'
}

df_ceragon_spare['Circle'] = df_ceragon_spare['Circle'].replace(circle_mapping)

#
df_scrap_ceragon = df_ceragon_spare[
    df_ceragon_spare['Repair Status'].str.contains(r'RNP|Scrap', case=False, na=False)
]
# Group by the "Circle" column and count the occurrences
circle_counts_ceragon = df_ceragon_spare.groupby('Circle').size().reset_index(name='Cards Received')
circle_counts_scarp_ceragon = df_scrap_ceragon.groupby('Circle').size().reset_index(name='Scrap')

# Perform a left join on the 'Circle' column
df_dashboard_ceragon = pd.merge(df_dashboard_ceragon, circle_counts_ceragon, on='Circle', how='left')
df_dashboard_ceragon['Cards Received'] = df_dashboard_ceragon['Cards Received'].fillna(0)
df_dashboard_ceragon = pd.merge(df_dashboard_ceragon, circle_counts_scarp_ceragon, on='Circle', how='left')
df_dashboard_ceragon['Scrap'] = df_dashboard_ceragon['Scrap'].fillna(0)
df_dashboard_ceragon['%'] = (df_dashboard_ceragon['Scrap'] / df_dashboard_ceragon['Cards Received']) * 100
# Format the % column as whole numbers followed by a percent sign, handling NaN values
df_dashboard_ceragon['%'] = df_dashboard_ceragon['%'].apply(lambda x: f"{int(round(x))}%" if pd.notna(x) else '0%')
total_cards_received_ceragon = df_dashboard_ceragon['Cards Received'].sum()
total_scrap_ceragon = df_dashboard_ceragon['Scrap'].sum()
# Calculate the total percentage
total_percentage_ceragon = (total_scrap_ceragon / total_cards_received_ceragon) * 100 if total_cards_received_ceragon != 0 else 0


# Append the total row
total_row_ceragon = pd.DataFrame({
    'Circle': ['Total'],
    'OEM': [''],
    'Cards Received': [total_cards_received_ceragon],
    'Scrap': [total_scrap_ceragon],
    '%': [f"{total_percentage_ceragon:.2f}%"]
})

df_dashboard_ceragon = pd.concat([df_dashboard_ceragon, total_row_ceragon], ignore_index=True)

#===========================================Huawei=============================================

df_dashboard_huawei  = pd.DataFrame({
    "Circle": ["GUJ", "J&K", "KL", "KK","MPCG","PB","RAJ","TN","CHN","UPW"],
    "OEM": ["Huawei"] * 10,    
})



# Specify the folder path, using either double backslashes or a raw string
folder_path = "D:\\Automation\\Spare Management\\Upload"
# Read Ericsson file and clean the data
file_path_huawei = f"{folder_path}\\Bharti- R&R Tracker_Huawei.xlsb"

columns_to_load = ['Circle','Good Parts Ship Date','RMA Status']
# Use pyxlsb engine to read .xlsb file
df_huawei_spare_md = pd.read_excel(file_path_huawei, sheet_name="Microwave Data", usecols=columns_to_load, engine="pyxlsb")
df_huawei_spare_wd = pd.read_excel(file_path_huawei, sheet_name="Wireless Data", usecols=columns_to_load, engine="pyxlsb")
df_huawei_spare = pd.concat([df_huawei_spare_md,df_huawei_spare_wd])

# Convert 'M5 SHIP DATE' from serial to date format
# df_nokia_spare['Dispatch date Physical'] = pd.TimedeltaIndex(df_nokia_spare['Dispatch date Physical'], unit='D') + pd.Timestamp('1899-12-30')
df_huawei_spare['Good Parts Ship Date'] = pd.to_timedelta(df_huawei_spare['Good Parts Ship Date'], unit='D') + pd.Timestamp('1899-12-30')

# df_huawei_spare = df_huawei_spare.dropna(subset="Circle")
# df_huawei_spare = df_huawei_spare.dropna(subset=['Good Parts Ship Date'])  
# df_huawei_spare = df_huawei_spare.dropna(subset="RMA Status")
df_huawei_spare['Good Parts Ship Month-Year'] = pd.to_datetime(
    df_huawei_spare['Good Parts Ship Date'], format='%b-%Y', errors='coerce'
)


# # Extract unique Month-Year values
# unique_month_years = df_huawei_spare['Good Parts Ship Month-Year'].unique()
# print("Unique Dispatch Month-Year values for Huawei:")
# for month_year in unique_month_years:
#     print(month_year)




# Get the max Month-Year
max_month_year = df_huawei_spare['Good Parts Ship Month-Year'].max()

# Filter data for the max Month-Year
if pd.notna(max_month_year):
    df_huawei_spare = df_huawei_spare[df_huawei_spare['Good Parts Ship Month-Year'] == max_month_year]
    print("Huawei: Auto-selected maximum Month-Year:", max_month_year.strftime('%b-%Y'))
else:
    print("No valid Month-Year found.")





# # Ask user for desired Month-Year values
# desired_month_years = input("Enter desired Month-Year in Mmm-YYYY format for Huawei: ").split(',')

# # Strip any extra whitespace from user input
# desired_month_years = [month_year.strip() for month_year in desired_month_years]

# # Filter the DataFrame based on user input
# df_huawei_spare = df_huawei_spare[df_huawei_spare['Good Parts Ship Month-Year'].isin(desired_month_years)]

circle_mapping = {
    'KTK-(Mobility)': 'KK', 
    'UP-West': 'UPW',
    'Rajasthan-Mobility': 'RAJ',
    'Rajasthan': 'RAJ',
    'Tamilnadu': 'TN',
    'TN-(Mobility)': 'TN',
    'CG-(Mobility)': 'MPCG',
    'MP-(Mobility)': 'MPCG',
    'Kerala': 'KL',
    'Gujarat-(Mobility)': 'GUJ',
    'Punjab': 'PB',    
    'Chennai': 'CHN',
    'UPW-(Mobility)': 'UPW',
    'J&K': 'J&K',
    'Chennai-(Mobility)': 'CHN'
}
df_huawei_spare['Circle'] = df_huawei_spare['Circle'].replace(circle_mapping)

df_scrap_huawei = df_huawei_spare[
    df_huawei_spare['RMA Status'].str.contains('Discard|Physical', case=False, na=False)
]
# Group by the "Circle" column and count the occurrences
circle_counts_huawei = df_huawei_spare.groupby('Circle').size().reset_index(name='Cards Received')
circle_counts_scarp_huawei = df_scrap_huawei.groupby('Circle').size().reset_index(name='Scrap')

# Perform a left join on the 'Circle' column
df_dashboard_huawei = pd.merge(df_dashboard_huawei, circle_counts_huawei, on='Circle', how='left')
df_dashboard_huawei['Cards Received'] = df_dashboard_huawei['Cards Received'].fillna(0)
df_dashboard_huawei = pd.merge(df_dashboard_huawei, circle_counts_scarp_huawei, on='Circle', how='left')
df_dashboard_huawei['Scrap'] = df_dashboard_huawei['Scrap'].fillna(0)
df_dashboard_huawei['%'] = (df_dashboard_huawei['Scrap'] / df_dashboard_huawei['Cards Received']) * 100
# Format the % column as whole numbers followed by a percent sign, handling NaN values
df_dashboard_huawei['%'] = df_dashboard_huawei['%'].apply(lambda x: f"{int(round(x))}%" if pd.notna(x) else '0%')
total_cards_received_huawei = df_dashboard_huawei['Cards Received'].sum()
total_scrap_huawei = df_dashboard_huawei['Scrap'].sum()
# Calculate the total percentage
total_percentage_huawei = (total_scrap_huawei / total_cards_received_huawei) * 100 if total_cards_received_huawei != 0 else 0


# Append the total row
total_row_huawei = pd.DataFrame({
    'Circle': ['Total'],
    'OEM': [''],

    'Cards Received': [total_cards_received_huawei],
    'Scrap': [total_scrap_huawei],
    '%': [f"{total_percentage_huawei:.2f}%"]
})

df_dashboard_huawei = pd.concat([df_dashboard_huawei, total_row_huawei], ignore_index=True)

#==================================================================================ZTE==================================================================================

df_dashboard_zte  = pd.DataFrame({
    "Circle": ["HR","PB","WB"],
    "OEM": ["ZTE"] * 3,    
})



# Specify the folder path, using either double backslashes or a raw string
folder_path = "D:\\Automation\\Spare Management\\Upload"
# Read Ericsson file and clean the data
file_path_zte = f"{folder_path}\\Airtel Monthly R&R Report_ZTE.xlsx"

columns_to_load = ['Circle Name','Dispatch Date','Repair Status','Dispatch Contact Person']
# Use pyxlsb engine to read .xlsb file
df_zte_spare = pd.read_excel(file_path_zte, sheet_name="Monthly Dispatch", usecols=columns_to_load)

df_zte_spare = df_zte_spare[
    df_zte_spare['Dispatch Contact Person'].str.contains('9831856351|9368480224|9899107189|9817362253', case=False, na=False)
]

# Convert 'Dispatch Date' to datetime format
df_zte_spare['Dispatch Date'] = pd.to_datetime(df_zte_spare['Dispatch Date'], errors='coerce')


# Drop rows with NaN values in required columns
df_zte_spare = df_zte_spare.dropna(subset=['Circle Name', 'Dispatch Date', 'Repair Status'])

# Format 'Dispatch Date' to 'Dispatch Month-Year'
#df_zte_spare['Dispatch Month-Year'] = df_zte_spare['Dispatch Date'].dt.strftime('%b-%Y')


df_zte_spare['Dispatch Month-Year'] = pd.to_datetime(
    df_zte_spare['Dispatch Date'], format='%b-%Y', errors='coerce'
)


# # Extract unique Month-Year values
# unique_month_years = df_zte_spare['Dispatch Month-Year'].unique()
# print("Unique Dispatch Month-Year values for ZTE:")
# for month_year in unique_month_years:
#     print(month_year)



# Get the max Month-Year
max_month_year = df_zte_spare['Dispatch Month-Year'].max()

# Filter data for the max Month-Year
if pd.notna(max_month_year):
    df_zte_spare = df_zte_spare[df_zte_spare['Dispatch Month-Year'] == max_month_year]
    print("ZTE: Auto-selected maximum Month-Year:", max_month_year.strftime('%b-%Y'))
else:
    print("No valid Month-Year found.")




# # Ask user for desired Month-Year values
# desired_month_years = input("Enter desired Month-Year in Mmm-YYYY format for ZTE: ").split(',')

# # Strip any extra whitespace from user input
# desired_month_years = [month_year.strip() for month_year in desired_month_years]

# # Filter the DataFrame based on user input
# df_zte_spare = df_zte_spare[df_zte_spare['Dispatch Month-Year'].isin(desired_month_years)]



circle_mapping = {
    'West Bengal (WB)': 'WB', 
    'Punjab (PB)': 'PB',
    'Haryana (HR)': 'HR'
}

df_zte_spare['Circle'] = df_zte_spare['Circle Name'].replace(circle_mapping)


df_scrap_zte = df_zte_spare[
    df_zte_spare['Repair Status'].str.contains('Discard', case=False, na=False)
]
# Group by the "Circle" column and count the occurrences
circle_counts_zte = df_zte_spare.groupby('Circle').size().reset_index(name='Cards Received')
circle_counts_scarp_zte = df_scrap_zte.groupby('Circle').size().reset_index(name='Scrap')

# Perform a left join on the 'Circle' column
df_dashboard_zte = pd.merge(df_dashboard_zte, circle_counts_zte, on='Circle', how='left')
df_dashboard_zte['Cards Received'] = df_dashboard_zte['Cards Received'].fillna(0)
df_dashboard_zte = pd.merge(df_dashboard_zte, circle_counts_scarp_zte, on='Circle', how='left')
df_dashboard_zte['Scrap'] = df_dashboard_zte['Scrap'].fillna(0)
df_dashboard_zte['%'] = (df_dashboard_zte['Scrap'] / df_dashboard_zte['Cards Received']) * 100
# Format the % column as whole numbers followed by a percent sign, handling NaN values
df_dashboard_zte['%'] = df_dashboard_zte['%'].apply(lambda x: f"{int(round(x))}%" if pd.notna(x) else '0%')
total_cards_received_zte = df_dashboard_zte['Cards Received'].sum()
total_scrap_zte = df_dashboard_zte['Scrap'].sum()
# Calculate the total percentage
total_percentage_zte = (total_scrap_zte / total_cards_received_zte) * 100 if total_cards_received_zte != 0 else 0


# Append the total row
total_row_zte = pd.DataFrame({
    'Circle': ['Total'],
    'OEM': [''],

    'Cards Received': [total_cards_received_zte],
    'Scrap': [total_scrap_zte],
    '%': [f"{total_percentage_zte:.2f}%"]
})

df_dashboard_zte = pd.concat([df_dashboard_zte, total_row_zte], ignore_index=True)

#===============================================================================Dashboard==============================================================================

output_file_path = "D:\\Automation\\Spare Management\\Output\\Dashboard.xlsx"

def auto_adjust_column_widths(df, worksheet, startcol=0):
    for i, col in enumerate(df.columns):
        max_length = max(df[col].astype(str).map(len).max(), len(col)) + 2  # Find max length
        worksheet.set_column(startcol + i, startcol + i, max_length)  # Set column width

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
     
    #df_ericsson_spare.to_excel(writer, sheet_name="ericsson_Spare", index=False)
    df_dashboard_eri.to_excel(writer, sheet_name="SCRAP", index=False,startrow = 3, startcol = 1)
    df_dashboard_nokia.to_excel(writer, sheet_name="SCRAP", index=False, startrow = 3, startcol = 7)
    df_dashboard_ceragon.to_excel(writer, sheet_name="SCRAP", index=False, startrow = 3, startcol = 13)
    df_dashboard_huawei.to_excel(writer, sheet_name="SCRAP", index=False, startrow = 24, startcol = 1)
    df_dashboard_zte.to_excel(writer, sheet_name="SCRAP", index=False, startrow = 24, startcol = 7)

    workbook = writer.book
    worksheet = writer.sheets["SCRAP"]

    merge_format = workbook.add_format({'align': 'center', 'bold': True, 'font_size': 14})

    worksheet.merge_range(1, 1, 1, 5, "Ericsson Spare Summary", merge_format)
    worksheet.merge_range(1, 7, 1, 11, "Nokia Spare Summary", merge_format)
    worksheet.merge_range(1, 13, 1, 17, "Ceragon Spare Summary", merge_format)
    worksheet.merge_range(22, 1, 22, 5, "Huawei Spare Summary", merge_format)
    worksheet.merge_range(22, 7, 22, 11, "ZTE Spare Summary", merge_format)

    # Auto-adjust column widths for each DataFrame
    auto_adjust_column_widths(df_dashboard_eri, worksheet, startcol=1)
    auto_adjust_column_widths(df_dashboard_nokia, worksheet, startcol=7)
    auto_adjust_column_widths(df_dashboard_ceragon, worksheet, startcol=13)
    auto_adjust_column_widths(df_dashboard_huawei, worksheet, startcol=1)
    auto_adjust_column_widths(df_dashboard_zte, worksheet, startcol=7)
 
print("Dashboard saved successfully to :" + output_file_path)