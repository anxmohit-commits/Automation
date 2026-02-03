import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Define the folder containing the Excel files
folder_path = 'D:\\Automation\\CBS_Auto\\Sample files\\2G'  # Replace with your folder path

# Initialize an empty list to store DataFrames
dataframes_2g = []

# Loop through all files in the folder
for file in os.listdir(folder_path):
    if file.endswith('.xlsx') or file.endswith('.xls'):  # Check for Excel files
        file_path = os.path.join(folder_path, file)
        try:
            # Read the required sheets into DataFrames
            df_cbs_a = pd.read_excel(file_path, sheet_name='CBS_A', header=1)
            df_cbs_l = pd.read_excel(file_path, sheet_name='CBS_L', header=1)

            # Print the file name, sheet name, and size
            print(f"Reading File: {file}")
            print(f"Sheet Name: CBS_A, Size: {df_cbs_a.shape}")
            print(f"Sheet Name: CBS_L, Size: {df_cbs_l.shape}")

            # Add a source column to identify data origin
            df_cbs_a['Source File'] = file
            df_cbs_a['Sheet Name'] = 'CBS_A'
            df_cbs_l['Source File'] = file
            df_cbs_l['Sheet Name'] = 'CBS_L'

            # Append the filtered DataFrames to the list
            dataframes_2g.append(df_cbs_a)
            dataframes_2g.append(df_cbs_l)

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

# Function to handle downloading selected data, splitting into sheets if size exceeds the limit
def download_data():
    selected_columns = [column for column, var in checkboxes.items() if var.get()]
    if not selected_columns:
        messagebox.showerror("Error", "No columns selected!")
        return

    # Filter the DataFrame with selected columns
    filtered_data = combined_df_2g[selected_columns]

    # Define the maximum number of rows per sheet
    max_rows_per_sheet = 1048570  # Adjust this based on Excel limits

    # Ask the user for a save location
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Split data into multiple sheets if needed
                num_sheets = (len(filtered_data) // max_rows_per_sheet) + 1
                for i in range(num_sheets):
                    start_row = i * max_rows_per_sheet
                    end_row = min((i + 1) * max_rows_per_sheet, len(filtered_data))
                    sheet_name = f"Sheet{i + 1}"
                    filtered_data.iloc[start_row:end_row].to_excel(writer, index=False, sheet_name=sheet_name)

            messagebox.showinfo("Success", f"Data saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

# Create the main Tkinter window
root = tk.Tk()
root.title("Select Columns to Download")

# Add a label
label = tk.Label(root, text="Select Columns to Include in the Download", font=("Arial", 12))
label.pack(pady=10)

# Add a scrollable frame for checkboxes
frame = tk.Frame(root)
frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

canvas = tk.Canvas(frame)
scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Get all unique columns from the combined DataFrame
all_columns = set()
for df in dataframes_2g:
    all_columns.update(df.columns)

# Create checkboxes for each unique column
checkboxes = {}
for column in sorted(all_columns):  # Sorting for better readability
    var = tk.BooleanVar()
    checkbox = tk.Checkbutton(scrollable_frame, text=column, variable=var, font=("Arial", 10))
    checkbox.pack(anchor="w", pady=2)
    checkboxes[column] = var

# Add a button to download data
download_button = tk.Button(root, text="Download Selected Data", command=download_data, font=("Arial", 12))
download_button.pack(pady=20)

# Start the Tkinter main loop
root.mainloop()
