import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

# Initialize an empty list to store DataFrames
data = []

# Function to update checkboxes with unique columns from DataFrames
def update_checkboxes():
    all_columns = set()
    for df in data:
        all_columns.update(df.columns)

    global checkboxes
    checkboxes = {}
    
    # Clear existing checkboxes
    for widget in scrollable_frame.winfo_children():
        widget.destroy()

    # Split columns into 4 equal parts
    columns_list = sorted(list(all_columns))
    chunk_size = len(columns_list) // 4
    columns_parts = [
        columns_list[i:i + chunk_size] for i in range(0, len(columns_list), chunk_size)
    ]

    # Create checkboxes for each part
    for i, part in enumerate(columns_parts):
        frame_part = tk.Frame(scrollable_frame, width=250)
        frame_part.grid(row=0, column=i, padx=5, pady=5, sticky="n")
        for column in part:
            var = tk.BooleanVar()
            checkbox = tk.Checkbutton(frame_part, text=column, variable=var, font=("Arial", 10))
            checkbox.pack(anchor="w", pady=2)
            checkboxes[column] = var

# Function to choose a folder and process files
def choose_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    if not folder_path:
        messagebox.showerror("Error", "No folder selected!")
        return

    status_label.config(text="Processing files, please wait...", fg="blue")
    folder_button.config(state="disabled")
    data.clear()
    progress_text.delete(1.0, tk.END)

    thread = threading.Thread(target=process_files)
    thread.start()

# Function to process files in a separate thread
def process_files():
    try:
        for file in os.listdir(folder_path):
            if file.endswith(('.xlsx', '.xls')):
                file_path = os.path.join(folder_path, file)
                try:
                    progress_text.insert(tk.END, f"Processing file: {file}\n")
                    progress_text.yview(tk.END)

                    # Read sheets only if they exist
                    sheet_names = pd.ExcelFile(file_path).sheet_names
                    if 'CBS_A' in sheet_names:
                        df_cbs_a = pd.read_excel(file_path, sheet_name='CBS_A', header=1)
                        df_cbs_a['Source File'] = file
                        df_cbs_a['Sheet Name'] = 'CBS_A'
                        data.append(df_cbs_a)
                        progress_text.insert(tk.END, f"'CBS_A' processed. Size: {df_cbs_a.shape}\n")
                    else:
                        progress_text.insert(tk.END, f"'CBS_A' not found in {file}\n")

                    if 'CBS_L' in sheet_names:
                        df_cbs_l = pd.read_excel(file_path, sheet_name='CBS_L', header=1)
                        df_cbs_l['Source File'] = file
                        df_cbs_l['Sheet Name'] = 'CBS_L'
                        data.append(df_cbs_l)
                        progress_text.insert(tk.END, f"'CBS_L' processed. Size: {df_cbs_l.shape}\n")
                    else:
                        progress_text.insert(tk.END, f"'CBS_L' not found in {file}\n")

                except Exception as e:
                    progress_text.insert(tk.END, f"Error processing {file}: {e}\n")
                finally:
                    progress_text.yview(tk.END)

        # Filter out empty or all-NA columns
        data_filtered = [df.dropna(axis=1, how='all') for df in data]

        global combined_df_2g
        if data_filtered:
            combined_df_2g = pd.concat(data_filtered, ignore_index=True)
            progress_text.insert(tk.END, "Data combined successfully!\n")
        else:
            combined_df_2g = pd.DataFrame()
            progress_text.insert(tk.END, "No valid data found.\n")

        # Automatically update columns after data processing
        update_checkboxes()

        status_label.config(text="Files processed successfully!", fg="green")
        folder_button.config(state="normal")
        update_button.config(state="normal")
        download_button.config(state="normal")

    except Exception as e:
        progress_text.insert(tk.END, f"Error: {e}\n")
        status_label.config(text=f"Error: {e}", fg="red")
        folder_button.config(state="normal")

# Function to reset the application
def reset_data():
    progress_text.delete(1.0, tk.END)
    status_label.config(text="")
    for widget in scrollable_frame.winfo_children():
        widget.destroy()
    global combined_df_2g
    combined_df_2g = pd.DataFrame()
    folder_button.config(state="normal")
    update_button.config(state="disabled")
    download_button.config(state="disabled")

# Function to download filtered data
def download_data():
    selected_columns = [column for column, var in checkboxes.items() if var.get()]
    if not selected_columns:
        messagebox.showerror("Error", "No columns selected!")
        return

    filtered_data = combined_df_2g[selected_columns]
    max_rows_per_sheet = 1048570

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
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
root.title("Excel Data Processor")
root.minsize(700, 500)

# Add buttons in a top-centered frame
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

folder_button = tk.Button(button_frame, text="Choose Folder", command=choose_folder, font=("Arial", 12))
folder_button.pack(side="left", padx=10)

update_button = tk.Button(button_frame, text="Update Columns", command=update_checkboxes, font=("Arial", 12), state="disabled")
update_button.pack(side="left", padx=10)

download_button = tk.Button(button_frame, text="Download Data", command=download_data, font=("Arial", 12), state="disabled")
download_button.pack(side="left", padx=10)

reset_button = tk.Button(button_frame, text="Reset", command=reset_data, font=("Arial", 12))
reset_button.pack(side="left", padx=10)

status_label = tk.Label(root, text="", font=("Arial", 10), fg="black")
status_label.pack(pady=5)

# Create the columns window
frame = tk.Frame(root)
frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

canvas = tk.Canvas(frame)
scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Create progress window at the bottom
progress_text = tk.Text(root, width=80, height=15, wrap=tk.WORD, font=("Arial", 10))
progress_text.pack(pady=10)
progress_text.insert(tk.END, "Progress log will appear here...\n")

root.mainloop()
