import pandas as pd
import os
import hashlib

def find_duplicate_excels(folder_path):
    hashes = {}  # Format: {hash_value: original_file_name}
    duplicates = []

    # Check agar path sahi hai
    if not os.path.exists(folder_path):
        print("Error: Diya gaya folder path sahi nahi hai.")
        return

    print(f"Scanning folder: {folder_path}\n")

    for filename in os.listdir(folder_path):
        # Sirf Excel files check karein
        if filename.endswith((".xlsx", ".xls", ".xlsm")):
            file_path = os.path.join(folder_path, filename)
            
            try:
                # File read karke content ka hash banana
                df = pd.read_excel(file_path)
                # Dataframe ko bytes mein convert karke hash nikalna
                file_hash = hashlib.md5(df.to_csv().encode('utf-8')).hexdigest()

                if file_hash in hashes:
                    duplicates.append((filename, hashes[file_hash]))
                else:
                    hashes[file_hash] = filename
            except Exception as e:
                print(f"File {filename} ko read nahi kar paye: {e}")

    # Results dikhana
    if duplicates:
        print("--- Duplicate Files Mil Gayi Hain ---")
        for dup, original in duplicates:
            print(f"DUPLICATE: {dup}  <== MATCHES ==>  ORIGINAL: {original}")
    else:
        print("Koi duplicate file nahi mili.")

# Aapka provide kiya gaya path
my_path = r'C:\Users\b0335496\Downloads\B0335496_Cell_Data_Report_Dec'

find_duplicate_excels(my_path)