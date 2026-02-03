import pandas as pd
import os

# Define the path to your parquet file
file_path = r'C:\Users\b0321755\Downloads\titanic.parquet'

# Check if the file exists
if os.path.exists(file_path):
    # Read the parquet file
    df = pd.read_parquet(file_path, engine='pyarrow')  # or engine='fastparquet'
    
    # Show first 5 rows
    print("✅ Parquet file read successfully!")
    print(df.head())
else:
    print("🚫 File not found at:", file_path)