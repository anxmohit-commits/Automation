import pandas as pd

# 1. New File Path
file_path = r"C:\Users\b0335496\Downloads\11 AM -2.xlsx"

# 2. Data Load
df = pd.read_excel(file_path)

# 3. Target Range Adjust (P se V tak)
# P is index 15, V is index 21. We use 15:22 to include V.
data_range = df.iloc[:, 15:22].astype(str).apply(lambda x: x.str.strip())

# 4. Row-wise Counting (W, X, Y columns mein mapping)
# W: Total < 11:00
df['W'] = ((data_range < "11:00") & (data_range != "NOSVD") & (data_range != "nan")).sum(axis=1)

# X: Total NOSVD
df['X'] = (data_range == "NOSVD").sum(axis=1)

# Y: Total >= 11:00
df['Y'] = ((data_range >= "11:00") & (data_range != "NOSVD") & (data_range != "nan")).sum(axis=1)

# 5. Column Headers Rename (Final Headers)
df.rename(columns={
    'W': 'Total < 11:00',
    'X': 'Total NOSVD',
    'Y': 'Total >= 11:00'
}, inplace=True)

# 6. Save back to the same file
df.to_excel(file_path, index=False)

print(f"Mapping Complete! Results saved in W, X, Y columns for file: {file_path}")