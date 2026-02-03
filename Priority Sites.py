import pandas as pd
import os
import re

# 1. Configuration
folder_path = r"C:\Users\b0335496\Downloads\Priority_Sites_11-01-2026"
source_sheet_1 = 'USD_SDO_LDI_Sites'
source_sheet_2 = 'summary'

# Metrics to be in SEPARATE sheets
separate_metrics = {
    source_sheet_1: ['Unique Sites Down', 'Total SDI', 'ISD Excluding SDI'],
    source_sheet_2: ['Total_y'] # Total_y ab yahan shamil hai
}

# Metrics to be COMBINED in one sheet
combined_metrics_list = ['>10 & <=20', '>20', 'Total_x']

final_separate_sheets = {}
summary_combined_list = []

# 2. Processing Files
files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]

for file_name in files:
    date_match = re.search(r'\d{2}-\d{2}-\d{4}', file_name)
    if not date_match: continue
    file_date = date_match.group()
    
    file_path = os.path.join(folder_path, file_name)
    xl = pd.ExcelFile(file_path)
    
    # Process Separate Sheets Metrics
    for src_sheet, cols in separate_metrics.items():
        actual_sheet = [n for n in xl.sheet_names if n.lower() == src_sheet.lower()]
        if actual_sheet:
            df = pd.read_excel(xl, sheet_name=actual_sheet[0])
            df['Circle'] = df['Circle'].astype(str).str.strip().str.upper()
            for col in cols:
                if col in df.columns:
                    s_name = f"Circle, {col}"
                    if s_name not in final_separate_sheets: final_separate_sheets[s_name] = pd.DataFrame()
                    final_separate_sheets[s_name][file_date] = df.set_index('Circle')[col]

    # Process Combined Summary Metrics
    s2_actual = [n for n in xl.sheet_names if n.lower() == source_sheet_2.lower()]
    if s2_actual:
        df2 = pd.read_excel(xl, sheet_name=s2_actual[0])
        df2['Circle'] = df2['Circle'].astype(str).str.strip().str.upper()
        needed_cols = ['Circle'] + [c for c in combined_metrics_list if c in df2.columns]
        temp_df = df2[needed_cols].copy().set_index('Circle')
        temp_df.columns = pd.MultiIndex.from_product([[file_date], temp_df.columns])
        summary_combined_list.append(temp_df)

# 3. Final Saving with Totals
output_file = os.path.join(folder_path, "Final_Full_Report_Fixed.xlsx")

with pd.ExcelWriter(output_file) as writer:
    # 1. Save Separate Sheets (including Total_y)
    for s_name, df_res in final_separate_sheets.items():
        if not df_res.empty:
            df_res = df_res.reindex(sorted(df_res.columns), axis=1).fillna(0)
            df_res['Grand Total'] = df_res.sum(axis=1)
            df_res.loc['Grand Total'] = df_res.sum(axis=0)
            df_res.to_excel(writer, sheet_name=s_name[:31])
    
    # 2. Save Combined Sheet
    if summary_combined_list:
        combined_df = pd.concat(summary_combined_list, axis=1).sort_index(axis=1).fillna(0)
        combined_df.loc['Grand Total'] = combined_df.sum(axis=0)
        # Row Total for combined metrics
        for metric in combined_metrics_list:
            m_cols = [c for c in combined_df.columns if c[1] == metric and c[0] != 'Grand Total']
            if m_cols:
                combined_df[('Overall Total', metric)] = combined_df[m_cols].sum(axis=1)
        combined_df.to_excel(writer, sheet_name='Summary_Combined_Metrics')

print(f"Ab saari 5 sheets ready hain, Total_y ke saath!\nPath: {output_file}")