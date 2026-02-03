import pandas as pd
import json

# --- Paths ---
json_path = r"C:\Users\b0335496\Downloads\user_location_data (1).json"
excel_output = r"C:\Users\b0335496\Downloads\user_location_data_converted.xlsx"

def convert_json():
    try:
        # 1. JSON file ko load karein
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 2. JSON ko Flat DataFrame mein convert karein
        # json_normalize complex (nested) JSON ko rows aur columns mein sahi se set kar deta hai
        if isinstance(data, list):
            df = pd.json_normalize(data)
        else:
            # Agar JSON ek single object hai
            df = pd.json_normalize([data])

        # 3. Excel mein save karein
        df.to_excel(excel_output, index=False, engine='openpyxl')
        
        print(f"✅ Conversion Successful!")
        print(f"📂 Excel file yahan hai: {excel_output}")
        print(f"📊 Total records converted: {len(df)}")

    except FileNotFoundError:
        print("❌ Error: JSON file nahi mili. Path check karein.")
    except Exception as e:
        print(f"❌ Kuch galat hua: {e}")

if __name__ == "__main__":
    convert_json()