import pandas as pd
import os
from datetime import datetime

# --- Paths ---
raw_folder = "/mnt/c/Users/MTECH COMPUTERS/Documents/RAW_JAMB_DB"
clean_base_folder = "/mnt/c/Users/MTECH COMPUTERS/Documents/CLEAN_JAMB_DB"

# Create time-stamped folder for this run
timestamp_folder = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
clean_folder = os.path.join(clean_base_folder, f"clean_jamb_DB_{timestamp_folder}")
os.makedirs(clean_folder, exist_ok=True)

# --- Get list of CSV or Excel files ---
files = [f for f in os.listdir(raw_folder) 
         if f.lower().endswith((".csv", ".xlsx", ".xls")) and not f.startswith("~$")]

if not files:
    print("⚠️ No CSV or Excel files found in RAW_JAMB_DB.")
    exit()

# --- Split name function ---
def split_name(full_name):
    parts = str(full_name).split()
    if len(parts) == 0:
        return "", "", ""
    elif len(parts) == 1:
        return parts[0], "", ""
    elif len(parts) == 2:
        return parts[0], parts[1], ""
    else:
        return parts[0], parts[1], " ".join(parts[2:])

# --- Process each file ---
for file in files:
    raw_file_path = os.path.join(raw_folder, file)
    try:
        # Load CSV or Excel
        if file.lower().endswith(".csv"):
            df = pd.read_csv(raw_file_path)
        elif file.lower().endswith(".xlsx"):
            df = pd.read_excel(raw_file_path, engine='openpyxl')
        elif file.lower().endswith(".xls"):
            try:
                df = pd.read_excel(raw_file_path, engine='xlrd')
            except Exception:
                try:
                    df = pd.read_excel(raw_file_path, engine='openpyxl')
                except Exception:
                    raise ValueError("Unsupported format or corrupt file")
        else:
            continue

        # --- Check required column ---
        if 'RG_CANDNAME' not in df.columns:
            print(f"⚠️ Skipping {file}: 'RG_CANDNAME' column not found")
            continue

        # Split names
        df[['lastName', 'firstName', 'otherNames']] = df['RG_CANDNAME'].apply(
            lambda x: pd.Series(split_name(x))
        )

        # Create final dataset
        final_df = pd.DataFrame({
            'jambId': df.get('RG_NUM', ''),
            'lastName': df['lastName'],
            'firstName': df['firstName'],
            'otherNames': df['otherNames'],
            'gender': df.get('RG_SEX', ''),
            'state': df.get('STATE_NAME', ''),
            'lga': df.get('LGA_NAME', ''),
            'aggregateScore': df.get('RG_AGGREGATE', '')
        })

        # Save cleaned file with timestamp only
        timestamp_file = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_file = os.path.join(clean_folder, f"clean_jamb_DB_{timestamp_file}.csv")
        final_df.to_csv(output_file, index=False)

        print(f"✅ File cleaned and saved: {output_file}")

    except Exception as e:
        print(f"⚠️ Skipping {file}: {e}")

