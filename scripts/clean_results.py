import pandas as pd
import glob
import os
from datetime import datetime
import re

# -----------------------------
# Step 1: Define directories
# -----------------------------
input_dir = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/INTERNAL_RESULT/RAW_INTERNAL_RESULT"

base_wsl_output = "/home/ernest/PROCESS_RESULT/INTERNAL_RESULT/CLEAN_INTERNAL_RESULT"
base_win_output = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/INTERNAL_RESULT/CLEAN_INTERNAL_RESULT"

# -----------------------------
# Step 2: Create timestamped folders with obj_result prefix
# -----------------------------
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
folder_name = f"obj_result_{timestamp}"
output_dir = os.path.join(base_wsl_output, folder_name)
windows_doc_folder = os.path.join(base_win_output, folder_name)

os.makedirs(output_dir, exist_ok=True)
os.makedirs(windows_doc_folder, exist_ok=True)

# -----------------------------
# Step 3: Helper to sanitize filenames for Windows
# -----------------------------
def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name)

# -----------------------------
# Step 4: Find all CSV and Excel files in RAW_INTERNAL_RESULT
# -----------------------------
csv_files = glob.glob(os.path.join(input_dir, "*.csv"))
xls_files = glob.glob(os.path.join(input_dir, "*.xls")) + glob.glob(os.path.join(input_dir, "*.xlsx"))

# Exclude temporary files (~$)
xls_files = [f for f in xls_files if not os.path.basename(f).startswith('~$')]
csv_files = [f for f in csv_files if not os.path.basename(f).startswith('~$')]

all_files = csv_files + xls_files

if not all_files:
    print("❌ No CSV or Excel files found in RAW_INTERNAL_RESULT folder.")
    exit()

all_cleaned_dfs = []

# -----------------------------
# Step 5: Process each file
# -----------------------------
for file in all_files:
    file_name = os.path.basename(file)
    print(f"Processing: {file_name}")

    try:
        if file.lower().endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file, engine='openpyxl')
    except Exception as e:
        print(f"⚠️ Error reading {file_name}: {e}")
        continue

    # Dynamically detect Grade column (any column starting with 'Grade/')
    grade_cols = [col for col in df.columns if str(col).startswith("Grade/")]
    if not grade_cols:
        print(f"⚠️ Skipping {file_name}: No Grade column found")
        continue
    grade_col = grade_cols[0]  # pick the first match

    # Check required columns
    required_cols = ['Surname', 'First name', grade_col]
    if not all(col in df.columns for col in required_cols):
        print(f"⚠️ Skipping {file_name}: required columns not found")
        continue

    # Keep only required columns
    cleaned_df = df[required_cols].copy()

    # Rename columns BUT keep Grade column name as-is
    cleaned_df.rename(columns={
        'Surname': 'MAT NO.',
        'First name': 'FULL NAME',
    }, inplace=True)

    # Sort by MAT NO.
    cleaned_df.sort_values(by='MAT NO.', key=lambda x: x.str.upper(), inplace=True)

    # Reset index + add Serial Number
    cleaned_df.reset_index(drop=True, inplace=True)
    cleaned_df.insert(0, 'SN', range(1, len(cleaned_df) + 1))

    # Handle "Overall average" rows → remove SN
    mask = cleaned_df['MAT NO.'].str.contains('Overall average', case=False, na=False)
    cleaned_df.loc[mask, 'SN'] = ''

    # Output filenames
    base_name = os.path.splitext(file_name)[0]
    safe_base_name = sanitize_filename(base_name)

    wsl_output_file = os.path.join(output_dir, f"cleaned_{safe_base_name}.csv")
    windows_output_file = os.path.join(windows_doc_folder, f"cleaned_{safe_base_name}.csv")

    # Save
    cleaned_df.to_csv(wsl_output_file, index=False)
    cleaned_df.to_csv(windows_output_file, index=False)

    print(f"✅ Cleaned CSV saved in WSL: {wsl_output_file}")
    print(f"✅ Cleaned CSV saved in Windows Documents: {windows_output_file}\n")

    all_cleaned_dfs.append(cleaned_df)

# -----------------------------
# Step 6: Master CSV
# -----------------------------
if all_cleaned_dfs:
    master_df = pd.concat(all_cleaned_dfs, ignore_index=True)

    mask_master = master_df['MAT NO.'].str.contains('Overall average', case=False, na=False)
    master_df.loc[mask_master, 'SN'] = ''

    master_wsl = os.path.join(output_dir, "master_cleaned_results.csv")
    master_windows = os.path.join(windows_doc_folder, "master_cleaned_results.csv")

    master_df.to_csv(master_wsl, index=False)
    master_df.to_csv(master_windows, index=False)

    print(f"🎉 Master CSV saved in WSL: {master_wsl}")
    print(f"🎉 Master CSV saved in Windows Documents: {master_windows}")

print("✅ All processing completed successfully!")
