import pandas as pd
import glob
import os
from datetime import datetime
import re

# -----------------------------
# Step 1: Define directories (Railway + Local compatible)
# -----------------------------
# Detect environment: Railway or Local
IS_RAILWAY = os.getenv('RAILWAY_ENVIRONMENT') is not None

if IS_RAILWAY:
    # Railway: BASE_DIR is set to /app/EXAMS_INTERNAL
    BASE_DIR = os.getenv('BASE_DIR', '/app/EXAMS_INTERNAL')
    print("🚂 Running on Railway")
    print(f"   BASE_DIR: {BASE_DIR}")
else:
    # Local: use home directory structure
    BASE_DIR = os.path.join(os.path.expanduser('~'), 'student_result_cleaner', 'EXAMS_INTERNAL')
    print("💻 Running locally")
    print(f"   BASE_DIR: {BASE_DIR}")

# Input directory - RAW_OBJ folder
input_dir = os.path.join(BASE_DIR, "OBJ_RESULT", "RAW_OBJ")

# Output directory - CLEAN_OBJ folder
output_base_dir = os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ")

# -----------------------------
# Step 2: Create timestamped folders with obj_result prefix
# -----------------------------
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
folder_name = f"obj_result_{timestamp}"
output_dir = os.path.join(output_base_dir, folder_name)

# Create output directory
os.makedirs(output_dir, exist_ok=True)

print(f"📁 Input Directory: {input_dir}")
print(f"📁 Output Directory: {output_dir}\n")

# -----------------------------
# Step 3: Helper to sanitize filenames
# -----------------------------
def sanitize_filename(name):
    """Remove characters that are invalid in filenames"""
    return re.sub(r'[<>:"/\\|?*]', "_", name)


# -----------------------------
# Step 4: Find all CSV and Excel files in RAW_OBJ
# -----------------------------
csv_files = glob.glob(os.path.join(input_dir, "*.csv"))
xls_files = glob.glob(os.path.join(input_dir, "*.xls")) + glob.glob(
    os.path.join(input_dir, "*.xlsx")
)

# Exclude temporary files (~$)
xls_files = [f for f in xls_files if not os.path.basename(f).startswith("~$")]
csv_files = [f for f in csv_files if not os.path.basename(f).startswith("~$")]

all_files = csv_files + xls_files

if not all_files:
    print("❌ No CSV or Excel files found in RAW_OBJ folder.")
    print(f"   Please check if files exist in: {input_dir}")
    exit()

print(f"📊 Found {len(all_files)} file(s) to process\n")

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
            df = pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        print(f"⚠️  Error reading {file_name}: {e}")
        continue

    # Dynamically detect Grade column (any column starting with 'Grade/')
    grade_cols = [col for col in df.columns if str(col).startswith("Grade/")]
    if not grade_cols:
        print(f"⚠️  Skipping {file_name}: No Grade column found")
        continue
    grade_col = grade_cols[0]  # pick the first match

    # Check required columns
    required_cols = ["Surname", "First name", grade_col]
    if not all(col in df.columns for col in required_cols):
        print(f"⚠️  Skipping {file_name}: required columns not found")
        continue

    # Keep only required columns
    cleaned_df = df[required_cols].copy()

    # Rename columns BUT keep Grade column name as-is
    cleaned_df.rename(
        columns={
            "Surname": "MAT NO.",
            "First name": "FULL NAME",
        },
        inplace=True,
    )

    # Sort by MAT NO.
    cleaned_df.sort_values(by="MAT NO.", key=lambda x: x.str.upper(), inplace=True)

    # Reset index + add Serial Number
    cleaned_df.reset_index(drop=True, inplace=True)
    cleaned_df.insert(0, "SN", range(1, len(cleaned_df) + 1))

    # Handle "Overall average" rows → remove SN
    mask = cleaned_df["MAT NO."].str.contains("Overall average", case=False, na=False)
    cleaned_df.loc[mask, "SN"] = ""

    # Output filename
    base_name = os.path.splitext(file_name)[0]
    safe_base_name = sanitize_filename(base_name)
    output_file = os.path.join(output_dir, f"cleaned_{safe_base_name}.csv")

    # Save cleaned CSV
    cleaned_df.to_csv(output_file, index=False)

    print(f"✅ Cleaned CSV saved: {output_file}\n")

    all_cleaned_dfs.append(cleaned_df)

# -----------------------------
# Step 6: Master CSV - Combine all cleaned results
# -----------------------------
if all_cleaned_dfs:
    master_df = pd.concat(all_cleaned_dfs, ignore_index=True)

    # Handle "Overall average" rows in master → remove SN
    mask_master = master_df["MAT NO."].str.contains(
        "Overall average", case=False, na=False
    )
    master_df.loc[mask_master, "SN"] = ""

    master_output = os.path.join(output_dir, "master_cleaned_results.csv")
    master_df.to_csv(master_output, index=False)

    print(f"🎉 Master CSV saved: {master_output}")
    print(f"   Total records: {len(master_df)}")

print("\n✅ All processing completed successfully!")
print(f"📂 Results saved in: {output_dir}")