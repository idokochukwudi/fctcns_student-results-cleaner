import pandas as pd
import os
from datetime import datetime

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

# Input directory - RAW_JAMB_DB folder
raw_folder = os.path.join(BASE_DIR, "JAMB_DB", "RAW_JAMB_DB")

# Output directory - CLEAN_JAMB_DB folder
clean_base_folder = os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB")

# -----------------------------
# Step 2: Create timestamped folder
# -----------------------------
timestamp_folder = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
clean_folder = os.path.join(clean_base_folder, f"clean_jamb_DB_{timestamp_folder}")
os.makedirs(clean_folder, exist_ok=True)

print(f"📁 Input Directory: {raw_folder}")
print(f"📁 Output Directory: {clean_folder}\n")

# -----------------------------
# Step 3: Get list of CSV or Excel files
# -----------------------------
if not os.path.exists(raw_folder):
    print(f"❌ Raw folder does not exist: {raw_folder}")
    exit()

files = [
    f
    for f in os.listdir(raw_folder)
    if f.lower().endswith((".csv", ".xlsx", ".xls")) and not f.startswith("~$")
]

if not files:
    print(f"❌ No CSV or Excel files found in {raw_folder}")
    exit()

print(f"📊 Found {len(files)} file(s) to process\n")


# -----------------------------
# Step 4: Split name function
# -----------------------------
def split_name(full_name):
    """Split full name into lastName, firstName, and otherNames"""
    parts = str(full_name).split()
    if len(parts) == 0:
        return "", "", ""
    elif len(parts) == 1:
        return parts[0], "", ""
    elif len(parts) == 2:
        return parts[0], parts[1], ""
    else:
        return parts[0], parts[1], " ".join(parts[2:])


# -----------------------------
# Step 5: Process each file
# -----------------------------
for file in files:
    raw_file_path = os.path.join(raw_folder, file)
    print(f"Processing: {file}")
    
    try:
        # Load CSV or Excel
        if file.lower().endswith(".csv"):
            df = pd.read_csv(raw_file_path)
        elif file.lower().endswith(".xlsx"):
            df = pd.read_excel(raw_file_path, engine="openpyxl")
        elif file.lower().endswith(".xls"):
            try:
                df = pd.read_excel(raw_file_path, engine="xlrd")
            except Exception:
                try:
                    df = pd.read_excel(raw_file_path, engine="openpyxl")
                except Exception:
                    raise ValueError("Unsupported format or corrupt file")
        else:
            continue

        # Check required column
        if "RG_CANDNAME" not in df.columns:
            print(f"⚠️  Skipping {file}: 'RG_CANDNAME' column not found\n")
            continue

        # Split names
        df[["lastName", "firstName", "otherNames"]] = df["RG_CANDNAME"].apply(
            lambda x: pd.Series(split_name(x))
        )

        # Create final dataset
        final_df = pd.DataFrame(
            {
                "jambId": df.get("RG_NUM", ""),
                "lastName": df["lastName"],
                "firstName": df["firstName"],
                "otherNames": df["otherNames"],
                "gender": df.get("RG_SEX", ""),
                "state": df.get("STATE_NAME", ""),
                "lga": df.get("LGA_NAME", ""),
                "aggregateScore": df.get("RG_AGGREGATE", ""),
            }
        )

        # Save cleaned file with timestamp only
        timestamp_file = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_file = os.path.join(clean_folder, f"clean_jamb_DB_{timestamp_file}.csv")
        final_df.to_csv(output_file, index=False)

        print(f"✅ Saved: {os.path.basename(output_file)}")
        print(f"   Total records: {len(final_df)}\n")

    except Exception as e:
        print(f"⚠️  Error processing {file}: {e}\n")

print("✅ Processing completed successfully!")
print(f"📂 Results saved in: {clean_folder}")