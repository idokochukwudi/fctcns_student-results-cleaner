import zipfile
import pandas as pd

zip_path = "/home/ernest/student_result_cleaner/EXAMS_INTERNAL/BN/SET47/CLEAN_RESULTS/SET47_RESULT-2025-11-11_090146.zip"

with zipfile.ZipFile(zip_path, "r") as z:
    excel_files = [f for f in z.namelist() if f.endswith(".xlsx")]
    if not excel_files:
        raise ValueError("No Excel files found in ZIP!")

    with z.open(excel_files[0]) as f:
        # Load all sheets
        sheets = pd.read_excel(f, sheet_name=None, header=5)  # row 6 as header

# Loop through semester sheets
for sheet_name in sheets:
    if "FIRST" in sheet_name.upper():  # adjust filter if needed
        df = sheets[sheet_name]
        
        # Find REMARKS column dynamically
        remarks_cols = [c for c in df.columns if str(c).strip().upper() in ["REMARKS", "STATUS"]]
        if not remarks_cols:
            print(f"❌ REMARKS column not found in {sheet_name}")
            continue
        
        remarks_col = remarks_cols[0]
        print(f"✅ {sheet_name}: REMARKS column -> '{remarks_col}' (index {df.columns.get_loc(remarks_col)})")
        
        # Sample output
        print(df[[df.columns[1], remarks_col]].head(10))  # Name + Status
