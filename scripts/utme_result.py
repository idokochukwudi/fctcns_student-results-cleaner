#!/usr/bin/env python3
"""
utme_result.py

Robust UTME/PUTME cleaning and reporting script (no matplotlib required).

Place raw files (CSV/XLS/XLSX) into:
  /mnt/c/Users/MTECH COMPUTERS/Documents/PUTME_RAW

Outputs cleaned CSV + formatted XLSX are saved to:
  /mnt/c/Users/MTECH COMPUTERS/Documents/PUTME_CLEAN

Features:
 - Maps common header variations to canonical names
 - Drops unwanted columns (Username, Department, State, Started on, Completed, Time taken)
 - Detects Grade/... column and renames to Score/...
 - Adds Score/100 (0 decimals) and optional Score/{N}% (0 decimals)
 - Removes 'overall average' rows and invalid rows
 - Sorts by STATE (A->Z) then Score (Z->A highest first)
 - Adds S/N (global) and STATE_SN (restarts per state)
 - Cleans phone numbers (remove .0, non-digits, ensure leading 0)
 - Highlights passes (>= PASS_THRESHOLD) in green
 - Applies a unique but stable row color per STATE
 - Produces Analysis sheet and polished Charts (no stray Series1)
"""

import os
import sys
import re
import math
import hashlib
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

# ---------------------------
# Configuration
# ---------------------------
RAW_DIR = "/mnt/c/Users/MTECH COMPUTERS/Documents/PUTME_RAW"
CLEAN_DIR = "/mnt/c/Users/MTECH COMPUTERS/Documents/PUTME_CLEAN"
PASS_THRESHOLD = 50.0                # highlight passes >= this
TIMESTAMP_FMT = "%Y%m%d_%H%M%S"      # timestamp format for output filenames

# Ensure directories exist
os.makedirs(RAW_DIR, exist_ok=True)
os.makedirs(CLEAN_DIR, exist_ok=True)

# ---------------------------
# Color palette (pastel-like) and stable hash mapping for states
# ---------------------------
PALETTE = [
    "FFF3E0", "E8F5E9", "E3F2FD", "F3E5F5", "FFFDE7",
    "E0F7FA", "FCE4EC", "EDE7F6", "FFF9C4", "E8F5E9",
    "F1F8E9", "E1F5FE", "F3E5F5", "FFEBEE", "ECEFF1",
    "FFF3E0", "F9FBE7", "EDE7F6", "F3E8FF", "E8F5E9"
]

def state_color_for(state_name):
    """
    Return a stable color for a state string by hashing the state name and
    selecting from PALETTE. Deterministic across runs.
    """
    if state_name is None:
        state_name = ""
    key = state_name.strip().lower().encode("utf8")
    h = int(hashlib.md5(key).hexdigest(), 16)
    color = PALETTE[h % len(PALETTE)]
    return PatternFill(start_color=color, end_color=color, fill_type="solid")

# ---------------------------
# Helpers
# ---------------------------
def find_column_by_names(df, candidate_names):
    """
    Find and return an actual column name present in df that matches any candidate (robust).
    Returns None if not found.
    """
    norm_map = {col: re.sub(r'\s+', ' ', str(col).strip().lower()) for col in df.columns}
    candidates = [re.sub(r'\s+', ' ', c.strip().lower()) for c in candidate_names]
    # exact normalized match
    for cand in candidates:
        for col, ncol in norm_map.items():
            if ncol == cand:
                return col
    # compacted match (remove spaces)
    for cand in candidates:
        candc = cand.replace(" ", "")
        for col, ncol in norm_map.items():
            if ncol.replace(" ", "") == candc:
                return col
    return None

def find_grade_column(df):
    """Return the first column that starts with 'Grade/' (case-insensitive), or a fallback containing 'grade' or 'score'."""
    for col in df.columns:
        if str(col).strip().lower().startswith("grade/"):
            return col
    # fallback: any column with 'grade' or 'score' in the name
    for col in df.columns:
        if re.search(r"grade|score", str(col), flags=re.I):
            return col
    return None

def to_numeric_safe(series):
    """Convert a pandas Series to numeric floats; coerce errors to NaN."""
    return pd.to_numeric(series.astype(str).str.replace(",", "").str.strip(), errors="coerce")

def clean_phone_value(s):
    """
    Normalize phone values:
     - treat NaN/None/blank as empty string
     - remove trailing .0 from Excel floats
     - keep only digits
     - convert leading '234' to '0...' (common pattern)
     - ensure leading 0
    """
    if pd.isna(s):
        return ""
    st = str(s).strip()
    if st.lower() in ("nan", "none", ""):
        return ""
    # remove .0 artifacts at end
    st = re.sub(r"\.0+$", "", st)
    # remove all non-digits
    digits = re.sub(r"\D", "", st)
    if digits == "":
        return ""
    # convert '234' prefix to '0'
    if digits.startswith("234") and len(digits) > 3:
        digits = "0" + digits[3:]
    if not digits.startswith("0"):
        digits = "0" + digits
    return digits

def drop_overall_average_rows(df):
    """
    Remove rows where any cell contains 'overall average' (case-insensitive).
    Useful to drop footers from exported grade files.
    """
    mask = df.apply(lambda row: row.astype(str).str.contains("overall average", case=False, na=False).any(), axis=1)
    if mask.any():
        return df[~mask].copy()
    return df

def auto_column_width(ws, min_width=8, max_width=60):
    """Auto-adjust column widths for an openpyxl worksheet."""
    for i, col in enumerate(ws.columns, 1):
        max_len = 0
        for cell in col:
            if cell.value is not None:
                l = len(str(cell.value))
                if l > max_len:
                    max_len = l
        ws.column_dimensions[get_column_letter(i)].width = min(max_width, max(min_width, max_len + 2))

# ---------------------------
# Core processing for a single file
# ---------------------------
def process_file(path):
    fname = os.path.basename(path)
    print(f"\n📂 Processing: {fname}")

    # Load CSV or Excel robustly
    try:
        if fname.lower().endswith(".csv"):
            df = pd.read_csv(path, dtype=str)
        else:
            try:
                df = pd.read_excel(path, dtype=str)
            except Exception:
                try:
                    df = pd.read_excel(path, dtype=str, engine="openpyxl")
                except Exception:
                    df = pd.read_excel(path, dtype=str, engine="xlrd")
    except Exception as e:
        print(f"❌ ERROR reading {fname}: {e}")
        return None

    # Normalize headers (strip)
    df.rename(columns=lambda c: str(c).strip(), inplace=True)

    # Map common names to canonical names
    # --- FULL NAME (user told 'First name' holds full name) ---
    fullname_col = find_column_by_names(df, ["First name", "Firstname", "Full name", "Name", "Surname"])
    if fullname_col:
        df.rename(columns={fullname_col: "FULL NAME"}, inplace=True)

    # APPLICATION ID if present (user sometimes has Surname/Mat no)
    appid_col = find_column_by_names(df, ["Surname", "Mat no", "Mat No", "MAT NO.", "APPLICATION ID", "Username"])
    if appid_col and appid_col != fullname_col:
        # map to APPLICATION ID only if different to FULL NAME column
        df.rename(columns={appid_col: "APPLICATION ID"}, inplace=True)

    # PHONE NUMBER mapping
    phone_col = find_column_by_names(df, ["Phone", "Phone number", "PHONE", "Mobile", "PhoneNumber"])
    if phone_col:
        df.rename(columns={phone_col: "PHONE NUMBER"}, inplace=True)

    # CITY/TOWN -> STATE
    city_col = find_column_by_names(df, ["City/town", "City /town", "City", "Town", "State of Origin", "STATE"])
    if city_col:
        df.rename(columns={city_col: "STATE"}, inplace=True)

    # Drop exactly requested columns (if present)
    for drop_col in ["Username", "Department", "State", "Started on", "Completed", "Time taken", "USERNAME", "DEPARTMENT"]:
        if drop_col in df.columns:
            df.drop(columns=[drop_col], inplace=True)

    # Ensure canonical columns exist for simpler downstream logic
    if "FULL NAME" not in df.columns:
        df["FULL NAME"] = pd.NA
    if "PHONE NUMBER" not in df.columns:
        df["PHONE NUMBER"] = pd.NA
    if "STATE" not in df.columns:
        df["STATE"] = pd.NA

    # Trim whitespace
    df["FULL NAME"] = df["FULL NAME"].astype(str).str.strip()
    df["STATE"] = df["STATE"].astype(str).str.strip()

    # Fill missing STATE with required label
    df.loc[df["STATE"].isin(["", "nan", "None"]), "STATE"] = "NO STATE OF ORIGIN"
    df["STATE"].fillna("NO STATE OF ORIGIN", inplace=True)

    # Find grade column: prefer Grade/... pattern
    grade_col = find_grade_column(df)
    if not grade_col:
        print(f"❌ Missing required column: a 'Grade/...' or 'Grade' column was not found in {fname}. Skipping.")
        return None

    # Rename Grade/... -> Score/... (preserve numeric suffix if present)
    if str(grade_col).strip().lower().startswith("grade/"):
        score_header = re.sub(r'(?i)^grade', 'Score', grade_col, flags=re.I)
    else:
        # fallback: if header contained digits use them, else generic 'Score'
        found = re.search(r"(\d+(?:\.\d+)?)", str(grade_col))
        suffix = found.group(1) if found else ""
        score_header = f"Score/{suffix}" if suffix else "Score"

    df.rename(columns={grade_col: score_header}, inplace=True)

    # Convert score to numeric for sorting and calculations
    df["_ScoreNum"] = to_numeric_safe(df[score_header])

    # Drop 'overall average' rows and invalid rows (missing name or score)
    df = drop_overall_average_rows(df)
    df["FULL NAME"] = df["FULL NAME"].astype(str).str.strip()
    df = df[~df["FULL NAME"].isin(["", "nan", "None"])].copy()
    df = df[df["_ScoreNum"].notna()].copy()

    if df.empty:
        print("⚠️ No valid rows remain after cleaning; skipping file.")
        return None

    # Clean phone column to avoid '0nan' etc.
    df["PHONE NUMBER"] = df["PHONE NUMBER"].apply(clean_phone_value)

    # Sort by STATE A->Z, Score Z->A (highest first)
    df.sort_values(by=["STATE", "_ScoreNum"], ascending=[True, False], inplace=True, na_position="last")
    df.reset_index(drop=True, inplace=True)

    # Insert global S/N and per-state STATE_SN
    df.insert(0, "S/N", range(1, len(df) + 1))
    df["STATE_SN"] = df.groupby("STATE").cumcount() + 1

    # Build cleaned DataFrame with the requested ordering
    out_cols = ["S/N", "STATE_SN", "FULL NAME"]
    if "APPLICATION ID" in df.columns:
        out_cols.append("APPLICATION ID")
    out_cols += ["PHONE NUMBER", "STATE", score_header]

    # Ensure out_cols exist
    for c in out_cols:
        if c not in df.columns:
            df[c] = pd.NA

    cleaned = df[out_cols].copy()

    # Add Score/100 (0 decimals) normalized column
    found_max = re.search(r"(\d+(?:\.\d+)?)", str(score_header))
    original_max = float(found_max.group(1)) if found_max else 100.0
    cleaned["Score/100"] = ((df["_ScoreNum"].astype(float) / original_max) * 100).round(0).astype("Int64")

    # Prompt for optional converted column (and create integer (0 decimals))
    add_conv = input("Add converted score column (e.g. convert Score/100 to Score/60)? (y/n): ").strip().lower()
    if add_conv in ("y", "yes"):
        tgt_raw = input("Enter target maximum (integer), e.g. 60: ").strip()
        try:
            tgt = float(tgt_raw)
            new_col = f"Score/{int(tgt)}%"
            cleaned[new_col] = ((df["_ScoreNum"].astype(float) / original_max) * tgt).round(0).astype("Int64")
            print(f"✅ Added converted column '{new_col}'.")
        except Exception as e:
            print("⚠️ Invalid conversion value; skipped converted column.", e)

    # Remove obvious stray empty rows where both FULL NAME and PHONE are empty
    cleaned = cleaned[~((cleaned["FULL NAME"].astype(str).str.strip() == "") &
                        (cleaned["PHONE NUMBER"].astype(str).str.strip() == ""))].copy()

    # Prepare timestamp-only output filenames (user requested no original filename appended)
    ts = datetime.now().strftime(TIMESTAMP_FMT)
    base_name = f"UTME_RESULT_{ts}"
    out_csv = os.path.join(CLEAN_DIR, base_name + ".csv")
    out_xlsx = os.path.join(CLEAN_DIR, base_name + ".xlsx")

    # Save cleaned CSV and initial XLSX (pandas)
    cleaned.to_csv(out_csv, index=False)
    try:
        cleaned.to_excel(out_xlsx, index=False, engine="openpyxl")
    except Exception:
        cleaned.to_excel(out_xlsx, index=False)
    print(f"Saved cleaned CSV: {out_csv}")
    print(f"Saved cleaned XLSX (pre-format): {out_xlsx}")

    # ---------------------------
    # Excel formatting (openpyxl) + Analysis + Charts
    # ---------------------------
    wb = load_workbook(out_xlsx)
    ws = wb.active
    ws.title = "Results"

    # Header formatting (bold + center)
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Left-align APPLICATION ID column if present
    if "APPLICATION ID" in cleaned.columns:
        try:
            col_idx = list(cleaned.columns).index("APPLICATION ID") + 1
            for r in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx, max_row=ws.max_row):
                for cell in r:
                    cell.alignment = Alignment(horizontal="left")
        except Exception:
            pass

    # Freeze header
    ws.freeze_panes = "A2"

    # Auto-adjust widths
    auto_column_width(ws)

    # Add thin borders to results
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in r:
            cell.border = border

    # Apply stable unique row color per state (light pastel) - do this first so pass highlights can override score cells
    state_col_index = list(cleaned.columns).index("STATE") + 1
    # iterate rows and fill row by state
    for row_idx in range(2, ws.max_row + 1):
        state_val = ws.cell(row=row_idx, column=state_col_index).value
        fill = state_color_for(state_val)
        for col_idx in range(1, ws.max_column + 1):
            # do not color header row (we start from row 2)
            ws.cell(row=row_idx, column=col_idx).fill = fill

    # Highlight passing scores (>= PASS_THRESHOLD) in Score/100 and any Score/{N}% and Score/... columns
    score_cols_indices = [i + 1 for i, c in enumerate(cleaned.columns) if str(c).startswith("Score/")]
    pass_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    pass_font = Font(color="006100")
    for col_idx in score_cols_indices:
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            try:
                # treat blank as not passing
                if cell.value is None or str(cell.value).strip() == "":
                    continue
                val = float(cell.value)
                if val >= PASS_THRESHOLD:
                    # override row fill for this cell with pass highlight
                    cell.fill = pass_fill
                    cell.font = pass_font
            except Exception:
                # ignore non-numeric
                continue

    # ---------------------------
    # Analysis sheet - summary + per-state table + narrative
    # ---------------------------
    if "Analysis" in wb.sheetnames:
        wb.remove(wb["Analysis"])
    anal = wb.create_sheet("Analysis")

    total_candidates = len(cleaned)
    score_series = df["_ScoreNum"].dropna().astype(float)
    highest_score = score_series.max() if not score_series.empty else None
    lowest_score = score_series.min() if not score_series.empty else None
    avg_score = round(score_series.mean(), 2) if not score_series.empty else None

    # Per-state statistics
    state_counts = df.groupby("STATE")["_ScoreNum"].count().sort_index()
    state_avg = df.groupby("STATE")["_ScoreNum"].mean().round(2).sort_index()
    state_pass_count = df[df["_ScoreNum"].astype(float) >= PASS_THRESHOLD].groupby("STATE")["_ScoreNum"].count()

    # Write high-level metrics
    anal.append(["Metric", "Value"])
    anal.append(["Total Candidates", int(total_candidates)])
    anal.append(["Highest Score (raw)", float(highest_score) if highest_score is not None else None])
    anal.append(["Lowest Score (raw)", float(lowest_score) if lowest_score is not None else None])
    anal.append(["Average Score (raw)", avg_score])
    anal.append([])

    # Per-state table header
    anal.append(["State", "Candidates", "Average Score", f"Pass Count (≥{int(PASS_THRESHOLD)})", "Pass Rate (%)"])
    per_state_rows = []
    for st in sorted(set(df["STATE"].tolist())):
        cnt = int(state_counts.get(st, 0))
        avg = float(state_avg.get(st, float("nan"))) if st in state_avg.index else None
        pcnt = int(state_pass_count.get(st, 0)) if st in state_pass_count.index else 0
        prate = round((pcnt / cnt * 100), 2) if cnt > 0 else 0.0
        per_state_rows.append([st, cnt, avg if not (isinstance(avg, float) and math.isnan(avg)) else None, pcnt, prate])
    for row in per_state_rows:
        anal.append(row)

    # Bold headers in analysis sheet
    anal["A1"].font = Font(bold=True)
    anal["B1"].font = Font(bold=True)
    # bold per-state header row (find the row)
    for i, r in enumerate(anal.iter_rows(values_only=True), start=1):
        if r and r[0] == "State" and r[1] == "Candidates":
            for cell in anal[i]:
                cell.font = Font(bold=True)
            break

    auto_column_width(anal)

    # Narrative summary (concise insights)
    if per_state_rows:
        sorted_by_count = sorted(per_state_rows, key=lambda r: r[1], reverse=True)
        most_state, most_cnt = sorted_by_count[0][0], sorted_by_count[0][1]
        least_state, least_cnt = sorted_by_count[-1][0], sorted_by_count[-1][1]
        sorted_by_avg = sorted([r for r in per_state_rows if r[2] is not None], key=lambda r: r[2], reverse=True)
        best_avg_state = sorted_by_avg[0][0] if sorted_by_avg else None
        sorted_by_pr = sorted(per_state_rows, key=lambda r: r[4], reverse=True)
        best_pass_state = sorted_by_pr[0][0] if sorted_by_pr else None

        anal.append([])
        anal.append(["Narrative Summary"])
        anal.append([f"Most candidates: {most_state} ({most_cnt})"])
        anal.append([f"Fewest candidates: {least_state} ({least_cnt})"])
        if best_avg_state:
            anal.append([f"Highest average score: {best_avg_state}"])
        if best_pass_state:
            anal.append([f"Highest pass rate: {best_pass_state}"])

    # ---------------------------
    # Charts sheet (polished)
    # ---------------------------
    if "Charts" in wb.sheetnames:
        wb.remove(wb["Charts"])
    chs = wb.create_sheet("Charts")

    # Prepare chart table sorted by candidate count descending
    chs.append(["State", "Candidates", "Average Score", "Pass Rate (%)"])
    counts_sorted = sorted(per_state_rows, key=lambda r: r[1], reverse=True)
    for r in counts_sorted:
        chs.append([r[0], r[1], r[2] if r[2] is not None else 0, r[4]])

    n = len(counts_sorted)
    if n > 0:
        # Candidates-per-state chart
        try:
            c1 = BarChart()
            c1.title = "Number of Candidates per State"
            c1.x_axis.title = "State"
            c1.y_axis.title = "Candidates"
            c1.width = 28
            c1.height = 14
            c1.gapWidth = 50
            data_ref = Reference(chs, min_col=2, min_row=2, max_row=1 + n)
            cats_ref = Reference(chs, min_col=1, min_row=2, max_row=1 + n)
            c1.add_data(data_ref, titles_from_data=False)
            # ensure series title is proper (avoid Series1)
            try:
                c1.series[0].title = "Candidates"
            except Exception:
                pass
            c1.set_categories(cats_ref)
            c1.dLbls = DataLabelList()
            c1.dLbls.showVal = True
            chs.add_chart(c1, "F2")
        except Exception:
            pass

        # Average Score per State chart
        try:
            c2 = BarChart()
            c2.title = "Average Score per State"
            c2.x_axis.title = "State"
            c2.y_axis.title = "Avg Score"
            c2.width = 28
            c2.height = 14
            c2.gapWidth = 50
            data_ref2 = Reference(chs, min_col=3, min_row=2, max_row=1 + n)
            cats_ref2 = Reference(chs, min_col=1, min_row=2, max_row=1 + n)
            c2.add_data(data_ref2, titles_from_data=False)
            try:
                c2.series[0].title = "Average Score"
            except Exception:
                pass
            c2.set_categories(cats_ref2)
            c2.dLbls = DataLabelList()
            c2.dLbls.showVal = True
            chs.add_chart(c2, "F22")
        except Exception:
            pass

        # Pass Rate per State chart
        try:
            c3 = BarChart()
            c3.title = f"Pass Rate per State (≥ {int(PASS_THRESHOLD)})"
            c3.x_axis.title = "State"
            c3.y_axis.title = "Pass Rate (%)"
            c3.width = 28
            c3.height = 14
            c3.gapWidth = 50
            data_ref3 = Reference(chs, min_col=4, min_row=2, max_row=1 + n)
            cats_ref3 = Reference(chs, min_col=1, min_row=2, max_row=1 + n)
            c3.add_data(data_ref3, titles_from_data=False)
            try:
                c3.series[0].title = "Pass Rate (%)"
            except Exception:
                pass
            c3.set_categories(cats_ref3)
            c3.dLbls = DataLabelList()
            c3.dLbls.showVal = True
            chs.add_chart(c3, "F42")
        except Exception:
            pass

    # Score distribution - 10-point bins added below charts table and charted
    try:
        score_vals = score_series.dropna().astype(float)
        if len(score_vals) > 0:
            bins = list(range(0, 101, 10))
            dist = pd.cut(score_vals, bins=bins, right=False)
            freq = dist.value_counts().sort_index()
            insert_row = chs.max_row + 2
            chs.cell(row=insert_row, column=1, value="Score Range")
            chs.cell(row=insert_row, column=2, value="Candidates")
            r = insert_row + 1
            for rng, cnt in freq.items():
                chs.cell(row=r, column=1, value=str(rng))
                chs.cell(row=r, column=2, value=int(cnt))
                r += 1
            n_bins = len(freq)
            try:
                c4 = BarChart()
                c4.title = "Score Distribution (10-point ranges)"
                c4.x_axis.title = "Range"
                c4.y_axis.title = "Candidates"
                c4.width = 28
                c4.height = 12
                c4.gapWidth = 50
                data_ref4 = Reference(chs, min_col=2, min_row=insert_row+1, max_row=insert_row+n_bins)
                cats_ref4 = Reference(chs, min_col=1, min_row=insert_row+1, max_row=insert_row+n_bins)
                c4.add_data(data_ref4, titles_from_data=False)
                try:
                    c4.series[0].title = "Candidates"
                except Exception:
                    pass
                c4.set_categories(cats_ref4)
                c4.dLbls = DataLabelList()
                c4.dLbls.showVal = True
                chs.add_chart(c4, f"F{insert_row}")
            except Exception:
                pass
    except Exception:
        pass

    # Auto width for Analysis and Charts
    auto_column_width(chs)
    auto_column_width(anal)

    # Save final workbook
    wb.save(out_xlsx)
    print(f"✅ Final Excel saved with Analysis & Charts: {out_xlsx}")
    print(f"  (CSV also available at: {out_csv})")

    # Console summary
    print("\nSummary:")
    print(f"  Total candidates processed: {total_candidates}")
    print(f"  Highest raw score: {highest_score}")
    print(f"  Lowest raw score: {lowest_score}")
    print(f"  Average raw score: {avg_score}")
    if per_state_rows:
        print(f"  State with most candidates: {most_state} ({most_cnt})")
        print(f"  State with fewest candidates: {least_state} ({least_cnt})")

    return out_csv, out_xlsx

# ---------------------------
# Entrypoint
# ---------------------------
def main():
    print("Starting UTME Results Cleaning...")
    files = [f for f in os.listdir(RAW_DIR) if f.lower().endswith((".csv", ".xlsx", ".xls"))]
    if not files:
        print(f"❌ No raw files found in {RAW_DIR}\nPut your raw file(s) there and re-run.")
        return

    outputs = []
    for f in files:
        try:
            res = process_file(os.path.join(RAW_DIR, f))
            if res:
                outputs.append(res)
        except Exception as e:
            print(f"❌ ERROR processing {f}: {e}", file=sys.stderr)

    print("\n✅ All done. Cleaned files (CSV + XLSX) are in:", CLEAN_DIR)
    if outputs:
        for csvp, xl in outputs:
            print(" -", csvp)
            print(" -", xl)

if __name__ == "__main__":
    main()

