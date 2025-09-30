#!/usr/bin/env python3
"""
utme_result.py

Robust UTME/PUTME cleaning and reporting script (no matplotlib required).

Place raw files (CSV/XLS/XLSX) into:
  PROCESS_RESULT/PUTME_RESULT/RAW_PUTME_RESULT

Put raw candidate batches (optional) into:
  PROCESS_RESULT/PUTME_RESULT/RAW_CANDIDATE_BATCHES

Outputs cleaned CSV + formatted XLSX are saved to:
  PROCESS_RESULT/PUTME_RESULT/CLEAN_PUTME_RESULT/UTME_RESULT-<timestamp>

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
 - Applies a soft, pastel row color per STATE
 - Produces Analysis sheet and clear, management-friendly Charts with a legend table
 - Lists registered-but-absent candidates in Absent sheet with APPLICATION ID
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
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

# ---------------------------
# Directory configuration
# ---------------------------
BASE_DIR = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/PUTME_RESULT"
RAW_DIR = os.path.join(BASE_DIR, "RAW_PUTME_RESULT")
CANDIDATE_DIR = os.path.join(BASE_DIR, "RAW_CANDIDATE_BATCHES")
CLEAN_DIR = os.path.join(BASE_DIR, "CLEAN_PUTME_RESULT")

PASS_THRESHOLD = 50.0                # highlight passes >= this
TIMESTAMP_FMT = "%Y-%m-%d_%H:%M:%S"  # well-structured timestamp format

# Ensure directories exist
os.makedirs(RAW_DIR, exist_ok=True)
os.makedirs(CLEAN_DIR, exist_ok=True)
os.makedirs(CANDIDATE_DIR, exist_ok=True)

# ---------------------------
# Soft pastel color palette for states and charts
# ---------------------------
PALETTE = [
    "A8E6CF",  # Soft Green
    "D7B9D5",  # Soft Purple
    "FFCCCB",  # Soft Red
    "B3CDE0",  # Soft Blue
    "F4E1A4",  # Soft Yellow
    "D1E8E2",  # Soft Mint
    "E4C1F9",  # Soft Lavender
    "F8E9A1",  # Soft Cream
    "A2D2FF",  # Soft Sky
    "D9BF77",  # Soft Beige
    "C6D4E1",  # Soft Grey-Blue
    "F7D6E0",  # Soft Pink
    "B5EAD7",  # Soft Seafoam
    "E2D1F9",  # Soft Lilac
    "F9AFA4",  # Soft Coral
    "B8E1FF",  # Soft Azure
    "D8E2DC",  # Soft Neutral
    "FFE5D9",  # Soft Peach
    "A1C7E0",  # Soft Slate
    "D4A5A5"   # Soft Rose
]

def state_color_for(state_name):
    """Deterministic pastel fill for a state label."""
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
    norm_map = {col: re.sub(r'\s+', ' ', str(col).strip().lower()) for col in df.columns}
    candidates = [re.sub(r'\s+', ' ', c.strip().lower()) for c in candidate_names]
    for cand in candidates:
        for col, ncol in norm_map.items():
            if ncol == cand:
                return col
    for cand in candidates:
        candc = cand.replace(" ", "")
        for col, ncol in norm_map.items():
            if ncol.replace(" ", "") == candc:
                return col
    return None

def find_grade_column(df):
    for col in df.columns:
        if str(col).strip().lower().startswith("grade/"):
            return col
    for col in df.columns:
        if re.search(r"grade|score", str(col), flags=re.I):
            return col
    return None

def to_numeric_safe(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").str.strip(), errors="coerce")

def clean_phone_value(s):
    if pd.isna(s):
        return ""
    st = str(s).strip()
    if st.lower() in ("nan", "none", ""):
        return ""
    st = re.sub(r"\.0+$", "", st)
    digits = re.sub(r"\D", "", st)
    if digits == "":
        return ""
    if digits.startswith("234") and len(digits) > 3:
        digits = "0" + digits[3:]
    if not digits.startswith("0"):
        digits = "0" + digits
    return digits

def drop_overall_average_rows(df):
    mask = df.apply(lambda row: row.astype(str).str.contains("overall average", case=False, na=False).any(), axis=1)
    if mask.any():
        return df[~mask].copy()
    return df

def auto_column_width(ws, min_width=8, max_width=60):
    for i, col in enumerate(ws.columns, 1):
        max_len = 0
        for cell in col:
            if cell.value is not None:
                l = len(str(cell.value))
                if l > max_len:
                    max_len = l
        ws.column_dimensions[get_column_letter(i)].width = min(max_width, max(min_width, max_len + 2))

# ---------------------------
# Candidate batches helpers
# ---------------------------
def load_candidate_batches(folder):
    """Load all candidate-batch files and return a DataFrame with standardized EXAM_NO, FULL_NAME, PHONE NUMBER, STATE."""
    if not os.path.isdir(folder):
        return pd.DataFrame(columns=["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE"])
    files = [f for f in os.listdir(folder) if f.lower().endswith((".csv", ".xlsx", ".xls")) and not f.startswith("~$")]
    rows = []
    for fname in sorted(files):
        path = os.path.join(folder, fname)
        try:
            if fname.lower().endswith(".csv"):
                cdf = pd.read_csv(path, dtype=str)
            else:
                cdf = pd.read_excel(path, dtype=str)
        except Exception:
            continue
        exam_col = find_column_by_names(cdf, ["username", "exam no", "reg no", "mat no", "regnum", "reg number"])
        name_col = find_column_by_names(cdf, ["firstname", "full name", "name", "candidate name", "user full name", "RG_CANDNAME"])
        phone_col = find_column_by_names(cdf, ["phone1", "phone", "phone number", "PHONE", "Mobile", "PhoneNumber"])
        state_col = find_column_by_names(cdf, ["city", "city/town", "City /town", "City", "Town", "State of Origin", "STATE"])
        for _, r in cdf.iterrows():
            ex = None
            name = None
            phone = None
            state = None
            if exam_col and pd.notna(r.get(exam_col)):
                ex = str(r.get(exam_col)).strip()
            if name_col and pd.notna(r.get(name_col)):
                name = str(r.get(name_col)).strip()
            if phone_col and pd.notna(r.get(phone_col)):
                phone = str(r.get(phone_col)).strip()
            if state_col and pd.notna(r.get(state_col)):
                state = str(r.get(state_col)).strip()
            if (not name) and ex and re.search(r"[A-Za-z]{2,}", ex):
                m = re.search(r"(.+?)\s+(\d{3,})$", ex)
                if m:
                    name = m.group(1).strip()
                    ex = m.group(2).strip()
            if not ex or not re.search(r"\d", ex):
                for c in cdf.columns:
                    v = r.get(c)
                    if pd.isna(v):
                        continue
                    s = str(v).strip()
                    m = re.search(r"\b(\d{3,})\b", s)
                    if m and not ex:
                        ex = m.group(1)
                    if not name and re.search(r"[A-Za-z]{2,}", s) and not re.fullmatch(r"\d+", s):
                        name = s
                    if not phone and re.search(r"\d{10,}", s):
                        phone = s
                    if not state and re.search(r"[A-Za-z]{2,}", s) and c.lower() in ["city", "state"]:
                        state = s
            if ex:
                rows.append({"EXAM_NO": ex, "FULL_NAME": name or "", "PHONE NUMBER": phone or "", "STATE": state or ""})
    if not rows:
        return pd.DataFrame(columns=["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE"])
    cdf_all = pd.DataFrame(rows).drop_duplicates(subset=["EXAM_NO"])
    if "PHONE NUMBER" in cdf_all.columns:
        cdf_all["PHONE NUMBER"] = cdf_all["PHONE NUMBER"].apply(clean_phone_value)
    return cdf_all

# ---------------------------
# Extract numeric tokens helper
# ---------------------------
def extract_numeric_token(s):
    """Return first sequence of 3+ digits found in string, else None."""
    if s is None:
        return None
    s = str(s)
    m = re.search(r"(\d{3,})", s)
    return m.group(1) if m else None

# ---------------------------
# Core processing for a single file
# ---------------------------
def process_file(path, output_dir, ts):
    fname = os.path.basename(path)
    print(f"\n📂 Processing: {fname}")

    # Load
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

    df.rename(columns=lambda c: str(c).strip(), inplace=True)

    # Map canonical names
    fullname_col = find_column_by_names(df, ["First name", "Firstname", "Full name", "Name", "Surname"])
    if fullname_col:
        df.rename(columns={fullname_col: "FULL NAME"}, inplace=True)

    appid_col = find_column_by_names(df, ["Surname", "Mat no", "Mat No", "MAT NO.", "APPLICATION ID", "Username"])
    if appid_col and appid_col != fullname_col:
        df.rename(columns={appid_col: "APPLICATION ID"}, inplace=True)

    phone_col = find_column_by_names(df, ["Phone", "Phone number", "PHONE", "Mobile", "PhoneNumber"])
    if phone_col:
        df.rename(columns={phone_col: "PHONE NUMBER"}, inplace=True)

    city_col = find_column_by_names(df, ["City/town", "City /town", "City", "Town", "State of Origin", "STATE"])
    if city_col:
        df.rename(columns={city_col: "STATE"}, inplace=True)

    # drop requested columns
    for drop_col in ["Username", "Department", "State", "Started on", "Completed", "Time taken", "USERNAME", "DEPARTMENT"]:
        if drop_col in df.columns:
            df.drop(columns=[drop_col], inplace=True)

    # ensure canonical columns
    if "FULL NAME" not in df.columns:
        df["FULL NAME"] = pd.NA
    if "PHONE NUMBER" not in df.columns:
        df["PHONE NUMBER"] = pd.NA
    if "STATE" not in df.columns:
        df["STATE"] = pd.NA

    df["FULL NAME"] = df["FULL NAME"].astype(str).str.strip()
    df["STATE"] = df["STATE"].astype(str).str.strip()
    df.loc[df["STATE"].isin(["", "nan", "None"]), "STATE"] = "NO STATE OF ORIGIN"
    df["STATE"].fillna("NO STATE OF ORIGIN", inplace=True)

    # find grade column
    grade_col = find_grade_column(df)
    if not grade_col:
        print(f"❌ Missing required column: a 'Grade/...' or 'Grade' column was not found in {fname}. Skipping.")
        return None

    if str(grade_col).strip().lower().startswith("grade/"):
        score_header = re.sub(r'(?i)^grade', 'Score', grade_col, flags=re.I)
    else:
        found = re.search(r"(\d+(?:\.\d+)?)", str(grade_col))
        suffix = found.group(1) if found else ""
        score_header = f"Score/{suffix}" if suffix else "Score"

    df.rename(columns={grade_col: score_header}, inplace=True)

    df["_ScoreNum"] = to_numeric_safe(df[score_header])

    df = drop_overall_average_rows(df)
    df["FULL NAME"] = df["FULL NAME"].astype(str).str.strip()
    df = df[~df["FULL NAME"].isin(["", "nan", "None"])].copy()
    df = df[df["_ScoreNum"].notna()].copy()

    if df.empty:
        print("⚠️ No valid rows remain after cleaning; skipping file.")
        return None

    df["PHONE NUMBER"] = df["PHONE NUMBER"].apply(clean_phone_value)

    # sort and serials
    df.sort_values(by=["STATE", "_ScoreNum"], ascending=[True, False], inplace=True, na_position="last")
    df.reset_index(drop=True, inplace=True)
    df.insert(0, "S/N", range(1, len(df) + 1))
    df["STATE_SN"] = df.groupby("STATE").cumcount() + 1

    out_cols = ["S/N", "STATE_SN", "FULL NAME"]
    if "APPLICATION ID" in df.columns:
        out_cols.append("APPLICATION ID")
    out_cols += ["PHONE NUMBER", "STATE", score_header]

    for c in out_cols:
        if c not in df.columns:
            df[c] = pd.NA

    cleaned = df[out_cols].copy()

    # Add normalized Score/100
    found_max = re.search(r"(\d+(?:\.\d+)?)", str(score_header))
    original_max = float(found_max.group(1)) if found_max else 100.0
    cleaned["Score/100"] = ((df["_ScoreNum"].astype(float) / original_max) * 100).round(0).astype("Int64")

    # Optional converted score column
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

    cleaned = cleaned[~((cleaned["FULL NAME"].astype(str).str.strip() == "") &
                        (cleaned["PHONE NUMBER"].astype(str).str.strip() == ""))].copy()

    # prepare output filenames
    base_name = f"UTME_RESULT_{os.path.splitext(fname)[0]}_{ts}"
    out_csv = os.path.join(output_dir, base_name + ".csv")
    out_xlsx = os.path.join(output_dir, base_name + ".xlsx")

    cleaned.to_csv(out_csv, index=False)
    try:
        cleaned.to_excel(out_xlsx, index=False, engine="openpyxl")
    except Exception:
        cleaned.to_excel(out_xlsx, index=False)
    print(f"Saved cleaned CSV: {out_csv}")
    print(f"Saved cleaned XLSX (pre-format): {out_xlsx}")

    # ---------------------------
    # Excel formatting + Analysis + Charts
    # ---------------------------
    wb = load_workbook(out_xlsx)
    ws = wb.active
    ws.title = "Results"

    header_font = Font(bold=True, size=12)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    if "APPLICATION ID" in cleaned.columns:
        try:
            col_idx = list(cleaned.columns).index("APPLICATION ID") + 1
            for r in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx, max_row=ws.max_row):
                for cell in r:
                    cell.alignment = Alignment(horizontal="left")
        except Exception:
            pass

    ws.freeze_panes = "A2"
    auto_column_width(ws)

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in r:
            cell.border = border

    # state row coloring with soft palette
    try:
        state_col_index = list(cleaned.columns).index("STATE") + 1
        for row_idx in range(2, ws.max_row + 1):
            state_val = ws.cell(row=row_idx, column=state_col_index).value
            fill = state_color_for(state_val)
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill
    except Exception:
        pass

    # highlight passes
    score_cols_indices = [i + 1 for i, c in enumerate(cleaned.columns) if str(c).startswith("Score/")]
    pass_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    pass_font = Font(color="006100")
    for col_idx in score_cols_indices:
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            try:
                if cell.value is None or str(cell.value).strip() == "":
                    continue
                val = float(cell.value)
                if val >= PASS_THRESHOLD:
                    cell.fill = pass_fill
                    cell.font = pass_font
            except Exception:
                continue

    # ---------------------------
    # Analysis sheet
    # ---------------------------
    if "Analysis" in wb.sheetnames:
        wb.remove(wb["Analysis"])
    anal = wb.create_sheet("Analysis")

    total_candidates = len(cleaned)
    score_series = df["_ScoreNum"].dropna().astype(float) if "_ScoreNum" in df.columns else pd.Series([], dtype=float)
    highest_score = score_series.max() if not score_series.empty else None
    lowest_score = score_series.min() if not score_series.empty else None
    avg_score = round(score_series.mean(), 2) if not score_series.empty else None

    state_counts = df.groupby("STATE")["_ScoreNum"].count().sort_index() if "_ScoreNum" in df.columns else pd.Series(dtype=int)
    state_avg = df.groupby("STATE")["_ScoreNum"].mean().round(2).sort_index() if "_ScoreNum" in df.columns else pd.Series(dtype=float)
    state_pass_count = df[df["_ScoreNum"].astype(float) >= PASS_THRESHOLD].groupby("STATE")["_ScoreNum"].count() if "_ScoreNum" in df.columns else pd.Series(dtype=int)

    anal.append(["Metric", "Value"])
    anal.append(["Total Candidates", int(total_candidates)])
    anal.append(["Highest Score (raw)", float(highest_score) if highest_score is not None else None])
    anal.append(["Lowest Score (raw)", float(lowest_score) if lowest_score is not None else None])
    anal.append(["Average Score (raw)", avg_score])
    anal.append([])

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

    anal["A1"].font = Font(bold=True, size=12)
    anal["B1"].font = Font(bold=True, size=12)
    for i, r in enumerate(anal.iter_rows(values_only=True), start=1):
        if r and r[0] == "State" and r[1] == "Candidates":
            for cell in anal[i]:
                cell.font = Font(bold=True, size=12)
            break

    auto_column_width(anal)

    # Narrative summary
    if per_state_rows:
        sorted_by_count = sorted(per_state_rows, key=lambda r: r[1], reverse=True)
        most_state, most_cnt = sorted_by_count[0][0], sorted_by_count[0][1]
        least_state, least_cnt = sorted_by_count[-1][0], sorted_by_count[-1][1]
        sorted_by_avg = sorted([r for r in per_state_rows if r[2] is not None], key=lambda r: r[2], reverse=True)
        best_avg_state = sorted_by_avg[0][0] if sorted_by_avg else None
        sorted_by_pr = sorted(per_state_rows, key=lambda r: r[4], reverse=True)
        best_pass_state = sorted_by_pr[0][0] if sorted_by_pr else None

        overall_pass = sum(r[3] for r in per_state_rows)
        overall_fail = total_candidates - overall_pass
        overall_pass_rate = round(overall_pass / total_candidates * 100, 2) if total_candidates > 0 else 0.0

        anal.append([])
        anal.append(["Narrative Summary"])
        anal.append([f"Most candidates: {most_state} ({most_cnt})"])
        anal.append([f"Fewest candidates: {least_state} ({least_cnt})"])
        if best_avg_state:
            anal.append([f"Highest average score: {best_avg_state}"])
        if best_pass_state:
            anal.append([f"Highest pass rate: {best_pass_state}"])
        anal.append([f"Overall pass rate: {overall_pass_rate}%"])

    # ---------------------------
    # Candidate batches: registered but absent handling
    # ---------------------------
    candidates_df = load_candidate_batches(CANDIDATE_DIR)
    absent_count = 0
    absent_df = pd.DataFrame(columns=["APPLICATION ID", "FULL NAME", "PHONE NO.", "STATE"])
    if not candidates_df.empty:
        present_ids = set()
        if "APPLICATION ID" in cleaned.columns:
            for v in cleaned["APPLICATION ID"].astype(str).tolist():
                tok = extract_numeric_token(v)
                if tok:
                    present_ids.add(tok)
        for v in cleaned["FULL NAME"].astype(str).tolist():
            tok = extract_numeric_token(v)
            if tok:
                present_ids.add(tok)
        candidates_df["EXAM_NO_NORMAL"] = candidates_df["EXAM_NO"].astype(str).apply(lambda x: extract_numeric_token(x) or x)
        candidate_ids = set(candidates_df["EXAM_NO_NORMAL"].astype(str).tolist())

        absent_ids = sorted([cid for cid in candidate_ids if cid not in present_ids])
        absent_rows = candidates_df[candidates_df["EXAM_NO_NORMAL"].isin(absent_ids)][["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE"]].drop_duplicates()
        absent_df = absent_rows.rename(columns={"EXAM_NO": "APPLICATION ID", "FULL_NAME": "FULL NAME", "PHONE NUMBER": "PHONE NO.", "STATE": "STATE"})
        absent_count = len(absent_df)

        anal.append([])
        anal.append(["Registered (candidate batches) total", int(len(candidates_df))])
        anal.append(["Registered but absent (did not sit)", int(absent_count)])
        anal.append([])

    # ---------------------------
    # Charts sheet
    # ---------------------------
    if "Charts" in wb.sheetnames:
        wb.remove(wb["Charts"])
    chs = wb.create_sheet("Charts")

    # Add chart overview and legend table
    chs["A1"] = "UTME Results Charts Overview"
    chs["A1"].font = Font(bold=True, size=16)
    chs["A2"] = (
        "1. Candidates per State: Number of candidates from each state.\n"
        "2. Average Score per State: Average score (raw) per state.\n"
        "3. Pass Rate per State: Percentage of candidates scoring ≥50 per state.\n"
        "4. Score Distribution: Number of candidates in 10-point score ranges.\n"
        "5. Pass/Fail Distribution: Proportion of all candidates who passed vs. failed."
    )
    chs["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    chs.column_dimensions["A"].width = 40

    # Legend table for state colors
    chs["C2"] = "State Color Legend"
    chs["C2"].font = Font(bold=True, size=12)
    chs["C3"] = "State"
    chs["D3"] = "Color"
    chs["C3"].font = Font(bold=True)
    chs["D3"].font = Font(bold=True)
    row = 4
    state_colors = {}
    for st in sorted(set(df["STATE"].tolist())):
        fill = state_color_for(st)
        state_colors[st] = fill.start_color.rgb[2:]  # Store color without 'FF' prefix
        chs[f"C{row}"] = st
        chs[f"D{row}"].fill = fill
        row += 1
    auto_column_width(chs, min_width=10, max_width=30)

    # Write data for state-based charts
    start_row = row + 2
    chs.cell(row=start_row, column=1, value="State")
    chs.cell(row=start_row, column=2, value="Candidates")
    chs.cell(row=start_row, column=3, value="Average Score")
    chs.cell(row=start_row, column=4, value="Pass Rate (%)")
    counts_sorted = sorted(per_state_rows, key=lambda r: r[1], reverse=True)
    for i, r in enumerate(counts_sorted):
        chs.cell(row=start_row + i + 1, column=1, value=r[0])
        chs.cell(row=start_row + i + 1, column=2, value=r[1])
        chs.cell(row=start_row + i + 1, column=3, value=r[2] if r[2] is not None else 0)
        chs.cell(row=start_row + i + 1, column=4, value=r[4])

    n = len(counts_sorted)
    if n > 0:
        # Chart 1: Candidates per State (Bar)
        # Purpose: Shows how many candidates are from each state, sorted by count.
        try:
            c1 = BarChart()
            c1.title = "Candidates per State"
            c1.x_axis.title = "State"
            c1.y_axis.title = "Number of Candidates"
            c1.width = 32
            c1.height = 16
            c1.gapWidth = 50
            data_ref = Reference(chs, min_col=2, min_row=start_row + 1, max_row=start_row + n)
            cats_ref = Reference(chs, min_col=1, min_row=start_row + 1, max_row=start_row + n)
            c1.add_data(data_ref, titles_from_data=True)
            c1.set_categories(cats_ref)
            c1.dLbls = DataLabelList()
            c1.dLbls.showVal = True
            c1.dLbls.showCatName = False
            c1.style = 10
            for i, series in enumerate(c1.series):
                state = counts_sorted[i][0]
                series.graphicalProperties.solidFill = state_colors[state]
            c1.title.font = Font(size=16, bold=True)
            c1.x_axis.titleFont = Font(size=14)
            c1.y_axis.titleFont = Font(size=14)
            chs.add_chart(c1, f"F{start_row}")
        except Exception as e:
            print(f"⚠️ Error creating Candidates chart: {e}")

        # Chart 2: Average Score per State (Bar)
        # Purpose: Shows the average score for candidates in each state.
        try:
            c2 = BarChart()
            c2.title = "Average Score per State"
            c2.x_axis.title = "State"
            c2.y_axis.title = "Average Score"
            c2.width = 32
            c2.height = 16
            c2.gapWidth = 50
            data_ref2 = Reference(chs, min_col=3, min_row=start_row + 1, max_row=start_row + n)
            cats_ref2 = Reference(chs, min_col=1, min_row=start_row + 1, max_row=start_row + n)
            c2.add_data(data_ref2, titles_from_data=True)
            c2.set_categories(cats_ref2)
            c2.dLbls = DataLabelList()
            c2.dLbls.showVal = True
            c2.dLbls.showCatName = False
            c2.style = 10
            for i, series in enumerate(c2.series):
                state = counts_sorted[i][0]
                series.graphicalProperties.solidFill = state_colors[state]
            c2.title.font = Font(size=16, bold=True)
            c2.x_axis.titleFont = Font(size=14)
            c2.y_axis.titleFont = Font(size=14)
            chs.add_chart(c2, f"F{start_row + 20}")
        except Exception as e:
            print(f"⚠️ Error creating Average Score chart: {e}")

        # Chart 3: Pass Rate per State (Bar)
        # Purpose: Shows the percentage of candidates who scored >=50 in each state.
        try:
            c3 = BarChart()
            c3.title = f"Pass Rate per State (≥ {int(PASS_THRESHOLD)})"
            c3.x_axis.title = "State"
            c3.y_axis.title = "Pass Rate (%)"
            c3.width = 32
            c3.height = 16
            c3.gapWidth = 50
            data_ref3 = Reference(chs, min_col=4, min_row=start_row + 1, max_row=start_row + n)
            cats_ref3 = Reference(chs, min_col=1, min_row=start_row + 1, max_row=start_row + n)
            c3.add_data(data_ref3, titles_from_data=True)
            c3.set_categories(cats_ref3)
            c3.dLbls = DataLabelList()
            c3.dLbls.showVal = True
            c3.dLbls.showCatName = False
            c3.style = 10
            for i, series in enumerate(c3.series):
                state = counts_sorted[i][0]
                series.graphicalProperties.solidFill = state_colors[state]
            c3.title.font = Font(size=16, bold=True)
            c3.x_axis.titleFont = Font(size=14)
            c3.y_axis.titleFont = Font(size=14)
            chs.add_chart(c3, f"F{start_row + 40}")
        except Exception as e:
            print(f"⚠️ Error creating Pass Rate chart: {e}")

    # Chart 4: Score Distribution (Bar)
    # Purpose: Shows how many candidates scored in each 10-point score range.
    try:
        score_vals = score_series.dropna().astype(float)
        if len(score_vals) > 0:
            bins = list(range(0, 101, 10))
            dist = pd.cut(score_vals, bins=bins, right=False)
            freq = dist.value_counts().sort_index()
            insert_row = start_row + n + 3
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
                c4.title = "Score Distribution (10-point Ranges)"
                c4.x_axis.title = "Score Range"
                c4.y_axis.title = "Number of Candidates"
                c4.width = 32
                c4.height = 12
                c4.gapWidth = 50
                data_ref4 = Reference(chs, min_col=2, min_row=insert_row + 1, max_row=insert_row + n_bins)
                cats_ref4 = Reference(chs, min_col=1, min_row=insert_row + 1, max_row=insert_row + n_bins)
                c4.add_data(data_ref4, titles_from_data=True)
                c4.set_categories(cats_ref4)
                c4.dLbls = DataLabelList()
                c4.dLbls.showVal = True
                c4.dLbls.showCatName = False
                c4.style = 10
                for i, series in enumerate(c4.series):
                    series.graphicalProperties.solidFill = PALETTE[i % len(PALETTE)]
                c4.title.font = Font(size=16, bold=True)
                c4.x_axis.titleFont = Font(size=14)
                c4.y_axis.titleFont = Font(size=14)
                chs.add_chart(c4, f"F{insert_row}")
            except Exception as e:
                print(f"⚠️ Error creating Score Distribution chart: {e}")
    except Exception as e:
        print(f"⚠️ Error processing score distribution: {e}")

    # Chart 5: Overall Pass/Fail Pie Chart
    # Purpose: Shows the proportion of all candidates who passed vs. failed.
    try:
        if total_candidates > 0:
            insert_row_pie = chs.max_row + 3
            chs.cell(row=insert_row_pie, column=1, value="Category")
            chs.cell(row=insert_row_pie, column=2, value="Count")
            chs.cell(row=insert_row_pie + 1, column=1, value="Pass")
            chs.cell(row=insert_row_pie + 1, column=2, value=overall_pass)
            chs.cell(row=insert_row_pie + 2, column=1, value="Fail")
            chs.cell(row=insert_row_pie + 2, column=2, value=overall_fail)
            pie = PieChart()
            pie.title = f"Overall Pass/Fail Distribution (≥ {int(PASS_THRESHOLD)})"
            pie.width = 15
            pie.height = 10
            data_ref_pie = Reference(chs, min_col=2, min_row=insert_row_pie + 1, max_row=insert_row_pie + 2)
            cats_ref_pie = Reference(chs, min_col=1, min_row=insert_row_pie + 1, max_row=insert_row_pie + 2)
            pie.add_data(data_ref_pie, titles_from_data=True)
            pie.set_categories(cats_ref_pie)
            pie.dLbls = DataLabelList()
            pie.dLbls.showVal = True
            pie.dLbls.showPercent = True
            pie.dLbls.showCatName = True
            pie.style = 10
            for i, series in enumerate(pie.series):
                series.graphicalProperties.solidFill = "A8E6CF" if i == 0 else "FFCCCB"  # Soft Green for Pass, Soft Red for Fail
            pie.title.font = Font(size=16, bold=True)
            chs.add_chart(pie, f"F{insert_row_pie}")
    except Exception as e:
        print(f"⚠️ Error creating Pass/Fail Pie chart: {e}")

    auto_column_width(chs)
    auto_column_width(anal)

    # Add Absent sheet after Charts
    if not absent_df.empty:
        if "Absent" in wb.sheetnames:
            wb.remove(wb["Absent"])
        abs_ws = wb.create_sheet("Absent")
        abs_ws.append(["APPLICATION ID", "FULL NAME", "PHONE NO.", "STATE"])
        for _, r in absent_df.iterrows():
            abs_ws.append([r.get("APPLICATION ID"), r.get("FULL NAME"), r.get("PHONE NO."), r.get("STATE")])
        auto_column_width(abs_ws)

    # Save workbook
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
    if not candidates_df.empty:
        print(f"  Registered (candidate batches): {len(candidates_df)}")
        print(f"  Registered but absent: {absent_count}")

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

    # Generate timestamp once per run
    ts = datetime.now().strftime(TIMESTAMP_FMT)
    output_dir = os.path.join(CLEAN_DIR, f"UTME_RESULT-{ts}")
    os.makedirs(output_dir, exist_ok=True)

    outputs = []
    for f in files:
        try:
            res = process_file(os.path.join(RAW_DIR, f), output_dir, ts)
            if res:
                outputs.append(res)
        except Exception as e:
            print(f"❌ ERROR processing {f}: {e}", file=sys.stderr)

    print("\n✅ All done. Cleaned files (CSV + XLSX) are in:", output_dir)
    if outputs:
        for csvp, xl in outputs:
            print(" -", csvp)
            print(" -", xl)

if __name__ == "__main__":
    main()