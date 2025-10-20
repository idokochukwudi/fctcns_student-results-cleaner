#!/usr/bin/env python3
"""
utme_result.py

Robust UTME/PUTME cleaning and reporting script (Railway + Local compatible).

Features:
 - Auto-detects Railway vs Local environment
 - Maps common header variations to canonical names
 - Processes results with comprehensive analysis and charts
 - Supports both sorted and unsorted combined results
 - Batch-specific processing and mismatch detection
"""

import os
import sys
import re
import math
import hashlib
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
import argparse

# ---------------------------
# Environment Detection & Directory Configuration
# ---------------------------
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

# UPDATED: Directories now under PUTME_RESULT
DEFAULT_BASE_DIR = os.path.join(BASE_DIR, "PUTME_RESULT")
DEFAULT_RAW_DIR = os.path.join(DEFAULT_BASE_DIR, "RAW_PUTME_RESULT")
DEFAULT_CANDIDATE_DIR = os.path.join(DEFAULT_BASE_DIR, "RAW_CANDIDATE_BATCHES")
DEFAULT_UTME_CANDIDATES_DIR = os.path.join(DEFAULT_BASE_DIR, "RAW_UTME_CANDIDATES")
DEFAULT_CLEAN_DIR = os.path.join(DEFAULT_BASE_DIR, "CLEAN_PUTME_RESULT")
DEFAULT_PASS_THRESHOLD = 50.0
TIMESTAMP_FMT = "%Y-%m-%d_%H:%M:%S"


def parse_args():
    """Parse command-line arguments with defaults."""
    parser = argparse.ArgumentParser(description="UTME/PUTME Results Cleaning Script")
    parser.add_argument(
        "--input-dir",
        default=DEFAULT_RAW_DIR,
        help=f"Input directory for raw files (default: {DEFAULT_RAW_DIR})",
    )
    parser.add_argument(
        "--candidate-dir",
        default=DEFAULT_CANDIDATE_DIR,
        help=f"Directory for candidate batch files (default: {DEFAULT_CANDIDATE_DIR})",
    )
    parser.add_argument(
        "--utme-candidates-dir",
        default=DEFAULT_UTME_CANDIDATES_DIR,
        help=f"Directory for UTME candidates from JAMB (default: {DEFAULT_UTME_CANDIDATES_DIR})",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_CLEAN_DIR,
        help=f"Output directory for cleaned files (default: {DEFAULT_CLEAN_DIR})",
    )
    parser.add_argument(
        "--pass-threshold",
        type=float,
        default=DEFAULT_PASS_THRESHOLD,
        help=f"Pass threshold for highlighting (default: {DEFAULT_PASS_THRESHOLD})",
    )
    parser.add_argument(
        "--batch-id",
        help="Specific batch ID to process (e.g., 'Batch1A' or 'Batch 1A')",
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="Skip interactive prompts (e.g., for converted score)",
    )
    parser.add_argument(
        "--converted-score-max",
        type=int,
        help="Target maximum for converted score (e.g., 60 for Score/60%) when non-interactive",
    )
    return parser.parse_args()


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
    "D4A5A5",  # Soft Rose
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
    norm_map = {
        col: re.sub(r"\s+", " ", str(col).strip().lower()) for col in df.columns
    }
    candidates = [re.sub(r"\s+", " ", c.strip().lower()) for c in candidate_names]
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
    return pd.to_numeric(
        series.astype(str).str.replace(",", "").str.strip(), errors="coerce"
    )


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
    mask = df.apply(
        lambda row: row.astype(str)
        .str.contains("overall average", case=False, na=False)
        .any(),
        axis=1,
    )
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
        ws.column_dimensions[get_column_letter(i)].width = min(
            max_width, max(min_width, max_len + 2)
        )


def normalize_id(s):
    if s is None:
        return None
    return str(s).strip().lower()


# ---------------------------
# Candidate batches helpers
# ---------------------------
def load_candidate_batches(folder, batch_id=None):
    """Load candidate-batch files. If batch_id is provided, load only the matching batch file; otherwise, load all with batch IDs."""
    if not os.path.isdir(folder):
        print(f"Warning: Candidate batch directory {folder} does not exist.")
        return (
            pd.DataFrame(
                columns=["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE", "BATCH_ID"]
            ),
            {},
            {},
        )

    files = [
        f
        for f in os.listdir(folder)
        if f.lower().endswith((".csv", ".xlsx", ".xls")) and not f.startswith("~$")
    ]
    if not files:
        print(f"Warning: No candidate batch files found in {folder}.")
        return (
            pd.DataFrame(
                columns=["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE", "BATCH_ID"]
            ),
            {},
            {},
        )

    batch_id_map = {}  # Maps normalized ID to registered BATCH_ID
    numeric_to_full = {}  # Maps numeric token to original EXAM_NO
    rows = []
    for fname in sorted(files):
        path = os.path.join(folder, fname)
        batch_id_match = re.search(r"(Batch\s*\d+[A-Za-z]?)", fname, re.IGNORECASE)
        current_batch_id = batch_id_match.group(1) if batch_id_match else fname
        if batch_id:
            batch_id_normalized = re.sub(r"\s+", "", batch_id.lower())
            if batch_id_normalized not in re.sub(r"\s+", "", fname.lower()):
                continue
        print(f"Loading candidate batch file: {fname} (Batch ID: {current_batch_id})")
        try:
            if fname.lower().endswith(".csv"):
                cdf = pd.read_csv(path, dtype=str)
            else:
                cdf = pd.read_excel(path, dtype=str)
        except Exception as e:
            print(f"Error reading candidate batch {fname}: {e}")
            continue
        exam_col = find_column_by_names(
            cdf, ["username", "exam no", "reg no", "mat no", "regnum", "reg number"]
        )
        name_col = find_column_by_names(
            cdf,
            [
                "firstname",
                "full name",
                "name",
                "candidate name",
                "user full name",
                "RG_CANDNAME",
            ],
        )
        phone_col = find_column_by_names(
            cdf, ["phone1", "phone", "phone number", "PHONE", "Mobile", "PhoneNumber"]
        )
        state_col = find_column_by_names(
            cdf,
            [
                "city",
                "city/town",
                "City /town",
                "City",
                "Town",
                "State of Origin",
                "STATE",
            ],
        )
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
                    if (
                        not name
                        and re.search(r"[A-Za-z]{2,}", s)
                        and not re.fullmatch(r"\d+", s)
                    ):
                        name = s
                    if not phone and re.search(r"\d{10,}", s):
                        phone = s
                    if (
                        not state
                        and re.search(r"[A-Za-z]{2,}", s)
                        and c.lower() in ["city", "state"]
                    ):
                        state = s
            if ex:
                rows.append(
                    {
                        "EXAM_NO": ex,
                        "FULL_NAME": name or "",
                        "PHONE NUMBER": phone or "",
                        "STATE": state or "",
                        "BATCH_ID": current_batch_id,
                    }
                )
                norm_id = normalize_id(ex)
                batch_id_map[norm_id] = current_batch_id
                tok = extract_numeric_token(ex)
                if tok:
                    numeric_to_full[tok] = ex
    if not rows:
        print(
            f"No valid candidate records found in batch files{' for ' + batch_id if batch_id else ''}."
        )
        return (
            pd.DataFrame(
                columns=["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE", "BATCH_ID"]
            ),
            {},
            {},
        )

    cdf_all = pd.DataFrame(rows)
    duplicates = cdf_all[cdf_all.duplicated(subset=["EXAM_NO"], keep=False)]
    if not duplicates.empty:
        print(
            f"Found {len(duplicates)} duplicate EXAM_NO entries in candidate batches{' for ' + batch_id if batch_id else ''}:"
        )
        print(duplicates[["EXAM_NO", "FULL_NAME", "BATCH_ID"]].head().to_string())
    cdf_all = cdf_all.drop_duplicates(subset=["EXAM_NO"], keep="first")
    print(
        f"Loaded {len(cdf_all)} unique registered candidates from batch files{' for ' + batch_id if batch_id else ''}."
    )
    if "PHONE NUMBER" in cdf_all.columns:
        cdf_all["PHONE NUMBER"] = cdf_all["PHONE NUMBER"].apply(clean_phone_value)
    return cdf_all, batch_id_map, numeric_to_full


# ---------------------------
# Load UTME Candidates from JAMB
# ---------------------------
def load_utme_candidates_from_jamb(folder):
    """Load all UTME candidates who chose the college from JAMB data."""
    if not os.path.isdir(folder):
        print(f"Warning: UTME candidates directory {folder} does not exist.")
        return pd.DataFrame(columns=["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE"])

    files = [
        f
        for f in os.listdir(folder)
        if f.lower().endswith((".csv", ".xlsx", ".xls")) and not f.startswith("~$")
    ]
    
    if not files:
        print(f"Warning: No UTME candidate files found in {folder}.")
        return pd.DataFrame(columns=["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE"])

    all_utme_rows = []
    for fname in sorted(files):
        path = os.path.join(folder, fname)
        print(f"Loading UTME candidates from JAMB: {fname}")
        try:
            if fname.lower().endswith(".csv"):
                udf = pd.read_csv(path, dtype=str)
            else:
                udf = pd.read_excel(path, dtype=str)
        except Exception as e:
            print(f"Error reading UTME candidates {fname}: {e}")
            continue

        # Try to find relevant columns (similar to candidate batch loading)
        exam_col = find_column_by_names(
            udf, ["username", "exam no", "reg no", "mat no", "regnum", "reg number", "RG_NUM"]
        )
        name_col = find_column_by_names(
            udf,
            [
                "firstname",
                "full name",
                "name",
                "candidate name",
                "user full name",
                "RG_CANDNAME",
            ],
        )
        phone_col = find_column_by_names(
            udf, ["phone1", "phone", "phone number", "PHONE", "Mobile", "PhoneNumber"]
        )
        state_col = find_column_by_names(
            udf,
            [
                "city",
                "city/town",
                "City /town",
                "City",
                "Town",
                "State of Origin",
                "STATE",
                "STATE_NAME",
            ],
        )

        for _, r in udf.iterrows():
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

            if ex:
                all_utme_rows.append(
                    {
                        "EXAM_NO": ex,
                        "FULL_NAME": name or "",
                        "PHONE NUMBER": phone or "",
                        "STATE": state or "",
                    }
                )

    if not all_utme_rows:
        print("No valid UTME candidate records found in JAMB files.")
        return pd.DataFrame(columns=["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE"])

    utme_df = pd.DataFrame(all_utme_rows)
    utme_df = utme_df.drop_duplicates(subset=["EXAM_NO"], keep="first")
    
    if "PHONE NUMBER" in utme_df.columns:
        utme_df["PHONE NUMBER"] = utme_df["PHONE NUMBER"].apply(clean_phone_value)
    
    print(f"Loaded {len(utme_df)} unique UTME candidates from JAMB data.")
    return utme_df


# ---------------------------
# Find candidates who didn't apply
# ---------------------------
def find_non_applicants(utme_df, applied_df):
    """
    Compare UTME candidates (chose college) with applied candidates.
    Returns candidates who chose the college but didn't apply.
    """
    if utme_df.empty:
        print("No UTME candidates data available for comparison.")
        return pd.DataFrame(columns=["EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE"])
    
    if applied_df.empty:
        print("No applied candidates data available for comparison.")
        return utme_df.copy()

    # Normalize EXAM_NO for comparison
    utme_df["EXAM_NO_NORM"] = utme_df["EXAM_NO"].apply(
        lambda x: extract_numeric_token(x) or normalize_id(x)
    )
    applied_df["EXAM_NO_NORM"] = applied_df["EXAM_NO"].apply(
        lambda x: extract_numeric_token(x) or normalize_id(x)
    )

    # Find candidates in UTME but not in applied
    applied_ids = set(applied_df["EXAM_NO_NORM"])
    non_applicants = utme_df[~utme_df["EXAM_NO_NORM"].isin(applied_ids)].copy()
    
    # Drop the normalized column before returning
    non_applicants = non_applicants.drop(columns=["EXAM_NO_NORM"])
    
    print(f"\nCandidate Comparison:")
    print(f"  Total UTME candidates (chose college): {len(utme_df)}")
    print(f"  Total applied candidates: {len(applied_df)}")
    print(f"  Candidates who chose but didn't apply: {len(non_applicants)}")
    
    return non_applicants


def extract_numeric_token(s):
    """Return first sequence of 3+ digits found in string, else None."""
    if s is None:
        return None
    s = re.sub(r"\W+", "", str(s).strip().lower())
    m = re.search(r"(\d{3,})", s)
    return m.group(1) if m else None


# ---------------------------
# Format Excel sheet helper
# ---------------------------
def format_excel_sheet(
    wb,
    ws_name,
    cleaned,
    state_colors,
    score_header,
    pass_threshold,
    apply_state_colors=True,
    highlight_passing_scores=True,
    title_row=None,
):
    ws = wb[ws_name]
    header_row = 1 if title_row is None else 2
    header_font = Font(bold=True, size=12)
    for cell in ws[header_row]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Apply title row formatting if provided
    if title_row is not None:
        for cell in ws[1]:
            cell.font = Font(bold=True, size=14)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(
            start_row=1, start_column=1, end_row=1, end_column=len(cleaned.columns)
        )
        ws.cell(row=1, column=1).value = title_row

    if "APPLICATION ID" in cleaned.columns:
        try:
            col_idx = list(cleaned.columns).index("APPLICATION ID") + 1
            for r in ws.iter_rows(
                min_row=header_row + 1,
                min_col=col_idx,
                max_col=col_idx,
                max_row=ws.max_row,
            ):
                for cell in r:
                    cell.alignment = Alignment(horizontal="left")
        except Exception:
            pass

    ws.freeze_panes = f"A{header_row + 1}"
    auto_column_width(ws)

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in ws.iter_rows(
        min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
    ):
        for cell in r:
            cell.border = border
            # Apply bold formatting to "Overall Average" row
            if (
                cell.row > header_row
                and ws.cell(row=cell.row, column=1).value == "Overall Average"
            ):
                cell.font = Font(bold=True)

    if apply_state_colors:
        try:
            state_col_index = list(cleaned.columns).index("STATE") + 1
            for row_idx in range(header_row + 1, ws.max_row + 1):
                state_val = ws.cell(row=row_idx, column=state_col_index).value
                # Skip coloring for "Overall Average" row
                if ws.cell(row=row_idx, column=1).value == "Overall Average":
                    continue
                fill = state_color_for(state_val)
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = fill
        except Exception:
            pass

    if highlight_passing_scores:
        score_cols_indices = [
            i + 1 for i, c in enumerate(cleaned.columns) if str(c).startswith("Score/")
        ]
        pass_fill = PatternFill(
            start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
        )
        pass_font = Font(color="006100")
        for col_idx in score_cols_indices:
            for row_idx in range(header_row + 1, ws.max_row + 1):
                # Skip highlighting for "Overall Average" row
                if ws.cell(row=row_idx, column=1).value == "Overall Average":
                    continue
                cell = ws.cell(row=row_idx, column=col_idx)
                try:
                    if cell.value is None or str(cell.value).strip() == "":
                        continue
                    val = float(cell.value)
                    if val >= pass_threshold:
                        cell.fill = pass_fill
                        cell.font = pass_font
                except Exception:
                    continue


# ---------------------------
# Analysis and Charts
# ---------------------------
def create_analysis_and_charts(wb, cleaned, score_header, pass_threshold, title=None):
    """Create analysis sheet with statistics and charts."""
    ws_analysis = wb.create_sheet("Analysis")
    
    # Title
    if title:
        ws_analysis.append([title])
        ws_analysis.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        title_cell = ws_analysis.cell(row=1, column=1)
        title_cell.font = Font(bold=True, size=16)
        title_cell.alignment = Alignment(horizontal="center")
        ws_analysis.append([])  # Empty row

    # Basic statistics
    scores = to_numeric_safe(cleaned[score_header]).dropna()
    if len(scores) == 0:
        ws_analysis.append(["No valid scores found for analysis"])
        return

    stats = [
        ["Statistic", "Value"],
        ["Total Candidates", len(cleaned)],
        ["Valid Scores", len(scores)],
        ["Average Score", f"{scores.mean():.2f}"],
        ["Highest Score", f"{scores.max():.2f}"],
        ["Lowest Score", f"{scores.min():.2f}"],
        ["Pass Threshold", f"{pass_threshold:.1f}"],
        ["Passed", f"{(scores >= pass_threshold).sum()}"],
        ["Failed", f"{(scores < pass_threshold).sum()}"],
        ["Pass Rate", f"{(scores >= pass_threshold).mean() * 100:.1f}%"],
    ]

    for row in stats:
        ws_analysis.append(row)

    # Score distribution
    ws_analysis.append([])
    ws_analysis.append(["Score Distribution"])
    bins = [0, pass_threshold, 100]
    labels = [f"Below {pass_threshold}", f"{pass_threshold} and above"]
    distribution = pd.cut(scores, bins=bins, labels=labels, right=False).value_counts()
    
    for label in labels:
        count = distribution.get(label, 0)
        ws_analysis.append([f"{label}", count])

    # State-wise performance (if STATE column exists)
    if "STATE" in cleaned.columns:
        ws_analysis.append([])
        ws_analysis.append(["State-wise Performance"])
        state_scores = cleaned.groupby("STATE")[score_header].agg(['count', 'mean']).round(2)
        state_scores = state_scores.sort_values('mean', ascending=False)
        
        ws_analysis.append(["State", "Count", "Average Score"])
        for state, row in state_scores.iterrows():
            ws_analysis.append([state, row['count'], row['mean']])

    # Create charts
    row_offset = len(stats) + 5  # Adjust based on content
    
    # Pie chart for pass/fail distribution
    pie_chart = PieChart()
    labels = Reference(ws_analysis, min_col=1, min_row=row_offset, max_row=row_offset+1)
    data = Reference(ws_analysis, min_col=2, min_row=row_offset-1, max_row=row_offset+1)
    pie_chart.add_data(data, titles_from_data=True)
    pie_chart.set_categories(labels)
    pie_chart.title = "Pass/Fail Distribution"
    pie_chart.dataLabels = DataLabelList()
    pie_chart.dataLabels.showPercent = True
    ws_analysis.add_chart(pie_chart, "E2")

    # Bar chart for state performance (if available)
    if "STATE" in cleaned.columns and len(state_scores) > 0:
        bar_chart = BarChart()
        state_labels = Reference(ws_analysis, min_col=1, min_row=row_offset+4, max_row=row_offset+3+len(state_scores))
        state_data = Reference(ws_analysis, min_col=3, min_row=row_offset+3, max_row=row_offset+3+len(state_scores))
        bar_chart.add_data(state_data, titles_from_data=True)
        bar_chart.set_categories(state_labels)
        bar_chart.title = "Average Score by State"
        bar_chart.y_axis.title = "Average Score"
        bar_chart.x_axis.title = "State"
        ws_analysis.add_chart(bar_chart, "E16")

    auto_column_width(ws_analysis)


# ---------------------------
# File Processing
# ---------------------------
def process_file(
    file_path,
    output_dir,
    timestamp,
    candidate_dir,
    full_candidates_df,
    full_batch_id_map,
    full_numeric_to_full,
    pass_threshold,
):
    """Process a single result file."""
    print(f"\n📄 Processing: {os.path.basename(file_path)}")
    
    try:
        if file_path.lower().endswith(".csv"):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return None, None, None, None

    # Remove overall average rows
    df = drop_overall_average_rows(df)

    # Find key columns
    exam_col = find_column_by_names(
        df, ["username", "exam no", "reg no", "mat no", "regnum", "reg number"]
    )
    name_col = find_column_by_names(
        df, ["firstname", "full name", "name", "candidate name", "user full name"]
    )
    score_col = find_grade_column(df)

    if not exam_col or not score_col:
        print(f"❌ Could not find required columns in {os.path.basename(file_path)}")
        print(f"   Available columns: {list(df.columns)}")
        return None, None, None, None

    print(f"   Using columns: EXAM={exam_col}, NAME={name_col}, SCORE={score_col}")

    # Extract exam numbers and scores
    exam_nos = df[exam_col].astype(str).str.strip()
    scores = to_numeric_safe(df[score_col])
    names = df[name_col].astype(str).str.strip() if name_col else None

    # Create cleaned dataframe
    cleaned_data = []
    for idx, (exam_no, score) in enumerate(zip(exam_nos, scores)):
        name = names.iloc[idx] if names is not None else ""
        cleaned_data.append({
            "EXAM_NO": exam_no,
            "FULL_NAME": name,
            "SCORE": score
        })

    cleaned = pd.DataFrame(cleaned_data)
    
    # Add candidate information
    cleaned = add_candidate_info(
        cleaned, full_candidates_df, full_batch_id_map, full_numeric_to_full
    )

    # Rename score column
    max_score = 100  # Default, can be adjusted
    if not cleaned.empty and cleaned["SCORE"].max() > 100:
        max_score = cleaned["SCORE"].max()
    score_header = f"Score/{max_score}%"
    cleaned = cleaned.rename(columns={"SCORE": score_header})

    # Sort by score descending
    cleaned = cleaned.sort_values(by=score_header, ascending=False)
    
    # Add rank
    cleaned["RANK"] = range(1, len(cleaned) + 1)

    # Reorder columns
    col_order = ["RANK", "EXAM_NO", "FULL_NAME", "PHONE NUMBER", "STATE", "BATCH_ID", score_header]
    existing_cols = [col for col in col_order if col in cleaned.columns]
    other_cols = [col for col in cleaned.columns if col not in col_order]
    cleaned = cleaned[existing_cols + other_cols]

    # Save files
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    csv_path = os.path.join(output_dir, f"UTME_RESULT_{base_name}_{timestamp}.csv")
    excel_path = os.path.join(output_dir, f"UTME_RESULT_{base_name}_{timestamp}.xlsx")

    # Save CSV
    cleaned.to_csv(csv_path, index=False)
    
    # Save Excel with formatting
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    
    for r in dataframe_to_rows(cleaned, index=False, header=True):
        ws.append(r)

    # Create state colors mapping
    state_colors = {}
    if "STATE" in cleaned.columns:
        for state in cleaned["STATE"].unique():
            state_colors[state] = state_color_for(state)

    format_excel_sheet(wb, "Results", cleaned, state_colors, score_header, pass_threshold)
    create_analysis_and_charts(wb, cleaned, score_header, pass_threshold, f"UTME Results - {base_name}")
    
    wb.save(excel_path)

    print(f"✅ Processed: {os.path.basename(file_path)}")
    print(f"   Candidates: {len(cleaned)}")
    print(f"   Saved: {os.path.basename(csv_path)}")
    print(f"   Saved: {os.path.basename(excel_path)}")

    # Find absent and mismatch candidates
    absent_df, mismatch_df = find_absent_and_mismatch(
        cleaned, full_candidates_df, full_batch_id_map
    )

    return cleaned, df, absent_df, mismatch_df


def add_candidate_info(cleaned_df, candidates_df, batch_id_map, numeric_to_full):
    """Add candidate information from batch files."""
    if cleaned_df.empty:
        return cleaned_df

    result = cleaned_df.copy()
    
    # Initialize new columns
    result["PHONE NUMBER"] = ""
    result["STATE"] = ""
    result["BATCH_ID"] = ""
    result["FULL_NAME_CANDIDATE"] = ""

    for idx, row in result.iterrows():
        exam_no = row["EXAM_NO"]
        norm_id = normalize_id(exam_no)
        numeric_tok = extract_numeric_token(exam_no)

        # Try to find candidate in batch files
        candidate_info = None
        
        # First try exact match
        if not candidate_info and norm_id in batch_id_map:
            candidate_match = candidates_df[candidates_df["EXAM_NO"].apply(normalize_id) == norm_id]
            if not candidate_match.empty:
                candidate_info = candidate_match.iloc[0]

        # Then try numeric token match
        if not candidate_info and numeric_tok and numeric_tok in numeric_to_full:
            full_exam_no = numeric_to_full[numeric_tok]
            candidate_match = candidates_df[candidates_df["EXAM_NO"] == full_exam_no]
            if not candidate_match.empty:
                candidate_info = candidate_match.iloc[0]

        # Then try partial match in candidate database
        if not candidate_info:
            for _, cand_row in candidates_df.iterrows():
                cand_exam = str(cand_row["EXAM_NO"]).strip()
                if norm_id in normalize_id(cand_exam) or (numeric_tok and numeric_tok in cand_exam):
                    candidate_info = cand_row
                    break

        # Update with candidate information
        if candidate_info is not None:
            if pd.notna(candidate_info.get("PHONE NUMBER")):
                result.at[idx, "PHONE NUMBER"] = clean_phone_value(candidate_info["PHONE NUMBER"])
            if pd.notna(candidate_info.get("STATE")):
                result.at[idx, "STATE"] = candidate_info["STATE"]
            if pd.notna(candidate_info.get("BATCH_ID")):
                result.at[idx, "BATCH_ID"] = candidate_info["BATCH_ID"]
            if pd.notna(candidate_info.get("FULL_NAME")):
                result.at[idx, "FULL_NAME_CANDIDATE"] = candidate_info["FULL_NAME"]

    return result


def find_absent_and_mismatch(cleaned_df, candidates_df, batch_id_map):
    """Find absent candidates and name mismatches."""
    absent_candidates = []
    mismatch_candidates = []

    if cleaned_df.empty or candidates_df.empty:
        return pd.DataFrame(absent_candidates), pd.DataFrame(mismatch_candidates)

    # Find absent candidates (registered but not in results)
    cleaned_exam_nos = set(cleaned_df["EXAM_NO"].apply(normalize_id))
    
    for _, cand_row in candidates_df.iterrows():
        exam_no = cand_row["EXAM_NO"]
        norm_id = normalize_id(exam_no)
        
        if norm_id not in cleaned_exam_nos:
            # Check if this might be a numeric token match
            numeric_tok = extract_numeric_token(exam_no)
            found = False
            if numeric_tok:
                for cleaned_exam in cleaned_df["EXAM_NO"]:
                    if numeric_tok in str(cleaned_exam):
                        found = True
                        break
            
            if not found:
                absent_candidates.append({
                    "EXAM_NO": exam_no,
                    "FULL_NAME": cand_row.get("FULL_NAME", ""),
                    "PHONE NUMBER": cand_row.get("PHONE NUMBER", ""),
                    "STATE": cand_row.get("STATE", ""),
                    "BATCH_ID": cand_row.get("BATCH_ID", "")
                })

    # Find name mismatches
    for idx, result_row in cleaned_df.iterrows():
        exam_no = result_row["EXAM_NO"]
        result_name = result_row.get("FULL_NAME", "").strip().lower()
        
        # Find candidate in batch files
        candidate_match = None
        norm_id = normalize_id(exam_no)
        
        if norm_id in batch_id_map:
            candidate_match = candidates_df[candidates_df["EXAM_NO"].apply(normalize_id) == norm_id]
        
        if candidate_match is not None and not candidate_match.empty:
            cand_name = candidate_match.iloc[0].get("FULL_NAME", "").strip().lower()
            if (result_name and cand_name and 
                result_name != cand_name and 
                not is_similar_name(result_name, cand_name)):
                mismatch_candidates.append({
                    "EXAM_NO": exam_no,
                    "RESULT_NAME": result_row.get("FULL_NAME", ""),
                    "CANDIDATE_NAME": candidate_match.iloc[0].get("FULL_NAME", ""),
                    "PHONE NUMBER": candidate_match.iloc[0].get("PHONE NUMBER", ""),
                    "STATE": candidate_match.iloc[0].get("STATE", ""),
                    "BATCH_ID": candidate_match.iloc[0].get("BATCH_ID", "")
                })

    return pd.DataFrame(absent_candidates), pd.DataFrame(mismatch_candidates)


def is_similar_name(name1, name2):
    """Check if two names are similar despite minor differences."""
    name1 = re.sub(r'\s+', ' ', name1.strip().lower())
    name2 = re.sub(r'\s+', ' ', name2.strip().lower())
    
    if name1 == name2:
        return True
    
    # Split into words and check if one is subset of another
    words1 = set(name1.split())
    words2 = set(name2.split())
    
    # If one name contains all words of the other, consider similar
    if words1.issubset(words2) or words2.issubset(words1):
        return True
    
    return False


def dataframe_to_rows(df, index=True, header=True):
    """Convert DataFrame to rows for openpyxl."""
    if header:
        yield df.columns.tolist()
    for _, row in df.iterrows():
        yield row.tolist()


# ---------------------------
# Batch Combination
# ---------------------------
def combine_batches(
    cleaned_dfs,
    raw_dfs,
    absent_dfs,
    mismatch_dfs,
    output_dir,
    timestamp,
    pass_threshold,
    candidate_dir,
    raw_dir,
    non_interactive=False,
    converted_score_max=None
):
    """Combine all batches into single files."""
    if not cleaned_dfs:
        print("❌ No data to combine")
        return

    print(f"\n🔄 Combining {len(cleaned_dfs)} batches...")

    # Combine all cleaned dataframes
    combined = pd.concat(cleaned_dfs, ignore_index=True)
    
    # Remove duplicates based on EXAM_NO, keeping the highest score
    combined = combined.sort_values(by=[col for col in combined.columns if col.startswith("Score/")][0], ascending=False)
    combined = combined.drop_duplicates(subset=["EXAM_NO"], keep="first")

    # Add overall rank
    score_col = [col for col in combined.columns if col.startswith("Score/")][0]
    combined = combined.sort_values(by=score_col, ascending=False)
    combined["OVERALL_RANK"] = range(1, len(combined) + 1)

    # Reorder columns to put overall rank first
    cols = combined.columns.tolist()
    cols.remove("OVERALL_RANK")
    combined = combined[["OVERALL_RANK"] + cols]

    # Save sorted combined results
    save_combined_results(combined, output_dir, timestamp, pass_threshold, "SORTED")
    
    # Save unsorted combined results (original order)
    save_combined_results(combined, output_dir, timestamp, pass_threshold, "UNSORTED", sorted=False)

    # Save absent and mismatch reports
    save_absent_mismatch_reports(absent_dfs, mismatch_dfs, output_dir, timestamp)

    print(f"✅ Combined processing completed")
    print(f"   Total unique candidates: {len(combined)}")


def save_combined_results(combined_df, output_dir, timestamp, pass_threshold, suffix, sorted=True):
    """Save combined results with proper formatting."""
    if sorted:
        score_col = [col for col in combined_df.columns if col.startswith("Score/")][0]
        result_df = combined_df.sort_values(by=score_col, ascending=False)
        filename = f"PUTME_COMBINE_RESULT_{timestamp}.xlsx"
        sheet_name = "Combined Results"
    else:
        result_df = combined_df
        filename = f"PUTME_COMBINE_RESULT_UNSORTED_{timestamp}.xlsx"
        sheet_name = "Combined Results Unsorted"

    filepath = os.path.join(output_dir, filename)
    
    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Add data
    for r in dataframe_to_rows(result_df, index=False, header=True):
        ws.append(r)

    # Apply formatting
    score_header = [col for col in result_df.columns if col.startswith("Score/")][0]
    state_colors = {}
    if "STATE" in result_df.columns:
        for state in result_df["STATE"].unique():
            state_colors[state] = state_color_for(state)

    format_excel_sheet(wb, sheet_name, result_df, state_colors, score_header, pass_threshold)
    
    # Add analysis
    title = f"Combined UTME Results - {timestamp}"
    create_analysis_and_charts(wb, result_df, score_header, pass_threshold, title)
    
    wb.save(filepath)
    print(f"✅ Saved: {filename}")


def save_absent_mismatch_reports(absent_dfs, mismatch_dfs, output_dir, timestamp):
    """Save absent and mismatch candidate reports."""
    # Combine all absent dataframes
    if absent_dfs:
        all_absent = pd.concat(absent_dfs, ignore_index=True)
        all_absent = all_absent.drop_duplicates(subset=["EXAM_NO"])
        if not all_absent.empty:
            absent_file = os.path.join(output_dir, f"ABSENT_CANDIDATES_{timestamp}.xlsx")
            all_absent.to_excel(absent_file, index=False)
            print(f"✅ Saved: ABSENT_CANDIDATES_{timestamp}.xlsx ({len(all_absent)} candidates)")

    # Combine all mismatch dataframes
    if mismatch_dfs:
        all_mismatch = pd.concat(mismatch_dfs, ignore_index=True)
        all_mismatch = all_mismatch.drop_duplicates(subset=["EXAM_NO"])
        if not all_mismatch.empty:
            mismatch_file = os.path.join(output_dir, f"NAME_MISMATCHES_{timestamp}.xlsx")
            all_mismatch.to_excel(mismatch_file, index=False)
            print(f"✅ Saved: NAME_MISMATCHES_{timestamp}.xlsx ({len(all_mismatch)} candidates)")


# ---------------------------
# Entrypoint
# ---------------------------
def main():
    args = parse_args()
    print("\nStarting UTME Results Cleaning...")
    print(f"📁 Input Directory: {args.input_dir}")
    print(f"📁 Candidate Directory: {args.candidate_dir}")
    print(f"📁 UTME Candidates Directory: {args.utme_candidates_dir}")
    print(f"📁 Output Directory: {args.output_dir}\n")

    # Use command-line arguments or defaults
    RAW_DIR = args.input_dir
    CANDIDATE_DIR = args.candidate_dir
    UTME_CANDIDATES_DIR = args.utme_candidates_dir
    CLEAN_DIR = args.output_dir
    PASS_THRESHOLD = args.pass_threshold
    BATCH_ID = args.batch_id
    NON_INTERACTIVE = args.non_interactive
    CONVERTED_SCORE_MAX = args.converted_score_max

    # Ensure directories exist
    os.makedirs(RAW_DIR, exist_ok=True)
    os.makedirs(CLEAN_DIR, exist_ok=True)
    os.makedirs(CANDIDATE_DIR, exist_ok=True)
    os.makedirs(UTME_CANDIDATES_DIR, exist_ok=True)

    # Load UTME candidates from JAMB
    utme_candidates_df = load_utme_candidates_from_jamb(UTME_CANDIDATES_DIR)
    
    # Load full candidates_df (applied candidates)
    full_candidates_df, full_batch_id_map, full_numeric_to_full = (
        load_candidate_batches(CANDIDATE_DIR)
    )
    
    # Find candidates who chose college but didn't apply
    non_applicants_df = find_non_applicants(utme_candidates_df, full_candidates_df)
    
    # Save non-applicants to a separate file
    if not non_applicants_df.empty:
        ts = datetime.now().strftime(TIMESTAMP_FMT)
        non_app_file = os.path.join(CLEAN_DIR, f"CHOSE_BUT_DID_NOT_APPLY_{ts}.xlsx")
        try:
            non_applicants_df.to_excel(non_app_file, index=False, engine="openpyxl")
            print(f"\n✅ Saved non-applicants list: {os.path.basename(non_app_file)}")
            print(f"   {len(non_applicants_df)} candidates chose the college but didn't apply\n")
        except Exception as e:
            print(f"⚠️  Error saving non-applicants file: {e}")

    # Filter files based on batch_id if provided
    files = [
        f for f in os.listdir(RAW_DIR) if f.lower().endswith((".csv", ".xlsx", ".xls"))
    ]
    if BATCH_ID:
        batch_id_normalized = re.sub(r"\s+", "", BATCH_ID.lower())
        files = [
            f for f in files if batch_id_normalized in re.sub(r"\s+", "", f.lower())
        ]
        if not files:
            print(f"❌ No raw files found matching batch ID '{BATCH_ID}' in {RAW_DIR}")
            return

    if not files:
        print(
            f"❌ No raw files found in {RAW_DIR}\nPut your raw file(s) there and re-run."
        )
        return

    print(f"📊 Found {len(files)} file(s) to process\n")

    ts = datetime.now().strftime(TIMESTAMP_FMT)
    output_dir = os.path.join(CLEAN_DIR, f"UTME_RESULT-{ts}")
    os.makedirs(output_dir, exist_ok=True)

    outputs = []
    cleaned_dfs = []
    dfs = []
    absent_dfs = []
    mismatch_dfs = []

    for f in files:
        try:
            cleaned, df, absent_df, mismatch_df = process_file(
                os.path.join(RAW_DIR, f),
                output_dir,
                ts,
                CANDIDATE_DIR,
                full_candidates_df,
                full_batch_id_map,
                full_numeric_to_full,
                PASS_THRESHOLD,
            )
            if cleaned is not None:
                outputs.append(
                    (
                        os.path.join(
                            output_dir, f"UTME_RESULT_{os.path.splitext(f)[0]}_{ts}.csv"
                        ),
                        os.path.join(
                            output_dir,
                            f"UTME_RESULT_{os.path.splitext(f)[0]}_{ts}.xlsx",
                        ),
                    )
                )
                cleaned_dfs.append(cleaned)
                dfs.append(df)
                absent_dfs.append(absent_df)
                mismatch_dfs.append(mismatch_df)
        except Exception as e:
            print(f"⚠️  Error processing {f}: {e}", file=sys.stderr)

    combine_batches(
        cleaned_dfs,
        dfs,
        absent_dfs,
        mismatch_dfs,
        output_dir,
        ts,
        PASS_THRESHOLD,
        CANDIDATE_DIR,
        RAW_DIR,
        NON_INTERACTIVE,
        CONVERTED_SCORE_MAX,
    )

    print("\n✅ Processing completed successfully!")
    print(f"📂 Results saved in: {output_dir}")
    if outputs:
        for csvp, xl in outputs:
            print(f"   - {os.path.basename(csvp)}")
            print(f"   - {os.path.basename(xl)}")
    print(f"   - PUTME_COMBINE_RESULT_{ts}.xlsx")
    print(f"   - PUTME_COMBINE_RESULT_UNSORTED_{ts}.xlsx")
    if not non_applicants_df.empty:
        print(f"\n📋 Non-applicants report:")
        print(f"   - CHOSE_BUT_DID_NOT_APPLY_{ts}.xlsx")
        print(f"   - {len(non_applicants_df)} candidates identified")


if __name__ == "__main__":
    main()