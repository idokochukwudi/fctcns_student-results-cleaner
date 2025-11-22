#!/usr/bin/env python3 
"""
caosce_result_fixed.py
FULLY FIXED & ENHANCED CAOSCE cleaning script (November 2025 revision)
- Dynamically detects Grade/X column names instead of hardcoded Grade/10
- Properly fits cells to maximum content length with autofit
- Ensures Date is directly under CAOSCE_2025 (centered)
- Forces MAT NO. and FULL NAME headers and cells to be LEFT aligned
- Removes any duplicate OVERALL AVERAGE rows (keaks a single computed overall average)
- Summary, Analysis and Signatories in separate, well-structured section
- Professional formatting with optimal column widths
- Fixed CLASS alignment to not interfere with column fitting
- Increased header font sizes and logo size
- Properly aligned Approved by section
- Removed detailed examination structure and performance categories
"""

import os
import re
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# ---------------------------
# Configuration
# ---------------------------
IS_RAILWAY = os.getenv("RAILWAY_ENVIRONMENT") is not None

if IS_RAILWAY:
    BASE_DIR = os.getenv("BASE_DIR", "/app/EXAMS_INTERNAL")
else:
    BASE_DIR = os.path.join(os.path.expanduser("~"), "student_result_cleaner", "EXAMS_INTERNAL")

DEFAULT_BASE_DIR = os.path.join(BASE_DIR, "CAOSCE_RESULT")
DEFAULT_RAW_DIR = os.path.join(DEFAULT_BASE_DIR, "RAW_CAOSCE_RESULT")
DEFAULT_CLEAN_DIR = os.path.join(DEFAULT_BASE_DIR, "CLEAN_CAOSCE_RESULT")

LOGO_PATH = os.path.join(os.path.expanduser("~"), "student_result_cleaner", "launcher", "static", "LOGO_YAN.jpg")

TIMESTAMP_FMT = "%Y-%m-%d_%H%M%S"
OUTPUT_BASENAME = "CAOSCE_PRE_COUNCIL_CLEANED"

# This will be dynamically built based on actual max scores found
STATION_COLUMN_MAP = {
    "procedure_station_one": "PS1_Score",
    "procedure_station_three": "PS3_Score",
    "procedure_station_five": "PS5_Score",
    "question_station_two": "QS2_Score",
    "question_station_four": "QS4_Score",
    "question_station_six": "QS6_Score",
    "viva": "VIVA",
}

# Will store the actual max scores for each station
station_max_scores = {}

# Styling - INCREASED FONT SIZES
NO_SCORE_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
NO_SCORE_FONT = Font(bold=True, color="9C0006", size=10, name="Calibri")
HEADER_FONT = Font(bold=True, size=10, name="Calibri", color="FFFFFF")
TITLE_FONT = Font(bold=True, size=16, name="Calibri", color="1F4E78")  # Increased from 14 to 16
SUBTITLE_FONT = Font(bold=True, size=14, name="Calibri", color="1F4E78")  # Increased from 12 to 14
DATE_FONT = Font(bold=True, size=11, name="Calibri", color="1F4E78")  # Increased from 10 to 11
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
AVERAGE_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
AVERAGE_FONT = Font(bold=True, size=10, name="Calibri", color="7F6000")
SUMMARY_HEADER_FONT = Font(bold=True, size=11, name="Calibri", color="1F4E78", underline="single")
SUMMARY_BODY_FONT = Font(size=10, name="Calibri")
ANALYSIS_HEADER_FONT = Font(bold=True, size=11, name="Calibri", color="1F4E78", underline="single")
ANALYSIS_BODY_FONT = Font(bold=True, size=10, name="Calibri")
SIGNATURE_FONT = Font(bold=True, size=10, name="Calibri")

UNWANTED_COL_PATTERNS = [
    r"phone", r"department", r"city", r"town", r"state",
    r"started on", r"started", r"completed", r"time taken", r"duration",
    r"q\.?\s*\d+", r"email", r"status", r"date", r"groups?",
]

# ---------------------------
# Helpers
# ---------------------------

def find_first_col(df, candidates):
    cols = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        cand_norm = cand.strip().lower()
        if cand_norm in cols:
            return cols[cand_norm]
        for c in df.columns:
            if cand_norm in str(c).strip().lower():
                return c
    return None


def find_username_col(df):
    return find_first_col(df, [
        "username", "user name", "exam no", "exam number", "registration no",
        "reg no", "mat no", "matno", "regnum", "last name", "surname",
        "groups", "group", "id"
    ])


def find_fullname_col(df):
    return find_first_col(df, [
        "user full name", "full name", "name", "candidate name", "student name",
        "first name", "given name", "first names"
    ])


def find_viva_score_col(df):
    return find_first_col(df, [
        "enter student score below", "enter student's score", "score", "grade"
    ])


def find_grade_column(df):
    """
    Dynamically find the grade/score column - looks for patterns like:
    Grade/10, Grade/3.25, Grade/20, Grade, Total, or any column containing 'grade' or 'total'
    Returns tuple: (column_name, max_score) where max_score is extracted from Grade/X
    """
    for c in df.columns:
        cn = str(c).strip()
        cn_lower = cn.lower()
        
        # Check for Grade/X pattern and extract the denominator
        match = re.match(r'^grade/([\d.]+)', cn_lower)
        if match:
            max_score = float(match.group(1))
            return (c, max_score)
        
        # Check for exact matches (assume /10 if not specified)
        if cn_lower in ["grade", "total", "score"]:
            return (c, 10.0)
        
        # Check for partial matches (assume /10 if not specified)
        if "grade" in cn_lower or "total" in cn_lower:
            return (c, 10.0)
    
    return (None, 10.0)


def extract_exam_number_from_fullname(text):
    if pd.isna(text):
        return None
    s = str(text).strip().upper()
    match = re.search(r'\b([A-Z]+/[A-Z0-9]+/\d+)\b', s)
    if match:
        return match.group(1)
    match = re.search(r'\b([A-Z]{1,3}\d{3,})\b', s)
    if match:
        return match.group(1)
    match = re.search(r'\b([A-Z]\d+/\d+)\b', s)
    if match:
        return match.group(1)
    return None


def extract_fullname_from_text(text, exam_no):
    if pd.isna(text):
        return None
    s = str(text).strip()
    if exam_no:
        s = s.replace(exam_no, "")
    s = re.sub(r'\s*-\s*', ' ', s)
    s = " ".join(s.split())
    if s and re.search(r"[A-Za-z]{3,}", s):
        return s
    return None


def sanitize_exam_no(v):
    if pd.isna(v):
        return ""
    s = str(v).strip().upper()
    s = re.sub(r"\.0+$", "", s)
    if " - " in s:
        s = s.split(" - ")[0].strip()
    return s


def numeric_safe(v):
    try:
        if pd.isna(v) or str(v).strip() == "":
            return None
        v2 = str(v).strip().replace(",", "")
        return float(v2)
    except:
        return None


def apply_autofit_columns(ws, header_row, data_end_row):
    """
    Apply optimal column widths based on content in data range only
    """
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_length = 0
        
        # Check header first
        header_cell = ws.cell(row=header_row, column=col_idx)
        header_value = str(header_cell.value or "")
        max_length = len(header_value)
        
        # Check data rows (from header+1 to data_end_row)
        for row_idx in range(header_row + 1, data_end_row + 1):
            try:
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value is not None:
                    cell_value = str(cell.value)
                    # For numeric values, use formatted length
                    if isinstance(cell.value, (int, float)):
                        if cell.number_format == '0.00':
                            cell_value = f"{cell.value:.2f}"
                        elif cell.number_format == '0':
                            cell_value = f"{int(cell.value)}"
                    
                    cell_length = len(cell_value)
                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass
        
        # Apply optimal width with padding
        optimal_width = max_length + 2  # Add padding
        
        # Set reasonable limits
        if header_value == "S/N":
            optimal_width = max(6, min(optimal_width, 8))
        elif header_value == "MAT NO.":
            optimal_width = max(12, min(optimal_width, 20))
        elif header_value == "FULL NAME":
            optimal_width = max(25, min(optimal_width, 35))
        elif "Score/" in header_value or "VIVA/" in header_value:
            optimal_width = max(10, min(optimal_width, 15))
        elif "Total Raw Score" in header_value:
            optimal_width = max(15, min(optimal_width, 20))
        elif "Percentage" in header_value:
            optimal_width = max(12, min(optimal_width, 15))
        else:
            optimal_width = max(10, min(optimal_width, 20))
        
        ws.column_dimensions[col_letter].width = optimal_width


def create_document_sections(ws, total_students, avg_percentage, highest_percentage, lowest_percentage, 
                           station_max_scores, total_max_score, data_end_row):
    """
    Create well-structured summary, analysis and signatories sections
    in a separate area that doesn't affect column widths
    """
    # Start the documentation section after the data with proper spacing
    doc_start_row = data_end_row + 3
    
    # ====================== SUMMARY SECTION ======================
    summary_header_row = doc_start_row
    ws.merge_cells(f"A{summary_header_row}:G{summary_header_row}")
    ws.cell(row=summary_header_row, column=1, value="EXAMINATION SUMMARY")
    ws.cell(row=summary_header_row, column=1).font = SUMMARY_HEADER_FONT
    ws.cell(row=summary_header_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    # Examination Structure - REMOVED DETAILED BREAKDOWN
    structure_rows = [
        "",
        f"Total Possible Score: {total_max_score} marks",
        "",
        "Scoring Methodology:",
        "Total Raw Score = Sum of all station scores",
        f"Percentage = (Total Raw Score ÷ {total_max_score}) × 100",
    ]
    
    for i, line in enumerate(structure_rows, 1):
        row_num = summary_header_row + i
        ws.merge_cells(f"A{row_num}:G{row_num}")
        cell = ws.cell(row=row_num, column=1, value=line)
        if "Methodology:" in line:
            cell.font = Font(bold=True, size=10, name="Calibri")
        else:
            cell.font = SUMMARY_BODY_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
    
    # ====================== ANALYSIS SECTION ======================
    analysis_start_row = summary_header_row + len(structure_rows) + 2
    
    ws.merge_cells(f"A{analysis_start_row}:G{analysis_start_row}")
    ws.cell(row=analysis_start_row, column=1, value="PERFORMANCE ANALYSIS")
    ws.cell(row=analysis_start_row, column=1).font = ANALYSIS_HEADER_FONT
    ws.cell(row=analysis_start_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    # Analysis rows - REMOVED PERFORMANCE CATEGORIES
    analysis_rows = [
        "",
        f"Total Candidates: {total_students}",
        f"Overall Average Percentage: {avg_percentage}%",
        f"Highest Percentage Score: {highest_percentage}%",
        f"Lowest Percentage Score: {lowest_percentage}%",
    ]
    
    for i, line in enumerate(analysis_rows, 1):
        row_num = analysis_start_row + i
        ws.merge_cells(f"A{row_num}:G{row_num}")
        cell = ws.cell(row=row_num, column=1, value=line)
        cell.font = ANALYSIS_BODY_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
    
    # ====================== SIGNATORIES SECTION ======================
    signatories_start_row = analysis_start_row + len(analysis_rows) + 3
    
    # Prepared by section (left aligned)
    ws.merge_cells(f"A{signatories_start_row}:C{signatories_start_row}")
    ws.cell(row=signatories_start_row, column=1, value="Prepared by:")
    ws.cell(row=signatories_start_row, column=1).font = SIGNATURE_FONT
    
    ws.merge_cells(f"A{signatories_start_row + 2}:C{signatories_start_row + 2}")
    ws.cell(row=signatories_start_row + 2, column=1, value="_________________________")
    
    ws.merge_cells(f"A{signatories_start_row + 3}:C{signatories_start_row + 3}")
    ws.cell(row=signatories_start_row + 3, column=1, value="Examiner's Signature")
    ws.cell(row=signatories_start_row + 3, column=1).font = SUMMARY_BODY_FONT
    
    ws.merge_cells(f"A{signatories_start_row + 5}:C{signatories_start_row + 5}")
    ws.cell(row=signatories_start_row + 5, column=1, value="Name: _________________________")
    
    ws.merge_cells(f"A{signatories_start_row + 6}:C{signatories_start_row + 6}")
    ws.cell(row=signatories_start_row + 6, column=1, value="Date: __________________________")
    
    # Approved by section (PROPERLY ALIGNED)
    # Use columns D to G for the Approved by section (4 columns wide)
    approved_col_start = 4  # Column D
    approved_col_end = 7    # Column G
    
    # Approved by header
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row}:{get_column_letter(approved_col_end)}{signatories_start_row}")
    approved_cell = ws.cell(row=signatories_start_row, column=approved_col_start, value="Approved by:")
    approved_cell.font = SIGNATURE_FONT
    approved_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Signature line
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row + 2}:{get_column_letter(approved_col_end)}{signatories_start_row + 2}")
    signature_cell = ws.cell(row=signatories_start_row + 2, column=approved_col_start, value="_________________________")
    signature_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Provost's Signature text
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row + 3}:{get_column_letter(approved_col_end)}{signatories_start_row + 3}")
    provost_cell = ws.cell(row=signatories_start_row + 3, column=approved_col_start, value="Provost's Signature")
    provost_cell.font = SUMMARY_BODY_FONT
    provost_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Name line
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row + 5}:{get_column_letter(approved_col_end)}{signatories_start_row + 5}")
    name_cell = ws.cell(row=signatories_start_row + 5, column=approved_col_start, value="Name: _________________________")
    name_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Date line
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row + 6}:{get_column_letter(approved_col_end)}{signatories_start_row + 6}")
    date_cell = ws.cell(row=signatories_start_row + 6, column=approved_col_start, value="Date: __________________________")
    date_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    return signatories_start_row + 8  # Return the last row used

# ---------------------------
# Main processing
# ---------------------------

def process_files():
    print("Starting CAOSCE Pre-Council Results Cleaning...\n")

    RAW_DIR = DEFAULT_RAW_DIR
    BASE_CLEAN_DIR = DEFAULT_CLEAN_DIR

    ts = datetime.now().strftime(TIMESTAMP_FMT)
    output_dir = os.path.join(BASE_CLEAN_DIR, f"CAOSCE_PRE_COUNCIL_{ts}")
    os.makedirs(output_dir, exist_ok=True)

    files = [f for f in os.listdir(RAW_DIR) if f.lower().endswith((".xlsx", ".xls", ".csv"))]

    if not files:
        print(f"No raw files found in {RAW_DIR}")
        return

    results = {}
    # Dictionary to track max scores found for each station
    station_max_scores = {}

    for fname in sorted(files):
        path = os.path.join(RAW_DIR, fname)
        lower = fname.lower()
        station_key = None
        if "procedure" in lower or "ps-" in lower or "ps1" in lower or "ps3" in lower or "ps5" in lower:
            if "one" in lower or "ps1" in lower or "_1" in lower:
                station_key = "procedure_station_one"
            elif "three" in lower or "ps3" in lower or "_3" in lower:
                station_key = "procedure_station_three"
            elif "five" in lower or "ps5" in lower or "_5" in lower:
                station_key = "procedure_station_five"
        elif "question" in lower or "qs-" in lower or "qs2" in lower or "qs4" in lower or "qs6" in lower:
            if "two" in lower or "qs2" in lower or "_2" in lower:
                station_key = "question_station_two"
            elif "four" in lower or "qs4" in lower or "_4" in lower:
                station_key = "question_station_four"
            elif "six" in lower or "qs6" in lower or "_6" in lower:
                station_key = "question_station_six"
        elif "viva" in lower:
            station_key = "viva"

        if not station_key:
            print(f"Could not determine station for {fname} – skipping")
            continue

        try:
            if fname.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str)
            else:
                df = pd.read_excel(path, dtype=str)
        except Exception as e:
            print(f"Error reading {fname}: {e}")
            continue

        df.rename(columns=lambda c: str(c).strip(), inplace=True)

        username_col = find_username_col(df)
        fullname_col = find_fullname_col(df)
        grade_col, max_score = find_grade_column(df)  # Now returns (column, max_score)
        viva_score_col = find_viva_score_col(df) if station_key == "viva" else None

        # Debug: Print what grade column was found
        if grade_col:
            print(f"  Found grade column: '{grade_col}' (max score: {max_score}) in {fname}")
            # Store the max score for this station
            station_max_scores[station_key] = max_score
        else:
            print(f"  Warning: No grade column found in {fname}")
            # Default to 10 if not found
            station_max_scores[station_key] = 10.0

        for pattern in UNWANTED_COL_PATTERNS:
            df.drop(columns=[c for c in df.columns if re.search(pattern, str(c), flags=re.I)], inplace=True, errors="ignore")

        rows_added = 0
        for _, row in df.iterrows():
            exam_no = None
            full_name = None
            if station_key == "viva":
                if fullname_col and pd.notna(row.get(fullname_col)):
                    fullname_value = str(row[fullname_col]).strip()
                    exam_no = extract_exam_number_from_fullname(fullname_value)
                    if exam_no:
                        full_name = extract_fullname_from_text(fullname_value, exam_no)
                if not exam_no and username_col:
                    exam_no = sanitize_exam_no(row.get(username_col))
                if not exam_no:
                    print(f"  Warning: Could not extract exam number from VIVA row in file {fname}")
                    continue
            else:
                if username_col:
                    raw_value = row.get(username_col)
                    exam_no = sanitize_exam_no(raw_value)
                    if exam_no and pd.notna(raw_value):
                        full_name = extract_fullname_from_text(str(raw_value), exam_no)
                if not full_name and fullname_col and pd.notna(row.get(fullname_col)):
                    fullname_value = str(row[fullname_col]).strip()
                    if not exam_no:
                        exam_no = extract_exam_number_from_fullname(fullname_value)
                    full_name = extract_fullname_from_text(fullname_value, exam_no)
                if not exam_no:
                    for c in df.columns:
                        val = sanitize_exam_no(row.get(c))
                        if val and len(val) > 2 and re.search(r'\d', val):
                            exam_no = val
                            break
                if not full_name:
                    for c in df.columns:
                        if c == username_col or c == fullname_col:
                            continue
                        val = str(row.get(c, "")).strip()
                        if val and re.search(r"[A-Za-z]{3,}", val) and not re.search(r'\d{2,}', val):
                            full_name = val
                            break
                if not exam_no:
                    continue

            if exam_no not in results:
                results[exam_no] = {
                    "MAT NO.": exam_no,
                    "FULL NAME": None,
                }
                # Initialize all station scores to None
                for sk in STATION_COLUMN_MAP.keys():
                    results[exam_no][STATION_COLUMN_MAP[sk]] = None

            if full_name and not results[exam_no]["FULL NAME"]:
                results[exam_no]["FULL NAME"] = full_name

            score_val = None
            if station_key == "viva" and viva_score_col:
                score_val = numeric_safe(row.get(viva_score_col))
            elif grade_col:
                score_val = numeric_safe(row.get(grade_col))

            if score_val is not None:
                out_col = STATION_COLUMN_MAP[station_key]
                results[exam_no][out_col] = round(score_val, 2)

            rows_added += 1

        print(f"Processed {fname} → {rows_added} rows (Station: {station_key})")

    if not results:
        print("No student data found.")
        return

    # Build the score columns with actual denominators
    score_cols = []
    for station_key in ["procedure_station_one", "procedure_station_three", "procedure_station_five",
                        "question_station_two", "question_station_four", "question_station_six", "viva"]:
        base_col = STATION_COLUMN_MAP[station_key]
        max_score = station_max_scores.get(station_key, 10.0)
        # Format the denominator nicely (remove .0 if it's a whole number)
        if max_score == int(max_score):
            col_name = f"{base_col}/{int(max_score)}"
        else:
            col_name = f"{base_col}/{max_score}"
        score_cols.append(col_name)
    
    # Rename columns in results to include denominators
    for exam_no in results:
        student_data = results[exam_no].copy()
        for i, station_key in enumerate(["procedure_station_one", "procedure_station_three", "procedure_station_five",
                                         "question_station_two", "question_station_four", "question_station_six", "viva"]):
            base_col = STATION_COLUMN_MAP[station_key]
            if base_col in student_data:
                results[exam_no][score_cols[i]] = student_data[base_col]
                if base_col != score_cols[i]:
                    del results[exam_no][base_col]

    base_cols = ["MAT NO.", "FULL NAME"] + score_cols
    df_out = pd.DataFrame.from_dict(results, orient="index")[base_cols]

    df_out["FULL NAME"] = df_out.groupby("MAT NO.")["FULL NAME"].transform(
        lambda x: x.fillna(method='ffill').fillna(method='bfill')
    )

    df_out["__sort"] = pd.to_numeric(df_out["MAT NO."].str.extract(r'(\d+)')[0], errors='coerce')
    df_out.sort_values(["__sort", "MAT NO."], inplace=True)
    df_out.drop(columns=["__sort"], inplace=True)
    df_out.reset_index(drop=True, inplace=True)

    df_out.insert(0, "S/N", range(1, len(df_out) + 1))

    df_out[score_cols] = df_out[score_cols].apply(pd.to_numeric, errors="coerce").round(2).fillna(0.00)

    # Calculate total using actual max scores
    total_max_score = sum(station_max_scores.values())
    df_out[f"Total Raw Score /{total_max_score:.2f}".replace(".00", "")] = df_out[score_cols].sum(axis=1).round(2)
    total_col_name = f"Total Raw Score /{total_max_score:.2f}".replace(".00", "")
    
    df_out["Percentage (%)"] = (df_out[total_col_name] / total_max_score * 100).round(0).astype(int)

    final_display_cols = ["S/N", "MAT NO.", "FULL NAME"] + score_cols + [total_col_name, "Percentage (%)"]
    df_out = df_out[final_display_cols]

    # Remove any pre-existing OVERALL AVERAGE rows (if they came from input)
    df_out = df_out[df_out["MAT NO."] != "OVERALL AVERAGE"].copy()

    total_students = len(df_out)
    student_percentages = df_out["Percentage (%)"].values
    avg_percentage = int(round(student_percentages.mean())) if total_students > 0 else 0
    highest_percentage = int(student_percentages.max()) if total_students > 0 else 0
    lowest_percentage = int(student_percentages.min()) if total_students > 0 else 0

    # Compute and add a single OVERALL AVERAGE row
    avg_row = {
        "S/N": "",
        "MAT NO.": "OVERALL AVERAGE",
        "FULL NAME": "",
    }
    for col in score_cols:
        avg_row[col] = df_out[col].mean().round(2) if total_students > 0 else 0.00
    avg_row[total_col_name] = df_out[total_col_name].mean().round(2) if total_students > 0 else 0.00
    avg_row["Percentage (%)"] = int(df_out["Percentage (%)"].mean().round(0)) if total_students > 0 else 0

    df_out = pd.concat([df_out, pd.DataFrame([avg_row])], ignore_index=True)

    # Save outputs
    out_csv = os.path.join(output_dir, f"{OUTPUT_BASENAME}_{ts}.csv")
    out_xlsx = os.path.join(output_dir, f"{OUTPUT_BASENAME}_{ts}.xlsx")

    df_out.to_csv(out_csv, index=False)
    df_out.to_excel(out_xlsx, index=False, engine="openpyxl")

    # ====================== Excel Formatting ======================
    wb = load_workbook(out_xlsx)
    ws = wb.active

    TITLE_ROWS = 6
    ws.insert_rows(1, TITLE_ROWS)
    header_row = TITLE_ROWS + 1

    last_col_letter = get_column_letter(ws.max_column)

    # Add logo if exists - INCREASED LOGO SIZE
    if os.path.exists(LOGO_PATH):
        try:
            img = XLImage(LOGO_PATH)
            img.width = 120  # Increased from 80 to 120
            img.height = 120  # Increased from 80 to 120
            ws.add_image(img, "A1")
        except Exception as e:
            print(f"Warning: Could not add logo: {e}")

    # Titles - WITH INCREASED FONT SIZES
    ws.merge_cells(f"B1:{last_col_letter}1")
    ws["B1"] = "YAGONGWO COLLEGE OF NURSING SCIENCE, KUJE, ABUJA"
    ws["B1"].font = TITLE_FONT  # Now size 16 (increased from 14)
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28  # Increased height for larger font

    ws.merge_cells(f"B2:{last_col_letter}2")
    ws["B2"] = "PRE-COUNCIL EXAMINATION RESULT"
    ws["B2"].font = SUBTITLE_FONT  # Now size 14 (increased from 12)
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24  # Increased height for larger font

    ws.merge_cells(f"B3:{last_col_letter}3")
    ws["B3"] = "CAOSCE_2025"
    ws["B3"].font = Font(bold=True, size=12, name="Calibri", color="1F4E78")  # Increased from 11 to 12
    ws["B3"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 20  # Increased height

    # DATE directly under CAOSCE_2025 (centered across full width)
    ws.merge_cells(f"B4:{last_col_letter}4")
    ws.cell(row=4, column=2, value=f"Date: {datetime.now().strftime('%d %B %Y')}")
    ws.cell(row=4, column=2).font = DATE_FONT  # Now size 11 (increased from 10)
    ws.cell(row=4, column=2).alignment = Alignment(horizontal="center", vertical="center")

    # CLASS - moved to separate row and properly aligned to not interfere with column fitting
    ws.merge_cells(f"A5:{last_col_letter}5")
    class_cell = ws.cell(row=5, column=1, value="CLASS: _________________________________________________________________________________________")
    class_cell.font = Font(bold=True, size=11, name="Calibri", color="1F4E78")  # Increased from 10 to 11
    class_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[4].height = 20  # Increased height
    ws.row_dimensions[5].height = 18  # Increased height

    # Empty row for spacing
    ws.row_dimensions[6].height = 10  # Increased spacing

    # Header styling - keep overall header centered but force specific headers left
    for cell in ws[header_row]:
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        cell.fill = HEADER_FILL
    ws.row_dimensions[header_row].height = 20

    # Force MAT NO. and FULL NAME header alignment to LEFT
    mat_no_idx = 2
    full_name_idx = 3
    ws.cell(row=header_row, column=mat_no_idx).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.cell(row=header_row, column=full_name_idx).alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Borders
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=header_row, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border

    ws.freeze_panes = f"A{header_row + 1}"

    station_col_indices = list(range(4, 4 + len(score_cols)))
    total_col_idx = 4 + len(score_cols)
    percent_col_idx = total_col_idx + 1

    # Identify avg row
    avg_row_num = None
    for r in range(header_row + 1, ws.max_row + 1):
        if ws.cell(row=r, column=mat_no_idx).value == "OVERALL AVERAGE":
            avg_row_num = r
            break

    # Format data rows and force MAT NO./FULL NAME left alignment
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        is_avg_row = (row[0].row == avg_row_num)
        for cell in row:
            if cell.column == 1:  # S/N
                if is_avg_row:
                    cell.value = ""
                    cell.fill = AVERAGE_FILL
                else:
                    cell.number_format = "0"
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(size=10, name="Calibri")
            elif cell.column == mat_no_idx:
                if is_avg_row:
                    cell.font = AVERAGE_FONT
                    cell.fill = AVERAGE_FILL
                else:
                    cell.font = Font(size=10, name="Calibri")
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif cell.column == full_name_idx:
                if is_avg_row:
                    cell.value = ""
                    cell.fill = AVERAGE_FILL
                else:
                    cell.font = Font(size=10, name="Calibri")
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif cell.column in station_col_indices:
                if is_avg_row:
                    cell.font = AVERAGE_FONT
                    cell.fill = AVERAGE_FILL
                elif cell.value == 0 or cell.value == 0.0 or cell.value is None:
                    cell.value = 0.00
                    cell.fill = NO_SCORE_FILL
                    cell.font = NO_SCORE_FONT
                else:
                    cell.font = Font(size=10, name="Calibri")
                cell.number_format = "0.00"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif cell.column in [total_col_idx, percent_col_idx]:
                if cell.column == percent_col_idx:
                    cell.number_format = "0"
                else:
                    cell.number_format = "0.00"
                if is_avg_row:
                    cell.font = AVERAGE_FONT
                    cell.fill = AVERAGE_FILL
                else:
                    cell.font = Font(bold=True, size=10, name="Calibri")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(size=10, name="Calibri")

    # Apply autofit columns based on content (only data rows)
    data_end_row = ws.max_row  # Store the end of data before adding summary
    apply_autofit_columns(ws, header_row, data_end_row)

    # ====================== CREATE SEPARATE DOCUMENTATION SECTION ======================
    last_row = create_document_sections(
        ws, total_students, avg_percentage, highest_percentage, lowest_percentage,
        station_max_scores, total_max_score, data_end_row
    )

    # Print setup - include the documentation section
    ws.print_area = f"A1:{last_col_letter}{last_row}"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    wb.save(out_xlsx)

    print(f"\n✓ Saved: {os.path.basename(out_csv)}")
    print(f"✓ Saved formatted Excel: {os.path.basename(out_xlsx)}")
    print(f"\n📁 Files saved in: {output_dir}")
    print(f"\n📊 Summary: {total_students} students processed")
    print(f"   Average: {int(avg_percentage)}% | Highest: {int(highest_percentage)}% | Lowest: {int(lowest_percentage)}%")


if __name__ == "__main__":
    process_files()