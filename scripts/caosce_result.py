#!/usr/bin/env python3 
"""
caosce_result_enhanced_refactored.py
ENHANCED CAOSCE cleaning script with multi-college support and multi-paper processing
- Processes Paper I, Paper II, and CAOSCE station-based exams
- Creates ONE workbook with TWO sheets: CAOSCE Results and Combined Results
- Removes "Overall average" student rows from raw data
- Extracts actual overall averages from raw files
- Creates single OVERALL AVERAGE row with proper values
- Dynamic logo selection based on college detection
"""

import os
import re
import copy
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.dataframe import dataframe_to_rows
import logging

# ---------------------------
# College Configuration
# ---------------------------
COLLEGE_CONFIGS = {
    "YAGONGWO": {
        "name": "YAGONGWO COLLEGE OF NURSING SCIENCE, KUJE, ABUJA",
        "exam_patterns": [r'BN/A\d{2}/\d{3}', r'BN/\w+/\d+'],
        "mat_no_label": "MAT NO.",
        "logo": "LOGO_YAN.png",
        "output_prefix": "YAGONGWO"
    },
    "FCT": {
        "name": "FCT COLLEGE OF NURSING SCIENCES, GWAGWALADA, FCT-ABUJA", 
        "exam_patterns": [r'^\d{4}$', r'FCTCONS/ND\d{2}/\d{3}', r'FCTCONS/\w+/\d+'],
        "mat_no_label": "EXAM NO.",
        "logo": "logo.png",
        "output_prefix": "FCT"
    }
}

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

LOGO_BASE_PATH = os.path.join(os.path.expanduser("~"), "student_result_cleaner", "launcher", "static")

TIMESTAMP_FMT = "%Y-%m-%d_%H%M%S"
CURRENT_YEAR = datetime.now().year

# Paper patterns
PAPER_I_PATTERNS = [
    r"PAPERI_PAPERII-PAPER I-grades",
    r"PAPER I", 
    r"PAPER 1",
    r"PAPERI"
]

PAPER_II_PATTERNS = [
    r"PAPERI_PAPERII-PAPER II-grades",
    r"PAPER II",
    r"PAPER 2", 
    r"PAPERII"
]

# Station mapping
STATION_COLUMN_MAP = {
    "procedure_station_one": "PS1_Score",
    "procedure_station_three": "PS3_Score",
    "procedure_station_five": "PS5_Score",
    "question_station_two": "QS2_Score",
    "question_station_four": "QS4_Score",
    "question_station_six": "QS6_Score",
    "viva": "VIVA",
}

# Station display names with denominators
STATION_DISPLAY_NAMES = {
    "procedure_station_one": "PS1",
    "procedure_station_three": "PS3", 
    "procedure_station_five": "PS5",
    "question_station_two": "QS2",
    "question_station_four": "QS4",
    "question_station_six": "QS6",
    "viva": "VIVA"
}

# Styling
NO_SCORE_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
NO_SCORE_FONT = Font(bold=True, color="9C0006", size=10, name="Calibri")
HEADER_FONT = Font(bold=True, size=10, name="Calibri", color="FFFFFF")
TITLE_FONT = Font(bold=True, size=16, name="Calibri", color="1F4E78")
SUBTITLE_FONT = Font(bold=True, size=14, name="Calibri", color="1F4E78")
DATE_FONT = Font(bold=True, size=11, name="Calibri", color="1F4E78")
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

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ---------------------------
# Helper Functions
# ---------------------------

def detect_college_from_exam_numbers(exam_numbers):
    """
    Detect college based on exam number patterns with better logic
    """
    if not exam_numbers:
        return "YAGONGWO", COLLEGE_CONFIGS["YAGONGWO"]
    
    yagongwo_count = 0
    fct_count = 0
    
    for exam_no in exam_numbers:
        exam_no_str = str(exam_no).strip().upper()
        
        # Check Yagongwo patterns
        yagongwo_matched = False
        for pattern in COLLEGE_CONFIGS["YAGONGWO"]["exam_patterns"]:
            if re.match(pattern, exam_no_str):
                yagongwo_count += 1
                yagongwo_matched = True
                break
        
        # Check FCT patterns only if not matched by Yagongwo
        if not yagongwo_matched:
            for pattern in COLLEGE_CONFIGS["FCT"]["exam_patterns"]:
                if re.match(pattern, exam_no_str):
                    fct_count += 1
                    break
    
    logger.info(f"College detection - Yagongwo: {yagongwo_count}, FCT: {fct_count}")
    
    if fct_count > yagongwo_count:
        return "FCT", COLLEGE_CONFIGS["FCT"]
    else:
        return "YAGONGWO", COLLEGE_CONFIGS["YAGONGWO"]

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
        "last name",           # ADDED FIRST - for paper files
        "surname",             # ADDED SECOND - for paper files
        "username", 
        "user name", 
        "exam no", 
        "exam number", 
        "registration no",
        "reg no", 
        "mat no", 
        "matno", 
        "regnum",
        "groups", 
        "group", 
        "id"
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
    Dynamically find the grade/score column
    Returns tuple: (column_name, max_score)
    """
    for c in df.columns:
        cn = str(c).strip()
        cn_lower = cn.lower()
        
        # Check for Grade/X pattern and extract the denominator
        match = re.match(r'^grade/([\d.]+)', cn_lower)
        if match:
            max_score = float(match.group(1))
            return (c, max_score)
        
        # Check for exact matches (assume /100 if not specified for papers, /10 for stations)
        if cn_lower in ["grade", "total", "score"]:
            return (c, 10.0)  # Default to 10 for stations
        
        # Check for partial matches
        if "grade" in cn_lower or "total" in cn_lower:
            return (c, 10.0)  # Default to 10 for stations
    
    return (None, 10.0)

def extract_exam_number_from_fullname(text):
    if pd.isna(text):
        return None
    s = str(text).strip().upper()
    
    # Try Yagongwo pattern: BN/A23/002
    match = re.search(r'\b(BN/A\d{2}/\d{3})\b', s)
    if match:
        return match.group(1)
    
    # Try FCT pattern: FCTCONS/ND24/001
    match = re.search(r'\b(FCTCONS/ND\d{2}/\d{3})\b', s)
    if match:
        return match.group(1)
    
    # Try 4-digit pattern: 7433
    match = re.search(r'\b(\d{4})\b', s)
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

def is_overall_average_row(row, username_col, fullname_col):
    """Check if a row represents an overall average row (not a real student)"""
    if username_col and pd.notna(row.get(username_col)):
        if 'overall' in str(row[username_col]).lower() and 'average' in str(row[username_col]).lower():
            return True
    if fullname_col and pd.notna(row.get(fullname_col)):
        if 'overall' in str(row[fullname_col]).lower() and 'average' in str(row[fullname_col]).lower():
            return True
    return False

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
        optimal_width = max_length + 2
        
        # Set reasonable limits
        if header_value == "S/N":
            optimal_width = max(6, min(optimal_width, 8))
        elif header_value in ["MAT NO.", "EXAM NO."]:
            optimal_width = max(12, min(optimal_width, 20))
        elif header_value == "FULL NAME":
            optimal_width = max(25, min(optimal_width, 35))
        elif "Score/" in header_value or "VIVA/" in header_value:
            optimal_width = max(10, min(optimal_width, 15))
        elif "Total Raw Score" in header_value:
            optimal_width = max(15, min(optimal_width, 20))
        elif "Percentage" in header_value:
            optimal_width = max(12, min(optimal_width, 15))
        elif "PAPER I" in header_value or "PAPER II" in header_value or "CAOSCE" in header_value:
            optimal_width = max(12, min(optimal_width, 15))
        elif "OVERALL AVERAGE" in header_value:
            optimal_width = max(15, min(optimal_width, 20))
        else:
            optimal_width = max(10, min(optimal_width, 20))
        
        ws.column_dimensions[col_letter].width = optimal_width

def create_document_sections(ws, total_students, avg_percentage, highest_percentage, lowest_percentage, 
                           total_max_score, data_end_row, college_config, sheet_type="CAOSCE"):
    """
    Create well-structured summary, analysis and signatories sections
    """
    doc_start_row = data_end_row + 3
    
    # ====================== SUMMARY SECTION ======================
    summary_header_row = doc_start_row
    last_col_letter = get_column_letter(ws.max_column)
    ws.merge_cells(f"A{summary_header_row}:{last_col_letter}{summary_header_row}")
    
    if sheet_type == "COMBINED":
        ws.cell(row=summary_header_row, column=1, value="EXAMINATION SUMMARY")
        summary_rows = [
            "",
            f"Total Possible Score: 300 marks (100 per paper)",
            "",
            "Scoring Methodology:",
            "- Paper I Score: 0-100 marks",
            "- Paper II Score: 0-100 marks",
            "- CAOSCE Score: 0-100 marks (percentage from station performance)",
            "- Overall Average = (Paper I + Paper II + CAOSCE) ÷ 3",
        ]
    else:
        ws.cell(row=summary_header_row, column=1, value="EXAMINATION SUMMARY")
        summary_rows = [
            "",
            f"Total Possible Score: {total_max_score} marks",
            "",
            "Scoring Methodology:",
            "Total Raw Score = Sum of all station scores",
            f"Percentage = (Total Raw Score ÷ {total_max_score}) × 100",
        ]
    
    ws.cell(row=summary_header_row, column=1).font = SUMMARY_HEADER_FONT
    ws.cell(row=summary_header_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    for i, line in enumerate(summary_rows, 1):
        row_num = summary_header_row + i
        ws.merge_cells(f"A{row_num}:{last_col_letter}{row_num}")
        cell = ws.cell(row=row_num, column=1, value=line)
        if "Methodology:" in line or line.startswith("-"):
            cell.font = Font(bold=True, size=10, name="Calibri")
        else:
            cell.font = SUMMARY_BODY_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
    
    # ====================== ANALYSIS SECTION ======================
    analysis_start_row = summary_header_row + len(summary_rows) + 2
    
    ws.merge_cells(f"A{analysis_start_row}:{last_col_letter}{analysis_start_row}")
    ws.cell(row=analysis_start_row, column=1, value="PERFORMANCE ANALYSIS")
    ws.cell(row=analysis_start_row, column=1).font = ANALYSIS_HEADER_FONT
    ws.cell(row=analysis_start_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    if sheet_type == "COMBINED":
        # Get paper-wise averages from the overall average row
        paper_i_avg = 0
        paper_ii_avg = 0
        caosce_avg = 0
        
        # Find the overall average row and extract values
        # header_row is at TITLE_ROWS + 1 = 7
        calculated_header_row = 7
        for row_idx in range(calculated_header_row + 1, data_end_row + 1):
            if ws.cell(row=row_idx, column=2).value == "OVERALL AVERAGE":  # MAT NO. column
                paper_i_avg = ws.cell(row=row_idx, column=4).value or 0  # PAPER I column
                paper_ii_avg = ws.cell(row=row_idx, column=5).value or 0  # PAPER II column
                caosce_avg = ws.cell(row=row_idx, column=6).value or 0  # CAOSCE column
                break
        
        analysis_rows = [
            "",
            f"Total Candidates: {total_students}",
            f"Overall Average: {avg_percentage}%",
            f"Highest Score: {highest_percentage}%",
            f"Lowest Score: {lowest_percentage}%",
            "",
            "Paper-wise Averages:",
            f"- Paper I Average: {paper_i_avg:.1f}%",
            f"- Paper II Average: {paper_ii_avg:.1f}%", 
            f"- CAOSCE Average: {caosce_avg:.1f}%",
        ]
    else:
        analysis_rows = [
            "",
            f"Total Candidates: {total_students}",
            f"Overall Average Percentage: {avg_percentage}%",
            f"Highest Percentage Score: {highest_percentage}%",
            f"Lowest Percentage Score: {lowest_percentage}%",
        ]
    
    for i, line in enumerate(analysis_rows, 1):
        row_num = analysis_start_row + i
        ws.merge_cells(f"A{row_num}:{last_col_letter}{row_num}")
        cell = ws.cell(row=row_num, column=1, value=line)
        cell.font = ANALYSIS_BODY_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
    
    # ====================== SIGNATORIES SECTION ======================
    signatories_start_row = analysis_start_row + len(analysis_rows) + 3
    
    # Prepared by section
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
    
    # Approved by section
    approved_col_start = 4
    approved_col_end = 7
    
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row}:{get_column_letter(approved_col_end)}{signatories_start_row}")
    approved_cell = ws.cell(row=signatories_start_row, column=approved_col_start, value="Approved by:")
    approved_cell.font = SIGNATURE_FONT
    approved_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row + 2}:{get_column_letter(approved_col_end)}{signatories_start_row + 2}")
    signature_cell = ws.cell(row=signatories_start_row + 2, column=approved_col_start, value="_________________________")
    signature_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row + 3}:{get_column_letter(approved_col_end)}{signatories_start_row + 3}")
    provost_cell = ws.cell(row=signatories_start_row + 3, column=approved_col_start, value="Provost's Signature")
    provost_cell.font = SUMMARY_BODY_FONT
    provost_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row + 5}:{get_column_letter(approved_col_end)}{signatories_start_row + 5}")
    name_cell = ws.cell(row=signatories_start_row + 5, column=approved_col_start, value="Name: _________________________")
    name_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells(f"{get_column_letter(approved_col_start)}{signatories_start_row + 6}:{get_column_letter(approved_col_end)}{signatories_start_row + 6}")
    date_cell = ws.cell(row=signatories_start_row + 6, column=approved_col_start, value="Date: __________________________")
    date_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    return signatories_start_row + 8

def detect_paper_type(filename):
    """
    Detect if file is Paper I, Paper II, or CAOSCE station with case-insensitive matching
    """
    # Normalize filename for better matching
    fname_upper = filename.upper()
    fname_lower = filename.lower()
    
    # Check for Paper I patterns (case-insensitive)
    paper_i_indicators = [
        "PAPER I", "PAPER 1", "PAPERI", "PAPER_I", "PAPER-I",
        "PAPERI_PAPERII-PAPER I", "PAPER I-GRADES"
    ]
    
    for indicator in paper_i_indicators:
        if indicator.upper() in fname_upper:
            # Make sure it's not Paper II
            if "PAPER II" not in fname_upper and "PAPER 2" not in fname_upper:
                return "PAPER_I"
    
    # Check for Paper II patterns (case-insensitive)
    paper_ii_indicators = [
        "PAPER II", "PAPER 2", "PAPERII", "PAPER_II", "PAPER-II",
        "PAPERI_PAPERII-PAPER II", "PAPER II-GRADES"
    ]
    
    for indicator in paper_ii_indicators:
        if indicator.upper() in fname_upper:
            return "PAPER_II"
    
    # Check for CAOSCE station patterns
    if any(station in fname_lower for station in ["procedure", "question", "viva", "ps-", "qs-", "ps1", "ps3", "ps5", "qs2", "qs4", "qs6"]):
        return "CAOSCE_STATION"
    
    return "UNKNOWN"

def process_paper_files(files, raw_dir):
    """
    Process Paper I and Paper II files
    Returns: dict with paper results {exam_no: {"PAPER I": score, "PAPER II": score}}
    """
    paper_results = {}
    paper_averages = {"PAPER I": [], "PAPER II": []}  # Track scores for averaging
    
    for fname in files:
        paper_type = detect_paper_type(fname)
        if paper_type not in ["PAPER_I", "PAPER_II"]:
            continue
            
        path = os.path.join(raw_dir, fname)
        logger.info(f"Processing {paper_type} file: {fname}")
        
        try:
            if fname.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str)
            else:
                df = pd.read_excel(path, dtype=str)
        except Exception as e:
            logger.error(f"Error reading {fname}: {e}")
            continue
            
        df.rename(columns=lambda c: str(c).strip(), inplace=True)
        
        username_col = find_username_col(df)
        fullname_col = find_fullname_col(df)
        grade_col, max_score = find_grade_column(df)
        
        if not grade_col:
            logger.warning(f"No grade column found in {fname}")
            continue
            
        logger.info(f"  Found grade column: '{grade_col}' (max score: {max_score})")
        
        # Remove unwanted columns
        for pattern in UNWANTED_COL_PATTERNS:
            df.drop(columns=[c for c in df.columns if re.search(pattern, str(c), flags=re.I)], 
                   inplace=True, errors="ignore")
        
        rows_processed = 0
        paper_label = paper_type.replace("_", " ")  # "PAPER I" or "PAPER II"
        
        for _, row in df.iterrows():
            # Skip overall average rows
            if is_overall_average_row(row, username_col, fullname_col):
                continue
                
            exam_no = None
            full_name = None
            
            # Extract exam number - try username column first (it's "Last name" in your data)
            if username_col:
                exam_no = sanitize_exam_no(row.get(username_col))
            
            # If still no exam_no, try fullname column
            if not exam_no and fullname_col:
                fullname_value = str(row.get(fullname_col, "")).strip()
                exam_no = extract_exam_number_from_fullname(fullname_value)
                if exam_no:
                    full_name = extract_fullname_from_text(fullname_value, exam_no)
            
            # If still no exam_no, try to find it in any column
            if not exam_no:
                for col in df.columns:
                    val = str(row.get(col, "")).strip()
                    # Look for patterns like BN/A23/011 or 4-digit numbers
                    if re.search(r'BN/A\d{2}/\d{3}|FCTCONS/ND\d{2}/\d{3}|\b\d{4}\b', val):
                        exam_no = sanitize_exam_no(val)
                        break
            
            if not exam_no:
                logger.debug(f"  Skipping row - no exam number found")
                continue
            
            # Extract full name from fullname_col if available
            if not full_name and fullname_col:
                full_name = str(row.get(fullname_col, "")).strip()
                if full_name and not re.search(r'[A-Za-z]{3,}', full_name):
                    full_name = None
                
            # Initialize student record if not exists
            if exam_no not in paper_results:
                paper_results[exam_no] = {
                    "PAPER I": 0.00,
                    "PAPER II": 0.00,
                    "FULL NAME": full_name
                }
            
            # Update full name if not set
            if full_name and not paper_results[exam_no]["FULL NAME"]:
                paper_results[exam_no]["FULL NAME"] = full_name
                
            # Extract and normalize score
            score_val = numeric_safe(row.get(grade_col))
            
            if score_val is not None:
                # CRITICAL: Normalize to percentage out of 100
                # If Grade/10.00, convert: (3.00/10.00) * 100 = 30.00%
                # If Grade/100, keep as is: (75/100) * 100 = 75.00%
                normalized_score = (score_val / max_score) * 100
                paper_results[exam_no][paper_label] = round(normalized_score, 2)
                paper_averages[paper_label].append(normalized_score)
                rows_processed += 1
                logger.debug(f"  {exam_no}: {score_val}/{max_score} = {normalized_score:.2f}%")
        
        # Log paper average
        if paper_averages[paper_label]:
            avg = sum(paper_averages[paper_label]) / len(paper_averages[paper_label])
            logger.info(f"  {paper_label} Average: {avg:.2f}%")
        
        logger.info(f"  Processed {rows_processed} rows from {fname}")
    
    return paper_results

def process_caosce_station_files(files, raw_dir):
    """
    Process CAOSCE station files (existing functionality)
    Returns: caosce_results, station_max_scores, station_overall_averages
    """
    caosce_results = {}
    station_max_scores = {}
    station_overall_averages = {}
    all_exam_numbers = set()
    
    # Initialize all station keys
    station_keys = list(STATION_COLUMN_MAP.keys())
    
    for fname in sorted(files):
        path = os.path.join(raw_dir, fname)
        paper_type = detect_paper_type(fname)
        
        # Only process CAOSCE station files
        if paper_type != "CAOSCE_STATION":
            continue
            
        lower = fname.lower()
        station_key = None
        
        # Determine station type
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
            logger.warning(f"Could not determine station for {fname} – skipping")
            continue

        try:
            if fname.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str)
            else:
                df = pd.read_excel(path, dtype=str)
        except Exception as e:
            logger.error(f"Error reading {fname}: {e}")
            continue

        df.rename(columns=lambda c: str(c).strip(), inplace=True)

        username_col = find_username_col(df)
        fullname_col = find_fullname_col(df)
        grade_col, max_score = find_grade_column(df)
        viva_score_col = find_viva_score_col(df) if station_key == "viva" else None

        if grade_col:
            logger.info(f"  Found grade column: '{grade_col}' (max score: {max_score}) in {fname}")
            station_max_scores[station_key] = max_score
        else:
            logger.warning(f"  No grade column found in {fname}")
            station_max_scores[station_key] = 10.0

        # Remove unwanted columns
        for pattern in UNWANTED_COL_PATTERNS:
            df.drop(columns=[c for c in df.columns if re.search(pattern, str(c), flags=re.I)], 
                   inplace=True, errors="ignore")

        rows_added = 0
        station_scores = []
        
        for _, row in df.iterrows():
            # Skip overall average rows
            if is_overall_average_row(row, username_col, fullname_col):
                logger.info(f"  Skipping overall average row in {fname}")
                continue
                
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

            if exam_no:
                all_exam_numbers.add(exam_no)

            # Initialize student with all station keys if not exists
            if exam_no not in caosce_results:
                caosce_results[exam_no] = {
                    "MAT NO.": exam_no,
                    "FULL NAME": None,
                }
                # Initialize all station scores to 0.00
                for sk in station_keys:
                    caosce_results[exam_no][STATION_COLUMN_MAP[sk]] = 0.00

            if full_name and not caosce_results[exam_no]["FULL NAME"]:
                caosce_results[exam_no]["FULL NAME"] = full_name

            score_val = None
            if station_key == "viva" and viva_score_col:
                score_val = numeric_safe(row.get(viva_score_col))
            elif grade_col:
                score_val = numeric_safe(row.get(grade_col))

            if score_val is not None:
                out_col = STATION_COLUMN_MAP[station_key]
                caosce_results[exam_no][out_col] = round(score_val, 2)
                station_scores.append(score_val)

            rows_added += 1

        # Calculate average for this station
        if station_scores:
            station_avg = sum(station_scores) / len(station_scores)
            station_overall_averages[station_key] = round(station_avg, 2)
            logger.info(f"  Calculated station average: {station_avg:.2f}")

        logger.info(f"Processed {fname} → {rows_added} rows (Station: {station_key})")
    
    return caosce_results, station_max_scores, station_overall_averages, all_exam_numbers

def merge_results(caosce_results, paper_results, station_max_scores):
    """
    Merge CAOSCE and paper results into combined structure
    """
    combined_results = {}
    all_exam_numbers = set(caosce_results.keys()) | set(paper_results.keys())
    
    for exam_no in all_exam_numbers:
        combined_results[exam_no] = {
            "MAT NO.": exam_no,
            "FULL NAME": None,
            "PAPER I": 0.00,
            "PAPER II": 0.00,
            "CAOSCE": 0.00,
            "OVERALL AVERAGE": 0.00
        }
        
        # Add CAOSCE data
        if exam_no in caosce_results:
            combined_results[exam_no]["FULL NAME"] = caosce_results[exam_no]["FULL NAME"]
            
            # Calculate CAOSCE percentage
            total_score = 0
            total_max = sum(station_max_scores.values())
            for station_key, base_col in STATION_COLUMN_MAP.items():
                score = caosce_results[exam_no].get(base_col, 0) or 0
                total_score += score
            
            if total_max > 0:
                caosce_percentage = (total_score / total_max) * 100
                combined_results[exam_no]["CAOSCE"] = round(caosce_percentage, 2)
        
        # Add paper data  
        if exam_no in paper_results:
            paper_data = paper_results[exam_no]
            if not combined_results[exam_no]["FULL NAME"] and paper_data.get("FULL NAME"):
                combined_results[exam_no]["FULL NAME"] = paper_data["FULL NAME"]
                
            combined_results[exam_no]["PAPER I"] = paper_data.get("PAPER I", 0.00) or 0.00
            combined_results[exam_no]["PAPER II"] = paper_data.get("PAPER II", 0.00) or 0.00
        
        # Calculate overall average
        paper_i = combined_results[exam_no]["PAPER I"] or 0
        paper_ii = combined_results[exam_no]["PAPER II"] or 0  
        caosce_score = combined_results[exam_no]["CAOSCE"] or 0
        
        overall_avg = (paper_i + paper_ii + caosce_score) / 3
        combined_results[exam_no]["OVERALL AVERAGE"] = round(overall_avg, 2)
    
    return combined_results

def create_caosce_sheet(wb, df_caosce, college_config, station_max_scores, station_overall_averages):
    """
    Create the CAOSCE Results sheet (existing functionality)
    """
    ws = wb.create_sheet("CAOSCE Results", 0)
    
    # Write data to worksheet
    for r in dataframe_to_rows(df_caosce, index=False, header=True):
        ws.append(r)
    
    data_end_row = apply_excel_formatting(ws, df_caosce, college_config, "CAOSCE", station_max_scores)
    return ws, data_end_row

def create_combined_sheet(wb, df_combined, college_config):
    """
    Create the Combined Results sheet
    """
    ws = wb.create_sheet("Combined Results")
    
    # Write data to worksheet
    for r in dataframe_to_rows(df_combined, index=False, header=True):
        ws.append(r)
    
    data_end_row = apply_excel_formatting(ws, df_combined, college_config, "COMBINED")
    return ws, data_end_row

def apply_excel_formatting(ws, df, college_config, sheet_type, station_max_scores=None):
    """
    Apply consistent Excel formatting to worksheets
    """
    TITLE_ROWS = 6
    ws.insert_rows(1, TITLE_ROWS)
    header_row = TITLE_ROWS + 1

    last_col_letter = get_column_letter(ws.max_column)

    # Add college-specific logo
    logo_path = None
    base_logo_name = college_config["logo"]
    
    for ext in ['.png', '.jpg', '.jpeg']:
        potential_path = os.path.join(LOGO_BASE_PATH, base_logo_name.replace('.png', ext))
        if os.path.exists(potential_path):
            logo_path = potential_path
            break
    
    if not logo_path:
        exact_path = os.path.join(LOGO_BASE_PATH, base_logo_name)
        if os.path.exists(exact_path):
            logo_path = exact_path
    
    if logo_path:
        try:
            img = XLImage(logo_path)
            img.width = 120
            img.height = 120
            ws.add_image(img, "A1")
            logger.info(f"✓ Added logo: {os.path.basename(logo_path)}")
        except Exception as e:
            logger.warning(f"Could not add logo {logo_path}: {e}")

    # College-specific titles
    ws.merge_cells(f"B1:{last_col_letter}1")
    ws["B1"] = college_config["name"]
    ws["B1"].font = TITLE_FONT
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(f"B2:{last_col_letter}2")
    ws["B2"] = "PRE-COUNCIL EXAMINATION RESULT"
    ws["B2"].font = SUBTITLE_FONT
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24

    ws.merge_cells(f"B3:{last_col_letter}3")
    exam_type = "COMBINED" if sheet_type == "COMBINED" else "CAOSCE"
    ws["B3"] = f"{exam_type}_{CURRENT_YEAR}"
    ws["B3"].font = Font(bold=True, size=12, name="Calibri", color="1F4E78")
    ws["B3"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 20

    # DATE
    ws.merge_cells(f"B4:{last_col_letter}4")
    ws.cell(row=4, column=2, value=f"Date: {datetime.now().strftime('%d %B %Y')}")
    ws.cell(row=4, column=2).font = DATE_FONT
    ws.cell(row=4, column=2).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 22  # Increased from 20

    # CLASS - with more spacing from logo
    ws.merge_cells(f"A5:{last_col_letter}5")
    class_cell = ws.cell(row=5, column=1, value="CLASS: _________________________________________________________________________________________")
    class_cell.font = Font(bold=True, size=11, name="Calibri", color="1F4E78")
    class_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[5].height = 24  # Increased from 18 for more breathing room

    # Empty row for spacing before header
    ws.row_dimensions[6].height = 15  # Increased from 10 for better separation

    # Header styling
    for cell in ws[header_row]:
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        cell.fill = HEADER_FILL
    ws.row_dimensions[header_row].height = 20

    # Update MAT NO. header to college-specific label
    mat_no_idx = 2
    full_name_idx = 3
    ws.cell(row=header_row, column=mat_no_idx).value = college_config["mat_no_label"]
    ws.cell(row=header_row, column=mat_no_idx).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.cell(row=header_row, column=full_name_idx).alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Borders
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=header_row, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border

    ws.freeze_panes = f"A{header_row + 1}"

    # Determine column indices based on sheet type
    if sheet_type == "COMBINED":
        paper_i_idx = 4      # Column D: PAPER I/100
        paper_ii_idx = 5     # Column E: PAPER II/100
        caosce_idx = 6       # Column F: CAOSCE/100
        overall_idx = 7      # Column G: OVERALL AVERAGE/100
        
        # Format data rows for combined sheet
        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
            is_avg_row = (row[1].value == "OVERALL AVERAGE")
            
            for cell in row:
                if cell.column == 1:  # S/N
                    if is_avg_row:
                        cell.value = ""
                        cell.font = AVERAGE_FONT
                        cell.fill = AVERAGE_FILL
                    else:
                        cell.number_format = "0"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(size=10, name="Calibri")
                elif cell.column == mat_no_idx:  # MAT NO./EXAM NO.
                    if is_avg_row:
                        cell.font = AVERAGE_FONT
                        cell.fill = AVERAGE_FILL
                    else:
                        cell.font = Font(size=10, name="Calibri")
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                elif cell.column == full_name_idx:  # FULL NAME
                    if is_avg_row:
                        cell.value = ""
                        cell.fill = AVERAGE_FILL
                    else:
                        cell.font = Font(size=10, name="Calibri")
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                elif cell.column in [paper_i_idx, paper_ii_idx, caosce_idx]:  # Paper scores
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
                elif cell.column == overall_idx:  # OVERALL AVERAGE
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
    else:
        # CAOSCE sheet formatting (existing logic)
        score_cols_count = len([col for col in df.columns if "Score/" in col or "VIVA/" in col])
        station_col_indices = list(range(4, 4 + score_cols_count))
        total_col_idx = 4 + score_cols_count
        percent_col_idx = total_col_idx + 1

        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
            is_avg_row = (row[1].value == "OVERALL AVERAGE")
            
            for cell in row:
                if cell.column == 1:  # S/N
                    if is_avg_row:
                        cell.value = ""
                        cell.font = AVERAGE_FONT
                        cell.fill = AVERAGE_FILL
                    else:
                        cell.number_format = "0"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(size=10, name="Calibri")
                elif cell.column == mat_no_idx:  # MAT NO.
                    if is_avg_row:
                        cell.font = AVERAGE_FONT
                        cell.fill = AVERAGE_FILL
                    else:
                        cell.font = Font(size=10, name="Calibri")
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                elif cell.column == full_name_idx:  # FULL NAME
                    if is_avg_row:
                        cell.value = ""
                        cell.fill = AVERAGE_FILL
                    else:
                        cell.font = Font(size=10, name="Calibri")
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                elif cell.column in station_col_indices:  # Station scores
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
                elif cell.column in [total_col_idx, percent_col_idx]:  # Total and Percentage
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

    # Apply autofit columns
    data_end_row = ws.max_row
    apply_autofit_columns(ws, header_row, data_end_row)

    return data_end_row

# ---------------------------
# Main Processing Function
# ---------------------------

def process_files():
    """
    Main function to process all files and generate ONE workbook with TWO sheets
    """
    logger.info("Starting Enhanced CAOSCE Pre-Council Results Cleaning...")
    logger.info(f"Processing year: CAOSCE_{CURRENT_YEAR}")

    RAW_DIR = DEFAULT_RAW_DIR
    BASE_CLEAN_DIR = DEFAULT_CLEAN_DIR

    ts = datetime.now().strftime(TIMESTAMP_FMT)
    
    # Get all files in raw directory
    files = [f for f in os.listdir(RAW_DIR) if f.lower().endswith((".xlsx", ".xls", ".csv"))]

    if not files:
        logger.error(f"No raw files found in {RAW_DIR}")
        return

    logger.info(f"Found {len(files)} files to process")

    # Process CAOSCE station files
    caosce_results, station_max_scores, station_overall_averages, caosce_exam_numbers = process_caosce_station_files(files, RAW_DIR)
    
    # Process Paper I and II files
    paper_results = process_paper_files(files, RAW_DIR)
    
    if not caosce_results and not paper_results:
        logger.error("No valid data found in any files")
        return

    # Detect college using all exam numbers
    all_exam_numbers = caosce_exam_numbers | set(paper_results.keys())
    college_key, college_config = detect_college_from_exam_numbers(all_exam_numbers)
    
    logger.info(f"Detected college: {college_config['name']}")
    logger.info(f"Using logo: {college_config['logo']}")
    logger.info(f"Exam number label: {college_config['mat_no_label']}")

    # Create college-specific output directory
    output_dir = os.path.join(BASE_CLEAN_DIR, f"{college_config['output_prefix']}_COMBINED_{ts}")
    os.makedirs(output_dir, exist_ok=True)

    # Generate only ONE workbook with TWO sheets
    combined_output = generate_combined_output(caosce_results, paper_results, station_max_scores, station_overall_averages,
                                             college_config, output_dir, ts)

    # Print summary
    logger.info("\n" + "="*50)
    logger.info("PROCESSING COMPLETE")
    logger.info("="*50)
    
    if combined_output:
        logger.info(f"✓ Combined Results (2 sheets): {os.path.basename(combined_output)}")
    
    logger.info(f"📁 Output directory: {output_dir}")
    
    caosce_count = len(caosce_results) if caosce_results else 0
    paper_count = len(paper_results) if paper_results else 0
    combined_count = len(set(caosce_results.keys()) | set(paper_results.keys())) if caosce_results or paper_results else 0
    
    logger.info(f"📊 Students processed: CAOSCE={caosce_count}, Papers={paper_count}, Combined={combined_count}")

def generate_combined_output(caosce_results, paper_results, station_max_scores, station_overall_averages,
                           college_config, output_dir, timestamp):
    """
    Generate ONE workbook with TWO sheets: CAOSCE Results and Combined Results
    """
    # Merge results for combined sheet
    combined_results = merge_results(caosce_results, paper_results, station_max_scores)
    
    if not combined_results:
        logger.warning("No combined results to process")
        return None

    # Create DataFrame for combined results
    df_combined = pd.DataFrame.from_dict(combined_results, orient="index")
    
    # Reorder columns
    column_order = ["MAT NO.", "FULL NAME", "PAPER I", "PAPER II", "CAOSCE", "OVERALL AVERAGE"]
    df_combined = df_combined[column_order]
    
    # Sort by exam number
    df_combined["__sort"] = pd.to_numeric(df_combined["MAT NO."].str.extract(r'(\d+)')[0], errors='coerce')
    df_combined.sort_values(["__sort", "MAT NO."], inplace=True)
    df_combined.drop(columns=["__sort"], inplace=True)
    df_combined.reset_index(drop=True, inplace=True)

    df_combined.insert(0, "S/N", range(1, len(df_combined) + 1))

    # Rename columns for display with /100 notation
    df_combined.rename(columns={
        "PAPER I": "PAPER I/100",
        "PAPER II": "PAPER II/100",
        "CAOSCE": "CAOSCE/100",
        "OVERALL AVERAGE": "OVERALL AVERAGE/100"
    }, inplace=True)

    # Calculate overall averages for the average row
    paper_i_avg = df_combined["PAPER I/100"].mean() if not df_combined["PAPER I/100"].isna().all() else 0
    paper_ii_avg = df_combined["PAPER II/100"].mean() if not df_combined["PAPER II/100"].isna().all() else 0
    caosce_avg = df_combined["CAOSCE/100"].mean() if not df_combined["CAOSCE/100"].isna().all() else 0
    overall_avg = (paper_i_avg + paper_ii_avg + caosce_avg) / 3

    # Add overall average row
    avg_row = {
        "S/N": "",
        "MAT NO.": "OVERALL AVERAGE", 
        "FULL NAME": "",
        "PAPER I/100": round(paper_i_avg, 2),
        "PAPER II/100": round(paper_ii_avg, 2),
        "CAOSCE/100": round(caosce_avg, 2),
        "OVERALL AVERAGE/100": round(overall_avg, 2)
    }
    
    df_combined = pd.concat([df_combined, pd.DataFrame([avg_row])], ignore_index=True)

    # Calculate statistics for documentation
    total_students = len(df_combined) - 1  # Exclude average row
    student_overall = df_combined["OVERALL AVERAGE/100"].iloc[:-1]  # Exclude average row
    
    avg_percentage = round(student_overall.mean(), 1) if total_students > 0 else 0
    highest_percentage = round(student_overall.max(), 1) if total_students > 0 else 0
    lowest_percentage = round(student_overall.min(), 1) if total_students > 0 else 0

    # Save combined output - only ONE file
    output_basename = f"{college_config['output_prefix']}_PRE_COUNCIL_CLEANED"
    out_xlsx = os.path.join(output_dir, f"{output_basename}_{timestamp}.xlsx")

    # Create Excel workbook with both sheets
    wb = Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Add CAOSCE sheet if we have CAOSCE data
    if caosce_results:
        caosce_df = generate_caosce_dataframe(caosce_results, station_max_scores, station_overall_averages)
        ws_caosce, data_end_row_caosce = create_caosce_sheet(wb, caosce_df, college_config, station_max_scores, station_overall_averages)
        
        # Add documentation to CAOSCE sheet
        total_max_score_caosce = sum(station_max_scores.values())
        create_document_sections(
            ws_caosce, len(caosce_df) - 1, avg_percentage, highest_percentage, lowest_percentage,
            total_max_score_caosce, data_end_row_caosce, college_config, "CAOSCE"
        )
    
    # Add Combined sheet
    ws_combined, data_end_row_combined = create_combined_sheet(wb, df_combined, college_config)
    
    # Add documentation to Combined sheet
    create_document_sections(
        ws_combined, total_students, avg_percentage, highest_percentage, lowest_percentage,
        300, data_end_row_combined, college_config, "COMBINED"
    )
    
    # Save the workbook
    wb.save(out_xlsx)
    
    logger.info(f"✓ Saved combined results: {os.path.basename(out_xlsx)}")
    return out_xlsx

def generate_caosce_dataframe(caosce_results, station_max_scores, station_overall_averages):
    """
    Generate DataFrame for CAOSCE results (helper function for combined output)
    """
    # Create a copy to avoid modifying original data
    caosce_results = copy.deepcopy(caosce_results)
    
    # Similar logic to generate_caosce_output but returns DataFrame instead of saving file
    score_cols = []
    for station_key in ["procedure_station_one", "procedure_station_three", "procedure_station_five",
                        "question_station_two", "question_station_four", "question_station_six", "viva"]:
        base_col = STATION_COLUMN_MAP[station_key]
        max_score = station_max_scores.get(station_key, 10.0)
        display_name = STATION_DISPLAY_NAMES[station_key]
        if max_score == int(max_score):
            col_name = f"{display_name}/{int(max_score)}"
        else:
            col_name = f"{display_name}/{max_score}"
        score_cols.append(col_name)

    # Create a new dictionary with the display column names
    processed_results = {}
    for exam_no, student_data in caosce_results.items():
        processed_results[exam_no] = {
            "MAT NO.": student_data["MAT NO."],
            "FULL NAME": student_data["FULL NAME"]
        }
        
        # Add station scores with display names
        for i, station_key in enumerate(["procedure_station_one", "procedure_station_three", "procedure_station_five",
                                         "question_station_two", "question_station_four", "question_station_six", "viva"]):
            base_col = STATION_COLUMN_MAP[station_key]
            processed_results[exam_no][score_cols[i]] = student_data.get(base_col, 0.00)

    base_cols = ["MAT NO.", "FULL NAME"] + score_cols
    df_out = pd.DataFrame.from_dict(processed_results, orient="index")
    
    # Ensure all required columns exist
    for col in base_cols:
        if col not in df_out.columns:
            df_out[col] = 0.00

    df_out = df_out[base_cols]

    df_out["FULL NAME"] = df_out.groupby("MAT NO.")["FULL NAME"].transform(
        lambda x: x.fillna(method='ffill').fillna(method='bfill')
    )

    # Sort by exam number
    df_out["__sort"] = pd.to_numeric(df_out["MAT NO."].str.extract(r'(\d+)')[0], errors='coerce')
    df_out.sort_values(["__sort", "MAT NO."], inplace=True)
    df_out.drop(columns=["__sort"], inplace=True)
    df_out.reset_index(drop=True, inplace=True)

    df_out.insert(0, "S/N", range(1, len(df_out) + 1))

    df_out[score_cols] = df_out[score_cols].apply(pd.to_numeric, errors="coerce")
    for col in score_cols:
        df_out[col] = df_out[col].apply(lambda x: round(x, 2) if pd.notna(x) else 0.00)

    total_max_score = sum(station_max_scores.values())
    df_out[f"Total Raw Score/{total_max_score}"] = df_out[score_cols].sum(axis=1)
    total_col_name = f"Total Raw Score/{total_max_score}"
    
    df_out[total_col_name] = df_out[total_col_name].apply(lambda x: round(x, 2) if pd.notna(x) else 0.00)
    
    df_out["Percentage (%)"] = (df_out[total_col_name] / total_max_score * 100)
    df_out["Percentage (%)"] = df_out["Percentage (%)"].apply(lambda x: int(round(x, 0)) if pd.notna(x) else 0)

    final_display_cols = ["S/N", "MAT NO.", "FULL NAME"] + score_cols + [total_col_name, "Percentage (%)"]
    df_out = df_out[final_display_cols]

    # Add overall average row
    avg_row = {
        "S/N": "",
        "MAT NO.": "OVERALL AVERAGE",
        "FULL NAME": "",
    }
    
    for i, station_key in enumerate(["procedure_station_one", "procedure_station_three", "procedure_station_five",
                                    "question_station_two", "question_station_four", "question_station_six", "viva"]):
        col_name = score_cols[i]
        avg_row[col_name] = station_overall_averages.get(station_key, 0.00)
    
    overall_total = sum([station_overall_averages.get(station_key, 0.00) for station_key in station_overall_averages])
    avg_row[total_col_name] = round(overall_total, 2)
    avg_row["Percentage (%)"] = int(round(overall_total / total_max_score * 100, 0))

    df_out = pd.concat([df_out, pd.DataFrame([avg_row])], ignore_index=True)

    return df_out

if __name__ == "__main__":
    process_files()