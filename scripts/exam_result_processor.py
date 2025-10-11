#!/usr/bin/env python3
"""
exam_result_processor.py

Complete script with proper semester ordering and Previous GPA tracking.
"""

from openpyxl.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os
import sys
import re
import pandas as pd
from datetime import datetime
import platform
import difflib
import math
import glob

# PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ----------------------------
# Configuration
# ----------------------------
BASE_DIR = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/EXAMS_INTERNAL"
ND_COURSES_DIR = os.path.join(BASE_DIR, "ND-COURSES")
DEFAULT_PASS_THRESHOLD = 50.0
TIMESTAMP_FMT = "%Y-%m-%d_%H%M%S"

DEFAULT_LOGO_PATH = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "launcher", "static", "logo.png"))

NAME_WIDTH_CAP = 40

# Define semester processing order
SEMESTER_ORDER = [
    "ND-FIRST-YEAR-FIRST-SEMESTER",
    "ND-FIRST-YEAR-SECOND-SEMESTER", 
    "ND-SECOND-YEAR-FIRST-SEMESTER",
    "ND-SECOND-YEAR-SECOND-SEMESTER"
]

# ----------------------------
# Utilities
# ----------------------------
def normalize_path(path: str) -> str:
    """Normalize user paths for Windows/WSL/Linux."""
    path = os.path.expanduser(path)
    path = os.path.normpath(path)
    if platform.system().lower() == "linux" and path.startswith("C:\\"):
        path = "/mnt/" + path[0].lower() + path[2:].replace("\\", "/")
    if platform.system().lower() == "linux" and path.startswith("c:\\"):
        path = "/mnt/" + path[0].lower() + path[2:].replace("\\", "/")
    if path.startswith("/c/") and os.path.exists("/mnt/c"):
        path = path.replace("/c/", "/mnt/c/", 1)
    return os.path.abspath(path)

def normalize_course_name(name):
    """Simple normalization for course title matching."""
    return re.sub(r'\s+', ' ', str(name).strip().lower()).replace('coomunication', 'communication')

def normalize_for_matching(s):
    if s is None:
        return ""
    s = str(s).lower()
    s = re.sub(r'\b1st\b', 'first', s)
    s = re.sub(r'\b2nd\b', 'second', s)
    s = re.sub(r'\b3rd\b', 'third', s)
    s = re.sub(r'[^a-z0-9\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

# ----------------------------
# Grade and GPA calculation
# ----------------------------
def get_grade(score):
    """Convert numeric score to letter grade - single letter only."""
    try:
        score = float(score)
        if score >= 70:
            return 'A'
        elif score >= 60:
            return 'B'
        elif score >= 50:
            return 'C'
        elif score >= 45:
            return 'D'
        elif score >= 40:
            return 'E'
        else:
            return 'F'
    except:
        return 'F'

def get_grade_point(score):
    """Convert score to grade point for GPA calculation."""
    try:
        score = float(score)
        if score >= 70:
            return 5.0  # A
        elif score >= 60:
            return 4.0  # B
        elif score >= 50:
            return 3.0  # C
        elif score >= 45:
            return 2.0  # D
        elif score >= 40:
            return 1.0  # E
        else:
            return 0.0  # F
    except:
        return 0.0

# ----------------------------
# Load Course Data
# ----------------------------
def load_course_data():
    """
    Reads course-code-creditUnit.xlsx and returns:
      (semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles)
    """
    course_file = os.path.join(ND_COURSES_DIR, "course-code-creditUnit.xlsx")
    print(f"Loading course data from: {course_file}")
    if not os.path.exists(course_file):
        raise FileNotFoundError(f"Course file not found: {course_file}")

    xl = pd.ExcelFile(course_file)
    semester_course_maps = {}
    semester_credit_units = {}
    semester_lookup = {}
    semester_course_titles = {}  # code -> title mapping

    for sheet in xl.sheet_names:
        df = pd.read_excel(course_file, sheet_name=sheet, engine='openpyxl', header=0)
        df.columns = [str(c).strip() for c in df.columns]
        expected = ['COURSE CODE', 'COURSE TITLE', 'CU']
        if not all(col in df.columns for col in expected):
            print(f"Warning: sheet '{sheet}' missing expected columns {expected} — skipped")
            continue
        dfx = df.dropna(subset=['COURSE CODE', 'COURSE TITLE'])
        dfx = dfx[~dfx['COURSE CODE'].astype(str).str.contains('TOTAL', case=False, na=False)]
        valid_mask = dfx['CU'].astype(str).str.replace('.', '', regex=False).str.isdigit()
        dfx = dfx[valid_mask]
        if dfx.empty:
            print(f"Warning: sheet '{sheet}' has no valid rows after cleaning — skipped")
            continue
        codes = dfx['COURSE CODE'].astype(str).str.strip().tolist()
        titles = dfx['COURSE TITLE'].astype(str).str.strip().tolist()
        cus = dfx['CU'].astype(float).astype(int).tolist()

        semester_course_maps[sheet] = dict(zip(titles, codes))
        semester_credit_units[sheet] = dict(zip(codes, cus))
        semester_course_titles[sheet] = dict(zip(codes, titles))

        # Create multiple lookup variations for flexible matching
        norm = normalize_for_matching(sheet)
        semester_lookup[norm] = sheet
        
        # Add variations without "ND-" prefix
        norm_no_nd = norm.replace('nd-', '').replace('nd ', '')
        semester_lookup[norm_no_nd] = sheet
        
        # Add variations with different separators
        norm_hyphen = norm.replace('-', ' ')
        semester_lookup[norm_hyphen] = sheet
        
        norm_space = norm.replace(' ', '-')
        semester_lookup[norm_space] = sheet

    if not semester_course_maps:
        raise ValueError("No course data loaded from course workbook")
    print(f"Loaded course sheets: {list(semester_course_maps.keys())}")
    return semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles

# ----------------------------
# Helper functions
# ----------------------------
def detect_semester_from_filename(filename):
    """
    Detect semester from filename.
    Returns: (semester_key, year, semester_num, level_display, semester_display, set_code)
    """
    filename_upper = filename.upper()
    
    # Map filename patterns to actual course sheet names
    if 'FIRST-YEAR-FIRST-SEMESTER' in filename_upper or 'FIRST_YEAR_FIRST_SEMESTER' in filename_upper:
        return "ND-FIRST-YEAR-FIRST-SEMESTER", 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
    elif 'FIRST-YEAR-SECOND-SEMESTER' in filename_upper or 'FIRST_YEAR_SECOND_SEMESTER' in filename_upper:
        return "ND-FIRST-YEAR-SECOND-SEMESTER", 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI"
    elif 'SECOND-YEAR-FIRST-SEMESTER' in filename_upper or 'SECOND_YEAR_FIRST_SEMESTER' in filename_upper:
        return "ND-SECOND-YEAR-FIRST-SEMESTER", 2, 1, "YEAR TWO", "FIRST SEMESTER", "NDII"
    elif 'SECOND-YEAR-SECOND-SEMESTER' in filename_upper or 'SECOND_YEAR_SECOND_SEMESTER' in filename_upper:
        return "ND-SECOND-YEAR-SECOND-SEMESTER", 2, 2, "YEAR TWO", "SECOND SEMESTER", "NDII"
    elif 'FIRST' in filename_upper and 'SECOND' not in filename_upper:
        return "ND-FIRST-YEAR-FIRST-SEMESTER", 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
    elif 'SECOND' in filename_upper:
        return "ND-FIRST-YEAR-SECOND-SEMESTER", 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI"
    else:
        # Default fallback
        print(f"⚠️ Could not detect semester from filename: {filename}, defaulting to ND-FIRST-YEAR-FIRST-SEMESTER")
        return "ND-FIRST-YEAR-FIRST-SEMESTER", 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"

def get_semester_display_info(semester_key):
    """
    Get display information for a given semester key.
    Returns: (year, semester_num, level_display, semester_display, set_code)
    """
    semester_lower = semester_key.lower()
    
    if 'first-year-first-semester' in semester_lower:
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
    elif 'first-year-second-semester' in semester_lower:
        return 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI"
    elif 'second-year-first-semester' in semester_lower:
        return 2, 1, "YEAR TWO", "FIRST SEMESTER", "NDII"
    elif 'second-year-second-semester' in semester_lower:
        return 2, 2, "YEAR TWO", "SECOND SEMESTER", "NDII"
    elif 'first' in semester_lower and 'second' not in semester_lower:
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
    elif 'second' in semester_lower:
        return 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI"
    else:
        # Default to first semester, first year
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"

def match_semester_from_filename(fname, semester_lookup):
    """Match semester using the lookup table with flexible matching."""
    fn = normalize_for_matching(fname)
    
    # Try exact matches first
    for norm, sheet in semester_lookup.items():
        if norm in fn:
            return sheet
    
    # Try close matches
    keys = list(semester_lookup.keys())
    best = difflib.get_close_matches(fn, keys, n=1, cutoff=0.55)
    if best:
        return semester_lookup[best[0]]
    
    # Fallback to filename-based detection
    sem, _, _, _, _, _ = detect_semester_from_filename(fname)
    return sem

def find_column_by_names(df, candidate_names):
    norm_map = {col: re.sub(r'\s+', ' ', str(col).strip().lower()) for col in df.columns}
    candidates = [re.sub(r'\s+', ' ', c.strip().lower()) for c in candidate_names]
    for cand in candidates:
        for col, ncol in norm_map.items():
            if ncol == cand:
                return col
    return None

def load_previous_gpas_from_processed_files(output_dir, current_semester_key, timestamp):
    """
    Load previous GPA data from previously processed mastersheets in the same run.
    Returns dict: {exam_number: previous_gpa}
    """
    previous_gpas = {}
    
    print(f"\n🔍 LOADING PREVIOUS GPA for: {current_semester_key}")
    
    # Determine previous semester based on current
    current_year, current_semester_num, _, _, _ = get_semester_display_info(current_semester_key)
    
    if current_semester_num == 1 and current_year == 1:
        # First semester of first year - no previous GPA
        print("📊 First semester of first year - no previous GPA available")
        return previous_gpas
    elif current_semester_num == 2 and current_year == 1:
        # Second semester of first year - look for first semester of first year
        prev_semester = "ND-FIRST-YEAR-FIRST-SEMESTER"
    elif current_semester_num == 1 and current_year == 2:
        # First semester of second year - look for second semester of first year
        prev_semester = "ND-FIRST-YEAR-SECOND-SEMESTER"
    elif current_semester_num == 2 and current_year == 2:
        # Second semester of second year - look for first semester of second year
        prev_semester = "ND-SECOND-YEAR-FIRST-SEMESTER"
    else:
        print(f"⚠️ Unknown semester combination: Year {current_year}, Semester {current_semester_num}")
        return previous_gpas
    
    print(f"🔍 Looking for previous GPA data from: {prev_semester}")
    
    # Look for the mastersheet file from the previous semester in the same timestamp directory
    mastersheet_pattern = os.path.join(output_dir, f"ND_RESULT-{timestamp}", f"mastersheet_{timestamp}.xlsx")
    print(f"🔍 Checking for mastersheet: {mastersheet_pattern}")
    
    if os.path.exists(mastersheet_pattern):
        print(f"✅ Found mastersheet: {mastersheet_pattern}")
        try:
            # Read the Excel file properly, skipping the header rows that contain merged cells
            df = pd.read_excel(mastersheet_pattern, sheet_name=prev_semester, header=5)  # Skip first 5 rows
            
            print(f"📋 Columns in {prev_semester}: {df.columns.tolist()}")
            
            # Find the actual column names by checking for exam number and GPA columns
            exam_col = None
            gpa_col = None
            
            for col in df.columns:
                col_str = str(col).upper().strip()
                if 'EXAM' in col_str or 'REG' in col_str or 'NUMBER' in col_str:
                    exam_col = col
                elif 'GPA' in col_str:
                    gpa_col = col
            
            if exam_col and gpa_col:
                print(f"✅ Found exam column: '{exam_col}', GPA column: '{gpa_col}'")
                
                gpas_loaded = 0
                for idx, row in df.iterrows():
                    exam_no = str(row[exam_col]).strip()
                    gpa = row[gpa_col]
                    
                    if pd.notna(gpa) and pd.notna(exam_no) and exam_no != 'nan' and exam_no != '':
                        try:
                            previous_gpas[exam_no] = float(gpa)
                            gpas_loaded += 1
                            if gpas_loaded <= 5:  # Show first 5 for debugging
                                print(f"📝 Loaded GPA: {exam_no} → {gpa}")
                        except (ValueError, TypeError):
                            continue
                
                print(f"✅ Loaded previous GPAs for {gpas_loaded} students from {prev_semester}")
                
                if gpas_loaded > 0:
                    # Show sample of loaded GPAs for verification
                    sample_gpas = list(previous_gpas.items())[:3]
                    print(f"📊 Sample GPAs loaded: {sample_gpas}")
                else:
                    print(f"⚠️ No valid GPA data found in {prev_semester}")
            else:
                print(f"❌ Could not find required columns in {prev_semester}")
                if not exam_col:
                    print("❌ Could not find exam number column")
                if not gpa_col:
                    print("❌ Could not find GPA column")
                    
        except Exception as e:
            print(f"⚠️ Could not read mastersheet: {str(e)}")
            import traceback
            traceback.print_exc()
    else:
        print(f"❌ Mastersheet not found: {mastersheet_pattern}")
        # Check if directory exists
        dir_path = os.path.dirname(mastersheet_pattern)
        if os.path.exists(dir_path):
            print(f"📁 Directory contents: {os.listdir(dir_path)}")
        else:
            print(f"📁 Directory not found: {dir_path}")
    
    print(f"📊 FINAL: Loaded {len(previous_gpas)} previous GPAs")
    return previous_gpas

def get_cumulative_gpa(current_gpa, previous_gpa, current_credits, previous_credits):
    """
    Calculate cumulative GPA based on current and previous semester performance.
    """
    if previous_gpa is None:
        return current_gpa
    
    # For simplicity, we'll assume equal credit weights if not provided
    if current_credits is None or previous_credits is None:
        return round((current_gpa + previous_gpa) / 2, 2)
    
    total_points = (current_gpa * current_credits) + (previous_gpa * previous_credits)
    total_credits = current_credits + previous_credits
    return round(total_points / total_credits, 2) if total_credits > 0 else 0.0

# ----------------------------
# PDF Generation - Individual Student Report
# ----------------------------
def generate_individual_student_pdf(mastersheet_df, out_pdf_path, semester_key, logo_path=None, 
                                   prev_mastersheet_df=None, filtered_credit_units=None, 
                                   ordered_codes=None, course_titles_map=None, previous_gpas=None):
    """
    Create a PDF with one page per student matching the sample format exactly.
    """
    doc = SimpleDocTemplate(out_pdf_path, pagesize=A4, rightMargin=40, leftMargin=40, 
                           topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    
    # Custom styles
    header_style = ParagraphStyle(
        'CustomHeader',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_CENTER,
        spaceAfter=2
    )
    
    main_header_style = ParagraphStyle(
        'MainHeader',
        parent=styles['Normal'],
        fontSize=16,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold',
        spaceAfter=6,
        textColor=colors.HexColor("#800080")
    )
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Normal'],
        fontSize=12,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold',
        spaceAfter=4
    )
    
    subtitle_style = ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_CENTER,
        spaceAfter=10,
        textColor=colors.red
    )
    
    # Left alignment style for course code and title
    left_align_style = ParagraphStyle(
        'LeftAlign',
        parent=styles['Normal'],
        fontSize=9,
        alignment=TA_LEFT,
        leftIndent=4
    )
    
    center_align_style = ParagraphStyle(
        'CenterAlign',
        parent=styles['Normal'],
        fontSize=9,
        alignment=TA_CENTER
    )
    
    elems = []

    for idx, r in mastersheet_df.iterrows():
        # Logo and header
        logo_img = None
        if logo_path and os.path.exists(logo_path):
            try:
                logo_img = Image(logo_path, width=0.8*inch, height=0.8*inch)
            except Exception as e:
                print(f"Warning: Could not load logo: {e}")
        
        # Header table with logo and title
        if logo_img:
            header_data = [[logo_img, Paragraph("FCT COLLEGE OF NURSING SCIENCES", main_header_style)]]
            header_table = Table(header_data, colWidths=[1.0*inch, 5.0*inch])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (0,0), 'LEFT'),
                ('ALIGN', (1,0), (1,0), 'CENTER'),
            ]))
            elems.append(header_table)
        else:
            elems.append(Paragraph("FCT COLLEGE OF NURSING SCIENCES", main_header_style))
        
        # Address and contact info
        elems.append(Paragraph("P.O.Box 507, Gwagwalada-Abuja, Nigeria", header_style))
        elems.append(Paragraph("<b>EXAMINATIONS OFFICE</b>", header_style))
        elems.append(Paragraph("fctsonexamsoffice@gmail.com", header_style))
        
        elems.append(Spacer(1, 8))
        elems.append(Paragraph("STUDENT'S ACADEMIC PROGRESS REPORT", title_style))
        elems.append(Paragraph("(THIS IS NOT A TRANSCRIPT)", subtitle_style))
        
        elems.append(Spacer(1, 8))
        
        # Student particulars - SEPARATE FROM PASSPORT PHOTO
        exam_no = str(r.get("EXAMS NUMBER", r.get("REG. No", "")))
        student_name = str(r.get("NAME", ""))
        
        # Determine level and semester using the new function
        year, semester_num, level_display, semester_display, set_code = get_semester_display_info(semester_key)
        
        # Create two tables: one for student particulars, one for passport photo
        particulars_data = [
            [Paragraph("<b>STUDENT'S PARTICULARS</b>", styles['Normal'])],
            [Paragraph("<b>NAME:</b>", styles['Normal']), student_name],
            [Paragraph("<b>LEVEL OF<br/>STUDY:</b>", styles['Normal']), level_display, 
             Paragraph("<b>SEMESTER:</b>", styles['Normal']), semester_display],
            [Paragraph("<b>REG NO.</b>", styles['Normal']), exam_no, 
             Paragraph("<b>SET:</b>", styles['Normal']), set_code],
        ]
        
        particulars_table = Table(particulars_data, colWidths=[1.2*inch, 2.3*inch, 0.8*inch, 1.5*inch])
        particulars_table.setStyle(TableStyle([
            ('SPAN', (0,0), (3,0)),
            ('SPAN', (1,1), (3,1)),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        
        # Passport photo table (separate box)
        passport_data = [
            [Paragraph("Affix Recent<br/>Passport<br/>Photograph", styles['Normal'])]
        ]
        
        passport_table = Table(passport_data, colWidths=[1.5*inch], rowHeights=[1.2*inch])
        passport_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('FONTSIZE', (0,0), (-1,-1), 9),
        ]))
        
        # Create a combined table with particulars and passport side by side
        combined_data = [
            [particulars_table, passport_table]
        ]
        
        combined_table = Table(combined_data, colWidths=[5.8*inch, 1.5*inch])
        combined_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        
        elems.append(combined_table)
        elems.append(Spacer(1, 12))
        
        # Semester result header
        elems.append(Paragraph("<b>SEMESTER RESULT</b>", title_style))
        elems.append(Spacer(1, 6))
        
        # Course results table - LEFT-ALIGNED CODE AND TITLE
        course_data = [[Paragraph("<b>S/N</b>", styles['Normal']), 
                       Paragraph("<b>CODE</b>", styles['Normal']), 
                       Paragraph("<b>COURSE TITLE</b>", styles['Normal']), 
                       Paragraph("<b>UNITS</b>", styles['Normal']), 
                       Paragraph("<b>SCORE</b>", styles['Normal']), 
                       Paragraph("<b>GRADE</b>", styles['Normal'])]]
        
        sn = 1
        total_grade_points = 0.0
        total_units = 0
        total_units_passed = 0
        total_units_failed = 0
        
        for code in ordered_codes if ordered_codes else []:
            score = r.get(code)
            if pd.isna(score) or score == "":
                continue
            
            try:
                score_val = float(score)
                score_display = str(int(round(score_val)))
                grade = get_grade(score_val)
                grade_point = get_grade_point(score_val)
            except:
                score_display = str(score)
                grade = "F"
                grade_point = 0.0
            
            cu = filtered_credit_units.get(code, 0) if filtered_credit_units else 0
            
            # Get course title
            course_title = course_titles_map.get(code, code) if course_titles_map else code
            
            # Calculate weighted grade points and unit counts
            total_grade_points += grade_point * cu
            total_units += cu
            
            # Track passed/failed units
            if score_val >= DEFAULT_PASS_THRESHOLD:
                total_units_passed += cu
            else:
                total_units_failed += cu
            
            # LEFT-ALIGNED course code and title
            course_data.append([
                Paragraph(str(sn), center_align_style), 
                Paragraph(code, left_align_style),
                Paragraph(course_title, left_align_style),
                Paragraph(str(cu), center_align_style), 
                Paragraph(score_display, center_align_style), 
                Paragraph(grade, center_align_style)
            ])
            sn += 1
        
        course_table = Table(course_data, colWidths=[0.4*inch, 0.7*inch, 2.8*inch, 0.6*inch, 0.6*inch, 0.6*inch])
        course_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E0E0E0")),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 9),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('ALIGN', (1,1), (2,-1), 'LEFT'),
        ]))
        elems.append(course_table)
        elems.append(Spacer(1, 14))
        
        # Calculate current semester GPA
        current_gpa = round(total_grade_points / total_units, 2) if total_units > 0 else 0.0
        
        # Get previous GPA if available
        exam_no = str(r.get("EXAMS NUMBER", "")).strip()
        previous_gpa = previous_gpas.get(exam_no, None) if previous_gpas else None
        
        print(f"📊 PDF GENERATION for {exam_no}:")
        print(f"   Current GPA: {current_gpa}")
        print(f"   Previous GPA available: {previous_gpa is not None}")
        if previous_gpa is not None:
            print(f"   Previous GPA value: {previous_gpa}")
        
        # Get values from dataframe
        tcpe = round(total_grade_points, 1)
        tcup = total_units_passed
        tcuf = total_units_failed
        remarks = str(r.get("REMARKS", ""))
        
        # Summary section - FIXED LOGIC
        summary_data = [
            [Paragraph("<b>SUMMARY</b>", styles['Normal']), "", "", ""],
            [Paragraph("<b>TCPE:</b>", styles['Normal']), str(tcpe), 
             Paragraph("<b>CURRENT GPA:</b>", styles['Normal']), str(current_gpa)],
        ]
        
        # Add previous GPA if available (from first year second semester upward)
        if previous_gpa is not None:
            print(f"✅ ADDING PREVIOUS GPA to PDF: {previous_gpa}")
            summary_data.append([
                Paragraph("<b>TCUP:</b>", styles['Normal']), str(tcup),
                Paragraph("<b>PREVIOUS GPA:</b>", styles['Normal']), str(previous_gpa)
            ])
            summary_data.append([
                Paragraph("<b>TCUF:</b>", styles['Normal']), str(tcuf), "", ""
            ])
        else:
            print(f"❌ NO PREVIOUS GPA available for {exam_no}")
            summary_data.append([
                Paragraph("<b>TCUP:</b>", styles['Normal']), str(tcup), "", ""
            ])
            summary_data.append([
                Paragraph("<b>TCUF:</b>", styles['Normal']), str(tcuf), "", ""
            ])
            
        summary_data.append([
            Paragraph("<b>REMARKS:</b>", styles['Normal']), remarks, "", ""
        ])
        
        summary_table = Table(summary_data, colWidths=[1.5*inch, 1.0*inch, 1.5*inch, 1.0*inch])
        summary_table.setStyle(TableStyle([
            ('SPAN', (0,0), (3,0)),
            ('SPAN', (1,len(summary_data)-1), (3,len(summary_data)-1)),
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (3,0), colors.HexColor("#E0E0E0")),
            ('ALIGN', (0,0), (3,0), 'CENTER'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        elems.append(summary_table)
        elems.append(Spacer(1, 25))
        
        # Signature section
        sig_data = [
            ["", ""],
            ["____________________", "____________________"],
            [Paragraph("<b>EXAMS SECRETARY</b>", ParagraphStyle('SigStyle', parent=styles['Normal'], 
                                                                fontSize=10, alignment=TA_CENTER)), 
             Paragraph("<b>V.P. ACADEMICS</b>", ParagraphStyle('SigStyle', parent=styles['Normal'], 
                                                              fontSize=10, alignment=TA_CENTER))]
        ]
        
        sig_table = Table(sig_data, colWidths=[3.0*inch, 3.0*inch])
        sig_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ]))
        elems.append(sig_table)
        
        # Page break for next student
        if idx < len(mastersheet_df) - 1:
            elems.append(PageBreak())
    
    doc.build(elems)
    print(f"✅ Individual student PDF written: {out_pdf_path}")

# ----------------------------
# Main file processing
# ----------------------------
def process_semester_files(semester_key, raw_files, raw_dir, clean_dir, ts, pass_threshold, 
                          semester_course_maps, semester_credit_units, semester_lookup, 
                          semester_course_titles, logo_path, previous_gpas=None):
    """
    Process all files for a specific semester.
    """
    print(f"\n{'='*60}")
    print(f"PROCESSING SEMESTER: {semester_key}")
    print(f"{'='*60}")
    
    # Filter files for this semester
    semester_files = []
    for rf in raw_files:
        detected_sem, _, _, _, _, _ = detect_semester_from_filename(rf)
        if detected_sem == semester_key:
            semester_files.append(rf)
    
    if not semester_files:
        print(f"⚠️ No files found for semester {semester_key}")
        return None
    
    print(f"📁 Found {len(semester_files)} files for {semester_key}: {semester_files}")
    
    # Process each file for this semester
    for rf in semester_files:
        raw_path = os.path.join(raw_dir, rf)
        print(f"\n📄 Processing: {rf}")
        
        try:
            # Load previous GPAs for this specific semester
            current_previous_gpas = load_previous_gpas_from_processed_files(clean_dir, semester_key, ts) if previous_gpas is None else previous_gpas
            
            # Process the file
            result = process_single_file(raw_path, clean_dir, ts, pass_threshold, semester_course_maps, 
                                       semester_credit_units, semester_lookup, semester_course_titles, 
                                       logo_path, semester_key, current_previous_gpas)
            
            if result is not None:
                print(f"✅ Successfully processed {rf}")
            else:
                print(f"❌ Failed to process {rf}")
                
        except Exception as e:
            print(f"❌ Error processing {rf}: {e}")
            import traceback
            traceback.print_exc()

def process_single_file(path, output_dir, ts, pass_threshold, semester_course_maps, semester_credit_units, 
                       semester_lookup, semester_course_titles, logo_path, semester_key, previous_gpas):
    """
    Process a single raw file and produce mastersheet Excel and PDFs.
    """
    fname = os.path.basename(path)
    
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        print(f"Error opening excel {path}: {e}")
        return None

    expected_sheets = ['CA', 'OBJ', 'EXAM']
    dfs = {}
    for s in expected_sheets:
        if s in xl.sheet_names:
            dfs[s] = pd.read_excel(path, sheet_name=s, dtype=str)
    if not dfs:
        print("No CA/OBJ/EXAM sheets detected — skipping file.")
        return None

    # Use the provided semester key
    sem = semester_key
    year, semester_num, level_display, semester_display, set_code = get_semester_display_info(sem)
    print(f"📁 Processing: {level_display} - {semester_display} - Set: {set_code}")
    print(f"📊 Using course sheet: {sem}")

    print(f"📊 Previous GPAs provided: {len(previous_gpas)} students")
    
    # Check if semester exists in course maps
    if sem not in semester_course_maps:
        print(f"❌ Semester '{sem}' not found in course data. Available semesters: {list(semester_course_maps.keys())}")
        return None
    
    course_map = semester_course_maps[sem]
    credit_units = semester_credit_units[sem]
    course_titles = semester_course_titles[sem]

    ordered_titles = list(course_map.keys())
    ordered_codes = [course_map[t] for t in ordered_titles if course_map.get(t)]
    ordered_codes = [c for c in ordered_codes if credit_units.get(c, 0) > 0]
    filtered_credit_units = {c: credit_units[c] for c in ordered_codes}

    reg_no_cols = {s: find_column_by_names(df, ["REG. No", "Reg No", "Registration Number", "Mat No", "Exam No", "Student ID"]) for s, df in dfs.items()}
    name_cols = {s: find_column_by_names(df, ["NAME", "Full Name", "Candidate Name"]) for s, df in dfs.items()}

    merged = None
    for s, df in dfs.items():
        df = df.copy()
        regcol = reg_no_cols.get(s)
        namecol = name_cols.get(s)
        if not regcol:
            regcol = df.columns[0] if len(df.columns) > 0 else None
        if not namecol and len(df.columns) > 1:
            namecol = df.columns[1]

        if regcol is None:
            print(f"Skipping sheet {s}: no reg column found")
            continue

        df["REG. No"] = df[regcol].astype(str).str.strip()
        if namecol:
            df["NAME"] = df[namecol].astype(str).str.strip()
        else:
            df["NAME"] = pd.NA

        to_drop = [c for c in [regcol, namecol] if c and c not in ["REG. No", "NAME"]]
        df.drop(columns=to_drop, errors="ignore", inplace=True)

        for col in [c for c in df.columns if c not in ["REG. No", "NAME"]]:
            norm = normalize_course_name(col)
            matched_code = None
            for title, code in zip(ordered_titles, [course_map[t] for t in ordered_titles]):
                if normalize_course_name(title) == norm:
                    matched_code = code
                    break
            if matched_code:
                newcol = f"{matched_code}_{s.upper()}"
                df.rename(columns={col: newcol}, inplace=True)

        cur_cols = ["REG. No", "NAME"] + [c for c in df.columns if c.endswith(f"_{s.upper()}")]
        cur = df[cur_cols].copy()
        if merged is None:
            merged = cur
        else:
            merged = merged.merge(cur, on="REG. No", how="outer", suffixes=('', '_dup'))
            if "NAME_dup" in merged.columns:
                merged["NAME"] = merged["NAME"].combine_first(merged["NAME_dup"])
                merged.drop(columns=["NAME_dup"], inplace=True)

    if merged is None or merged.empty:
        print("No data merged from sheets — skipping file.")
        return None

    mastersheet = merged[["REG. No", "NAME"]].copy()
    mastersheet.rename(columns={"REG. No": "EXAMS NUMBER"}, inplace=True)

    for code in ordered_codes:
        ca_col = f"{code}_CA"
        obj_col = f"{code}_OBJ"
        exam_col = f"{code}_EXAM"

        ca_series = pd.to_numeric(merged[ca_col], errors="coerce") if ca_col in merged.columns else pd.Series([0]*len(merged), index=merged.index)
        obj_series = pd.to_numeric(merged[obj_col], errors="coerce") if obj_col in merged.columns else pd.Series([0]*len(merged), index=merged.index)
        exam_series = pd.to_numeric(merged[exam_col], errors="coerce") if exam_col in merged.columns else pd.Series([0]*len(merged), index=merged.index)

        ca_norm = (ca_series / 20) * 100
        obj_norm = (obj_series / 20) * 100
        exam_norm = (exam_series / 80) * 100
        ca_norm = ca_norm.fillna(0).clip(upper=100)
        obj_norm = obj_norm.fillna(0).clip(upper=100)
        exam_norm = exam_norm.fillna(0).clip(upper=100)
        total = (ca_norm * 0.2) + (((obj_norm + exam_norm) / 2) * 0.8)
        mastersheet[code] = total.round(0).clip(upper=100).values

    for c in ordered_codes:
        if c not in mastersheet.columns:
            mastersheet[c] = 0

    def compute_remarks(row):
        """Compute remarks with expanded failed courses list."""
        fails = [c for c in ordered_codes if float(row.get(c, 0) or 0) < pass_threshold]
        if not fails:
            return "Passed"
        # Expanded remarks to accommodate maximum failed courses
        failed_courses_str = ", ".join(sorted(fails))
        return f"Failed: {failed_courses_str}"

    # Calculate TCPE, TCUP, TCUF correctly
    def calc_tcpe_tcup_tcuf(row):
        tcpe = 0.0
        tcup = 0
        tcuf = 0
        
        for code in ordered_codes:
            score = float(row.get(code, 0) or 0)
            cu = filtered_credit_units.get(code, 0)
            gp = get_grade_point(score)
            
            # TCPE: Grade Point × Credit Units
            tcpe += gp * cu
            
            # TCUP/TCUF: Count credit units based on pass/fail
            if score >= pass_threshold:
                tcup += cu
            else:
                tcuf += cu
                
        return tcpe, tcup, tcuf

    # Apply calculations to each row
    results = mastersheet.apply(calc_tcpe_tcup_tcuf, axis=1, result_type='expand')
    mastersheet["TCPE"] = results[0].round(1)
    mastersheet["CU Passed"] = results[1]
    mastersheet["CU Failed"] = results[2]

    mastersheet["REMARKS"] = mastersheet.apply(compute_remarks, axis=1)

    total_cu = sum(filtered_credit_units.values()) if filtered_credit_units else 0
    mastersheet["GPA"] = mastersheet["TCPE"].apply(lambda x: round((x / total_cu), 2) if total_cu > 0 else 0.0)
    mastersheet["AVERAGE"] = mastersheet[[c for c in ordered_codes]].mean(axis=1).round(0)

    def sort_key(remark):
        if remark == "Passed":
            return (0, "")
        else:
            failed_courses = remark.replace("Failed: ", "").split(", ")
            return (1, len(failed_courses), ",".join(sorted(failed_courses)))
    mastersheet = mastersheet.sort_values(by="REMARKS", key=lambda x: x.map(sort_key)).reset_index(drop=True)

    if "S/N" not in mastersheet.columns:
        mastersheet.insert(0, "S/N", range(1, len(mastersheet) + 1))
    else:
        mastersheet["S/N"] = range(1, len(mastersheet) + 1)
        cols = list(mastersheet.columns)
        if cols[0] != "S/N":
            cols.remove("S/N")
            mastersheet = mastersheet[["S/N"] + cols]

    course_cols = ordered_codes
    out_cols = ["S/N", "EXAMS NUMBER", "NAME"] + course_cols + ["REMARKS", "CU Passed", "CU Failed", "TCPE", "GPA", "AVERAGE"]
    for c in out_cols:
        if c not in mastersheet.columns:
            mastersheet[c] = pd.NA
    mastersheet = mastersheet[out_cols]

    output_subdir = os.path.join(output_dir, f"ND_RESULT-{ts}")
    os.makedirs(output_subdir, exist_ok=True)
    out_xlsx = os.path.join(output_subdir, f"mastersheet_{ts}.xlsx")

    if not os.path.exists(out_xlsx):
        wb = Workbook()
        if wb.active:
            wb.remove(wb.active)
    else:
        wb = load_workbook(out_xlsx)

    if sem not in wb.sheetnames:
        ws = wb.create_sheet(title=sem)
    else:
        ws = wb[sem]

    try:
        ws.delete_rows(1, ws.max_row)
        ws.delete_cols(1, ws.max_column)
    except Exception:
        pass

    ws.insert_rows(1, 2)
    logo_path_norm = os.path.normpath(logo_path) if logo_path else None
    if logo_path_norm and os.path.exists(logo_path_norm):
        try:
            img = XLImage(logo_path_norm)
            img.width, img.height = 110, 110
            ws.add_image(img, "A1")
        except Exception as e:
            print(f"⚠ Could not place logo: {e}")

    ws.merge_cells("C1:Q1")
    title_cell = ws["C1"]
    title_cell.value = "FCT COLLEGE OF NURSING SCIENCES, GWAGWALADA-ABUJA"
    title_cell.font = Font(bold=True, size=16, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    border = Border(left=Side(style="medium"), right=Side(style="medium"), top=Side(style="medium"), bottom=Side(style="medium"))
    title_cell.border = border

    # Use expanded semester name in the subtitle
    expanded_semester_name = f"{level_display} {semester_display}"
    
    ws.merge_cells("C2:Q2")
    subtitle_cell = ws["C2"]
    subtitle_cell.value = f"{datetime.now().year}/{datetime.now().year + 1} SESSION  NATIONAL DIPLOMA {expanded_semester_name} EXAMINATIONS RESULT — {datetime.now().strftime('%B %d, %Y')}"
    subtitle_cell.font = Font(bold=True, size=12, color="000000")
    subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")

    display_course_titles = []
    for t, c in zip(ordered_titles, [course_map[t] for t in ordered_titles]):
        if c in ordered_codes:
            display_course_titles.append(t)

    ws.append([""] * 3 + display_course_titles + [""] * 5)
    for i, cell in enumerate(ws[3][3:3+len(display_course_titles)], start=3):
        cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=45)
        cell.font = Font(bold=True, size=9)
    ws.row_dimensions[3].height = 18

    cu_list = [filtered_credit_units.get(c, "") for c in ordered_codes]
    ws.append([""] * 3 + cu_list + [""] * 5)
    for cell in ws[4][3:3+len(cu_list)]:
        cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=135)
        cell.font = Font(bold=True, size=9)
        cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    headers = out_cols
    ws.append(headers)
    for cell in ws[5]:
        cell.font = Font(bold=True, size=10, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    for _, r in mastersheet.iterrows():
        rowvals = [r[col] for col in headers]
        ws.append(rowvals)

    # Freeze column F (column 6 - first course column)
    ws.freeze_panes = "F6"

    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for row in ws.iter_rows(min_row=6, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

    # Colorize course columns
    for idx, code in enumerate(ordered_codes, start=4):
        col_letter = get_column_letter(idx)
        for r_idx in range(6, ws.max_row + 1):
            cell = ws.cell(row=r_idx, column=idx)
            try:
                val = float(cell.value) if cell.value not in (None, "") else 0
                if val >= pass_threshold:
                    cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    cell.font = Font(color="006100")
                else:
                    cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                    cell.font = Font(color="FF0000", bold=True)
            except Exception:
                continue

    # Adjust column widths - EXPANDED REMARKS AND SEMESTER COLUMNS
    longest_name_len = max([len(str(x)) for x in mastersheet["NAME"].fillna("")]) if "NAME" in mastersheet.columns else 10
    name_col_width = min(max(longest_name_len + 2, 10), NAME_WIDTH_CAP)
    
    # Find longest remarks for proper column width
    longest_remark_len = max([len(str(x)) for x in mastersheet["REMARKS"].fillna("")]) if "REMARKS" in mastersheet.columns else 20
    remarks_col_width = min(max(longest_remark_len + 4, 35), 60)  # Expanded remarks column

    for col_idx, col in enumerate(ws.columns, start=1):
        column_letter = get_column_letter(col_idx)
        if col_idx == 1:  # S/N
            ws.column_dimensions[column_letter].width = 6
        elif column_letter == "B" or headers[col_idx-1] in ["EXAMS NUMBER", "EXAM NO"]:
            ws.column_dimensions[column_letter].width = 18
        elif headers[col_idx-1] == "NAME":
            ws.column_dimensions[column_letter].width = name_col_width
        elif 4 <= col_idx < 4 + len(ordered_codes):  # course columns
            ws.column_dimensions[column_letter].width = 8
        elif headers[col_idx-1] in ["REMARKS"]:
            ws.column_dimensions[column_letter].width = remarks_col_width  # Expanded remarks
        else:
            ws.column_dimensions[column_letter].width = 12

    # Fails per course row
    fails_per_course = mastersheet[ordered_codes].apply(lambda x: (x < pass_threshold).sum()).tolist()
    footer_vals = [""] * 2 + ["FAILS PER COURSE:"] + fails_per_course + [""] * (len(headers) - 3 - len(ordered_codes))
    ws.append(footer_vals)
    for cell in ws[ws.max_row]:
        if 4 <= cell.column < 4 + len(ordered_codes):
            cell.fill = PatternFill(start_color="F0E68C", end_color="F0E68C", fill_type="solid")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        elif cell.column == 3:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

    # Summary block
    total_students = len(mastersheet)
    passed_all = len(mastersheet[mastersheet["REMARKS"] == "Passed"])
    failed_over45 = len(mastersheet[mastersheet["CU Failed"] > 0.45 * total_cu]) if total_cu else 0

    ws.append([])
    ws.append(["SUMMARY"])
    ws.append([f"A total of {total_students} students registered and sat for the Examination"])
    ws.append([f"A total of {passed_all} students passed in all courses registered."])
    ws.append([f"A total of {failed_over45} students failed in more than 45% of their registered credit units."])
    ws.append(["The above decisions are in line with the provisions of the General Information Section of the NMCN/NBTE Examinations Regulations."])
    ws.append([])
    ws.append(["________________________", "", "", "________________________", "", "", "", "", "", "", "", "", ""])
    ws.append(["Mrs. Abini Hauwa", "", "", "Mrs. Olukemi Ogunleye", "", "", "", "", "", "", "", "", ""])
    ws.append(["Head of Exams", "", "", "Chairman, ND/HND Program C'tee", "", "", "", "", "", "", "", "", ""])

    wb.save(out_xlsx)
    print(f"✅ Mastersheet saved: {out_xlsx}")

    # Generate individual student PDF with previous GPAs
    safe_sem = re.sub(r'[^\w\-]', '_', sem)
    student_pdf_path = os.path.join(output_subdir, f"mastersheet_students_{ts}_{safe_sem}.pdf")
    
    print(f"📊 FINAL CHECK before PDF generation:")
    print(f"   Previous GPAs loaded: {len(previous_gpas)}")
    if previous_gpas:
        sample = list(previous_gpas.items())[:3]
        print(f"   Sample GPAs: {sample}")
    
    try:
        generate_individual_student_pdf(mastersheet, student_pdf_path, sem, logo_path=logo_path_norm, 
                                       prev_mastersheet_df=None, filtered_credit_units=filtered_credit_units,
                                       ordered_codes=ordered_codes, course_titles_map=course_titles,
                                       previous_gpas=previous_gpas)
    except Exception as e:
        print(f"⚠ Failed to generate student PDF: {e}")
        import traceback
        traceback.print_exc()

    return mastersheet

# ----------------------------
# Main runner
# ----------------------------
def main():
    print("Starting ND Examination Results Processing...")
    ts = datetime.now().strftime(TIMESTAMP_FMT)

    base_dir_norm = normalize_path(BASE_DIR)
    print(f"Using base directory: {base_dir_norm}")

    try:
        semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_course_data()
    except Exception as e:
        print(f"❌ Could not load course data: {e}")
        return

    nd_dirs = [d for d in os.listdir(base_dir_norm) if os.path.isdir(os.path.join(base_dir_norm, d)) and d.upper().startswith("ND-")]
    if not nd_dirs:
        print(f"No ND-* directories found in {base_dir_norm}. Nothing to process.")
        return

    for nd in nd_dirs:
        print(f"\n--- Processing ND folder: {nd} ---")
        raw_dir = normalize_path(os.path.join(base_dir_norm, nd, "RAW_RESULTS"))
        clean_dir = normalize_path(os.path.join(base_dir_norm, nd, "CLEAN_RESULTS"))
        os.makedirs(raw_dir, exist_ok=True)
        os.makedirs(clean_dir, exist_ok=True)

        raw_files = [f for f in os.listdir(raw_dir) if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
        if not raw_files:
            print(f"⚠️ No raw files in {raw_dir}; skipping {nd}")
            continue

        print(f"📁 Found {len(raw_files)} raw files")
        
        # Process semesters in the correct order
        for semester_key in SEMESTER_ORDER:
            process_semester_files(semester_key, raw_files, raw_dir, clean_dir, ts, 
                                 DEFAULT_PASS_THRESHOLD, semester_course_maps, semester_credit_units, 
                                 semester_lookup, semester_course_titles, DEFAULT_LOGO_PATH)

    print("\n✅ ND Examination Results Processing completed successfully.")

if __name__ == "__main__":
    main()