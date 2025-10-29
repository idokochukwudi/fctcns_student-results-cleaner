#!/usr/bin/env python3
"""
exam_result_processor.py
Complete script with integrated carryover student management.
FIXED VERSION - Regular carryover file generation without enhanced formatting
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
import tempfile
import shutil
import json
import subprocess

# ----------------------------
# Configuration
# ----------------------------
def is_running_on_railway():
    """Check if we're running on Railway"""
    return any(key in os.environ for key in [
        'RAILWAY_ENVIRONMENT', 
        'RAILWAY_STATIC_URL', 
        'RAILWAY_PROJECT_ID',
        'RAILWAY_SERVICE_NAME'
    ])

def get_base_directory():
    """Get base directory - compatible with both local and Railway environments"""
    # Check if BASE_DIR is explicitly set in environment (highest priority)
    base_dir_env = os.getenv('BASE_DIR')
    if base_dir_env:
        if os.path.exists(base_dir_env):
            print(f"✅ Using BASE_DIR from environment: {base_dir_env}")
            return base_dir_env
        else:
            print(f"⚠️ BASE_DIR from environment doesn't exist: {base_dir_env}, trying alternatives...")
    
    # Check if we're running on Railway
    if is_running_on_railway():
        # Create the directory structure on Railway
        railway_base = '/app/EXAMS_INTERNAL'
        os.makedirs(railway_base, exist_ok=True)
        os.makedirs(os.path.join(railway_base, 'ND', 'ND-COURSES'), exist_ok=True)
        print(f"✅ Using Railway base directory: {railway_base}")
        return railway_base
    
    # Local development fallbacks - check multiple possible locations
    local_paths = [
        # Your specific structure
        os.path.join(os.path.expanduser('~'), 'student_result_cleaner', 'EXAMS_INTERNAL'),
        # Common development locations
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'EXAMS_INTERNAL'),
        os.path.join(os.getcwd(), 'EXAMS_INTERNAL'),
        # Relative to script location
        os.path.join(os.path.dirname(__file__), 'EXAMS_INTERNAL'),
    ]
    
    for local_path in local_paths:
        if os.path.exists(local_path):
            print(f"✅ Using local base directory: {local_path}")
            return local_path
    
    # Final fallback - create in current working directory
    fallback_path = os.path.join(os.getcwd(), 'EXAMS_INTERNAL')
    print(f"⚠️ No existing directory found, creating fallback: {fallback_path}")
    os.makedirs(fallback_path, exist_ok=True)
    os.makedirs(os.path.join(fallback_path, 'ND', 'ND-COURSES'), exist_ok=True)
    return fallback_path

BASE_DIR = get_base_directory()
# UPDATED: ND directories now under ND folder
ND_BASE_DIR = os.path.join(BASE_DIR, "ND")
ND_COURSES_DIR = os.path.join(ND_BASE_DIR, "ND-COURSES")

# Ensure directories exist
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(ND_BASE_DIR, exist_ok=True)
os.makedirs(ND_COURSES_DIR, exist_ok=True)

print(f"📁 Base directory: {BASE_DIR}")
print(f"📁 ND base directory: {ND_BASE_DIR}")
print(f"📁 ND courses directory: {ND_COURSES_DIR}")

# Global variables for threshold upgrade
THRESHOLD_UPGRADED = False
ORIGINAL_THRESHOLD = 50.0
UPGRADE_MIN = None
UPGRADE_MAX = 49

# Global carryover tracker
CARRYOVER_STUDENTS = {}

def is_web_mode():
    """Check if running in web mode (file upload)"""
    return os.getenv('WEB_MODE') == 'true'

def get_uploaded_file_path():
    """Get path of uploaded file in web mode"""
    return os.getenv('UPLOADED_FILE_PATH')

def should_use_interactive_mode():
    """Check if we should use interactive mode (CLI) or non-interactive mode (web)."""
    # If specific environment variables are set by web form, use non-interactive
    if os.getenv('SELECTED_SET') or os.getenv('PROCESSING_MODE') or is_web_mode():
        return False
    # If we're running in a terminal with stdin available, use interactive mode
    if sys.stdin.isatty():
        return True
    # Default to interactive for backward compatibility
    return True

def get_upgrade_threshold_from_env():
    """Get upgrade threshold from environment variables"""
    upgrade_threshold_str = os.getenv('UPGRADE_THRESHOLD', '0').strip()
    if upgrade_threshold_str and upgrade_threshold_str.isdigit():
        upgrade_value = int(upgrade_threshold_str)
        if 0 <= upgrade_value <= 49:
            return upgrade_value if upgrade_value > 0 else None
    return None

def get_form_parameters():
    """Get parameters from environment variables set by the web form."""
    selected_set = os.getenv('SELECTED_SET', 'all')
    processing_mode = os.getenv('PROCESSING_MODE', 'auto')
    selected_semesters_str = os.getenv('SELECTED_SEMESTERS', '')
    pass_threshold = float(os.getenv('PASS_THRESHOLD', '50.0'))
    generate_pdf = os.getenv('GENERATE_PDF', 'True').lower() == 'true'
    track_withdrawn = os.getenv('TRACK_WITHDRAWN', 'True').lower() == 'true'
    
    # NEW: Check for carryover processing mode
    process_carryover = os.getenv('PROCESS_CARRYOVER', 'False').lower() == 'true'
    carryover_file_path = os.getenv('CARRYOVER_FILE_PATH', '')
    
    # Convert semester string to list - handle both comma-separated and single values
    selected_semesters = []
    if selected_semesters_str:
        if ',' in selected_semesters_str:
            selected_semesters = [sem.strip() for sem in selected_semesters_str.split(',') if sem.strip()]
        else:
            selected_semesters = [selected_semesters_str.strip()]
    
    # If no semesters selected or 'all' in selected, use all semesters
    if not selected_semesters or 'all' in selected_semesters:
        selected_semesters = SEMESTER_ORDER.copy()
    
    # FIX: Convert selected semesters to UPPERCASE to match course data
    selected_semesters = [sem.upper() for sem in selected_semesters]
    
    print(f"🎯 FORM PARAMETERS:")
    print(f"   Selected Set: {selected_set}")
    print(f"   Processing Mode: {processing_mode}")
    print(f"   Selected Semesters: {selected_semesters}")
    print(f"   Pass Threshold: {pass_threshold}")
    print(f"   Generate PDF: {generate_pdf}")
    print(f"   Track Withdrawn: {track_withdrawn}")
    print(f"   Process Carryover: {process_carryover}")
    print(f"   Carryover File Path: {carryover_file_path}")
    
    return {
        'selected_set': selected_set,
        'processing_mode': processing_mode,
        'selected_semesters': selected_semesters,
        'pass_threshold': pass_threshold,
        'generate_pdf': generate_pdf,
        'track_withdrawn': track_withdrawn,
        'process_carryover': process_carryover,
        'carryover_file_path': carryover_file_path
    }

def get_pass_threshold():
    """Get pass threshold - now handles upgrade logic interactively."""
    threshold_str = os.getenv('PASS_THRESHOLD', '50.0')
    try:
        threshold = float(threshold_str)
    except ValueError:
        threshold = 50.0
    
    return threshold

DEFAULT_PASS_THRESHOLD = get_pass_threshold()

TIMESTAMP_FMT = "%Y-%m-%d_%H%M%S"

# FIXED: More robust logo path detection
def get_logo_path():
    """Get the logo path with multiple fallback options."""
    possible_paths = [
        # Relative to script
        os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "launcher", "static", "logo.png")),
        # Common locations
        os.path.join(os.path.dirname(__file__), "logo.png"),
        os.path.join(os.getcwd(), "logo.png"),
        # Absolute path fallback
        "/app/launcher/static/logo.png"  # For Railway deployment
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            print(f"✅ Found logo at: {path}")
            return path
    
    print("⚠️ Logo not found, PDF generation will proceed without logo")
    return None

DEFAULT_LOGO_PATH = get_logo_path()

NAME_WIDTH_CAP = 40

# Define semester processing order - FIXED: Use consistent uppercase
SEMESTER_ORDER = [
    "ND-FIRST-YEAR-FIRST-SEMESTER",
    "ND-FIRST-YEAR-SECOND-SEMESTER",
    "ND-SECOND-YEAR-FIRST-SEMESTER",
    "ND-SECOND-YEAR-SECOND-SEMESTER"
]

# Global student tracker
STUDENT_TRACKER = {}
WITHDRAWN_STUDENTS = {}

# ----------------------------
# DATA TRANSFORMATION FUNCTIONS - NEW ADDITION
# ----------------------------

def transform_transposed_data(df, sheet_type):
    """
    Transform transposed data format to wide format.
    Input: Each student appears multiple times with different courses
    Output: Each student appears once with all courses as columns
    """
    print(f"🔄 Transforming {sheet_type} sheet from transposed to wide format...")
    
    # Find the registration and name columns
    reg_col = find_column_by_names(df, ["REG. No", "Reg No", "Registration Number", "EXAM NUMBER"])
    name_col = find_column_by_names(df, ["NAME", "Full Name", "Candidate Name"])
    
    if not reg_col:
        print("❌ Could not find registration column for transformation")
        return df
    
    # Get all course columns (columns that contain course codes)
    course_columns = [col for col in df.columns 
                     if col not in [reg_col, name_col] and col not in ['', None]]
    
    print(f"📊 Found {len(course_columns)} course columns: {course_columns}")
    
    # Create a new dataframe to store transformed data
    transformed_data = []
    student_dict = {}
    
    # Process each row
    for idx, row in df.iterrows():
        exam_no = str(row[reg_col]).strip()
        student_name = str(row[name_col]).strip() if name_col and pd.notna(row.get(name_col)) else ""
        
        if exam_no not in student_dict:
            student_dict[exam_no] = {
                'REG. No': exam_no,
                'NAME': student_name
            }
        
        # Add course scores for this student
        for course_col in course_columns:
            score = row.get(course_col)
            if pd.notna(score) and score != "" and score != " ":
                # Create column name with sheet type suffix
                column_name = f"{course_col}_{sheet_type}"
                student_dict[exam_no][column_name] = score
    
    # Convert dictionary to list
    transformed_data = list(student_dict.values())
    
    # Create new DataFrame
    if transformed_data:
        transformed_df = pd.DataFrame(transformed_data)
        print(f"✅ Transformed data: {len(transformed_df)} students, {len(transformed_df.columns)} columns")
        return transformed_df
    else:
        print("❌ No data after transformation")
        return df

def detect_data_format(df, sheet_type):
    """
    Detect if data is in transposed format (students appear multiple times)
    Returns True if transposed format is detected
    """
    reg_col = find_column_by_names(df, ["REG. No", "Reg No", "Registration Number", "EXAM NUMBER"])
    
    if not reg_col:
        return False
    
    # Count occurrences of each student
    student_counts = df[reg_col].value_counts()
    max_occurrences = student_counts.max()
    
    # If any student appears more than once, it's likely transposed format
    if max_occurrences > 1:
        print(f"📊 Data format detection for {sheet_type}:")
        print(f"   Total students: {len(student_counts)}")
        print(f"   Max occurrences per student: {max_occurrences}")
        print(f"   Students with multiple entries: {(student_counts > 1).sum()}")
        return True
    
    return False

# ----------------------------
# Enhanced Course Name Matching Functions
# ----------------------------

def normalize_course_name(name):
    """Enhanced normalization for course title matching with better variations handling."""
    if not isinstance(name, str):
        return ""
    
    # Convert to lowercase and remove extra spaces
    normalized = name.lower().strip()
    
    # Replace multiple spaces with single space
    normalized = re.sub(r'\s+', ' ', normalized)
    
    # Remove special characters and extra words
    normalized = re.sub(r'[^\w\s]', '', normalized)
    
    # Enhanced substitutions for variations
    substitutions = {
        'coomunication': 'communication',
        'nsg': 'nursing',
        'foundation': 'foundations',
        'of of': 'of',  # handle double "of"
        'emergency care': 'emergency',
        'nursing/ emergency': 'nursing emergency',
        'care i': 'care',
        'foundations of nursing': 'foundations nursing',
        'foundation of nsg': 'foundations nursing',
        'foundation of nursing': 'foundations nursing',
    }
    
    for old, new in substitutions.items():
        normalized = normalized.replace(old, new)
        
    return normalized.strip()

def find_best_course_match(column_name, course_map):
    """Find the best matching course using enhanced matching algorithm."""
    if not isinstance(column_name, str):
        return None
        
    normalized_column = normalize_course_name(column_name)
    
    # First try exact match
    if normalized_column in course_map:
        return course_map[normalized_column]
    
    # Try partial matches with higher priority
    for course_norm, course_info in course_map.items():
        # Check if one is contained in the other
        if course_norm in normalized_column or normalized_column in course_norm:
            return course_info
    
    # Try word-based matching
    column_words = set(normalized_column.split())
    best_match = None
    best_score = 0
    
    for course_norm, course_info in course_map.items():
        course_words = set(course_norm.split())
        common_words = column_words.intersection(course_words)
        
        if common_words:
            score = len(common_words)
            # Bonus for matching key words
            key_words = ['foundation', 'nursing', 'emergency', 'care', 'communication', 'anatomy', 'physiology']
            for word in key_words:
                if word in column_words and word in course_words:
                    score += 2
            
            if score > best_score:
                best_score = score
                best_match = course_info
    
    # Require at least 2 common words or 1 key word with good match
    if best_match and best_score >= 2:
        return best_match
    
    # Final fallback: use difflib for fuzzy matching
    best_match = None
    best_ratio = 0
    
    for course_norm, course_info in course_map.items():
        ratio = difflib.SequenceMatcher(None, normalized_column, course_norm).ratio()
        if ratio > best_ratio and ratio > 0.6:  # Lower threshold for fuzzy matching
            best_ratio = ratio
            best_match = course_info
    
    return best_match

# ----------------------------
# Carryover Management Functions - UPDATED: No enhanced formatting
# ----------------------------

def initialize_carryover_tracker():
    """Initialize the global carryover tracker."""
    global CARRYOVER_STUDENTS
    CARRYOVER_STUDENTS = {}

def identify_carryover_students(mastersheet_df, semester_key, set_name, pass_threshold=50.0):
    """
    Identify students with carryover courses from current semester processing.
    Returns list of carryover students with their failed courses.
    """
    carryover_students = []
    
    # Get course columns (excluding student info columns)
    course_columns = [col for col in mastersheet_df.columns 
                     if col not in ['S/N', 'EXAM NUMBER', 'NAME', 'REMARKS', 
                                   'CU Passed', 'CU Failed', 'TCPE', 'GPA', 'AVERAGE']]
    
    for idx, student in mastersheet_df.iterrows():
        failed_courses = []
        exam_no = str(student['EXAM NUMBER'])
        student_name = student['NAME']
        
        for course in course_columns:
            score = student.get(course, 0)
            try:
                score_val = float(score) if pd.notna(score) else 0
                if score_val < pass_threshold:
                    failed_courses.append({
                        'course_code': course,
                        'original_score': score_val,
                        'semester': semester_key,
                        'set': set_name,
                        'resit_attempts': 0,
                        'best_score': score_val,
                        'status': 'Failed'  # Failed, Passed_Resit, Carryover
                    })
            except (ValueError, TypeError):
                continue
        
        if failed_courses:
            carryover_data = {
                'exam_number': exam_no,
                'name': student_name,
                'failed_courses': failed_courses,
                'semester': semester_key,
                'set': set_name,
                'identified_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'total_resit_attempts': 0,
                'status': 'Active'  # Active, Cleared, Withdrawn
            }
            carryover_students.append(carryover_data)
            
            # Update global tracker
            student_key = f"{exam_no}_{semester_key}"
            CARRYOVER_STUDENTS[student_key] = carryover_data
    
    return carryover_students

def save_carryover_records(carryover_students, output_dir, set_name, semester_key):
    """
    Save carryover student records to the clean results folder.
    UPDATED: SIMPLE Excel structure WITHOUT enhanced formatting
    """
    if not carryover_students:
        print("ℹ️ No carryover students to save")
        return None
    
    # Create carryover subdirectory in clean results
    carryover_dir = os.path.join(output_dir, "CARRYOVER_RECORDS")
    os.makedirs(carryover_dir, exist_ok=True)
    
    # Generate filename with set and semester tags
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"co_student_{set_name}_{semester_key}_{timestamp}"
    
    # Save as Excel
    excel_file = os.path.join(carryover_dir, f"{filename}.xlsx")
    
    # Prepare data for Excel - simple structure without enhanced formatting
    records_data = []
    for student in carryover_students:
        for course in student['failed_courses']:
            records_data.append({
                'EXAM NUMBER': student['exam_number'],
                'NAME': student['name'],
                'COURSE CODE': course['course_code'],
                'ORIGINAL SCORE': course['original_score'],
                'SEMESTER': student['semester'],
                'SET': student['set'],
                'RESIT ATTEMPTS': course['resit_attempts'],
                'BEST SCORE': course['best_score'],
                'STATUS': course['status'],
                'IDENTIFIED DATE': student['identified_date']
            })
    
    if records_data:
        df = pd.DataFrame(records_data)
        df.to_excel(excel_file, index=False)
        print(f"✅ Carryover records saved: {excel_file}")
        
        # UPDATED: NO enhanced formatting - keep it simple for regular processing
        try:
            # Just add basic header formatting without logo or title
            wb = load_workbook(excel_file)
            ws = wb.active
            
            # Simple header formatting only
            header_row = ws[1]
            for cell in header_row:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
            
            # Auto-adjust column widths for readability
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                ws.column_dimensions[column_letter].width = adjusted_width
            
            wb.save(excel_file)
            print("✅ Added basic formatting to carryover Excel file")
            
        except Exception as e:
            print(f"⚠️ Could not add basic formatting to carryover file: {e}")
    
    # Save as JSON for easy processing
    json_file = os.path.join(carryover_dir, f"{filename}.json")
    with open(json_file, 'w') as f:
        json.dump(carryover_students, f, indent=2)
    
    print(f"📁 Regular carryover records saved in: {carryover_dir}")
    return carryover_dir

def check_existing_carryover_files(raw_dir, set_name, semester_key):
    """
    Check if carryover files already exist for a given set and semester.
    Returns list of existing carryover file paths.
    """
    carryover_dir = os.path.join(raw_dir, "CARRYOVER")
    if not os.path.exists(carryover_dir):
        return []
    
    # Look for carryover files matching the pattern
    pattern = f"CARRYOVER-{semester_key}-{set_name}*.xlsx"
    existing_files = []
    
    for file in os.listdir(carryover_dir):
        if file.startswith(f"CARRYOVER-{semester_key}-{set_name}") and file.endswith('.xlsx'):
            existing_files.append(os.path.join(carryover_dir, file))
    
    print(f"🔍 Found {len(existing_files)} existing carryover files for {set_name}/{semester_key}")
    return existing_files

# ----------------------------
# CGPA Tracking Functions
# ----------------------------

def create_cgpa_summary_sheet(mastersheet_path, timestamp):
    """
    Create a CGPA summary sheet that aggregates GPA across all semesters.
    """
    try:
        print("📊 Creating CGPA Summary Sheet...")
        
        # Load the mastersheet workbook
        wb = load_workbook(mastersheet_path)
        
        # Collect GPA data from all semesters
        cgpa_data = {}
        
        for sheet_name in wb.sheetnames:
            if sheet_name in SEMESTER_ORDER:
                df = pd.read_excel(mastersheet_path, sheet_name=sheet_name, header=5)
                
                # Find exam number and GPA columns
                exam_col = find_exam_number_column(df)
                gpa_col = None
                name_col = None
                
                for col in df.columns:
                    col_str = str(col).upper()
                    if 'GPA' in col_str:
                        gpa_col = col
                    elif 'NAME' in col_str:
                        name_col = col
                
                if exam_col and gpa_col:
                    for idx, row in df.iterrows():
                        exam_no = str(row[exam_col]).strip()
                        if exam_no and exam_no != 'nan':
                            if exam_no not in cgpa_data:
                                cgpa_data[exam_no] = {
                                    'name': row[name_col] if name_col and pd.notna(row.get(name_col)) else '',
                                    'gpas': {}
                                }
                            cgpa_data[exam_no]['gpas'][sheet_name] = row[gpa_col]
        
        # Create CGPA summary dataframe
        summary_data = []
        for exam_no, data in cgpa_data.items():
            row = {
                'EXAM NUMBER': exam_no,
                'NAME': data['name']
            }
            
            # Add GPA for each semester
            total_gpa = 0
            semester_count = 0
            
            for semester in SEMESTER_ORDER:
                if semester in data['gpas']:
                    row[semester] = data['gpas'][semester]
                    if pd.notna(data['gpas'][semester]):
                        total_gpa += data['gpas'][semester]
                        semester_count += 1
                else:
                    row[semester] = None
            
            # Calculate CGPA
            row['CGPA'] = round(total_gpa / semester_count, 2) if semester_count > 0 else 0.0
            summary_data.append(row)
        
        # Create summary dataframe
        summary_df = pd.DataFrame(summary_data)
        
        # Add the summary sheet to the workbook
        if 'CGPA_SUMMARY' in wb.sheetnames:
            del wb['CGPA_SUMMARY']
        
        ws = wb.create_sheet('CGPA_SUMMARY')
        
        # Write header
        headers = ['EXAM NUMBER', 'NAME'] + SEMESTER_ORDER + ['CGPA']
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)
        
        # Write data
        for row_idx, row_data in enumerate(summary_data, 2):
            for col_idx, header in enumerate(headers, 1):
                ws.cell(row=row_idx, column=col_idx, value=row_data.get(header, ''))
        
        # Style the header
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        
        wb.save(mastersheet_path)
        print("✅ CGPA Summary sheet created successfully")
        
        return summary_df
        
    except Exception as e:
        print(f"❌ Error creating CGPA summary sheet: {e}")
        return None

def create_analysis_sheet(mastersheet_path, timestamp):
    """
    Create an analysis sheet with comprehensive statistics.
    """
    try:
        print("📈 Creating Analysis Sheet...")
        
        wb = load_workbook(mastersheet_path)
        
        # Collect data from all semesters
        analysis_data = {
            'semester': [],
            'total_students': [],
            'passed_all': [],
            'carryover_students': [],
            'withdrawn_students': [],
            'average_gpa': [],
            'pass_rate': []
        }
        
        for sheet_name in wb.sheetnames:
            if sheet_name in SEMESTER_ORDER:
                df = pd.read_excel(mastersheet_path, sheet_name=sheet_name, header=5)
                
                # Basic statistics
                total_students = len(df)
                passed_all = len(df[df['REMARKS'] == 'Passed']) if 'REMARKS' in df.columns else 0
                
                # Calculate carryover students
                carryover_count = 0
                if 'REMARKS' in df.columns:
                    carryover_count = len(df[df['REMARKS'].str.contains('Failed', na=False)])
                
                # Calculate average GPA
                avg_gpa = df['GPA'].mean() if 'GPA' in df.columns and not df['GPA'].isna().all() else 0
                
                # Calculate pass rate
                pass_rate = (passed_all / total_students * 100) if total_students > 0 else 0
                
                analysis_data['semester'].append(sheet_name)
                analysis_data['total_students'].append(total_students)
                analysis_data['passed_all'].append(passed_all)
                analysis_data['carryover_students'].append(carryover_count)
                analysis_data['withdrawn_students'].append(0)  # You can track this if available
                analysis_data['average_gpa'].append(round(avg_gpa, 2))
                analysis_data['pass_rate'].append(round(pass_rate, 2))
        
        # Create analysis dataframe
        analysis_df = pd.DataFrame(analysis_data)
        
        # Add overall statistics
        overall_stats = {
            'semester': 'OVERALL',
            'total_students': analysis_df['total_students'].sum(),
            'passed_all': analysis_df['passed_all'].sum(),
            'carryover_students': analysis_df['carryover_students'].sum(),
            'withdrawn_students': analysis_df['withdrawn_students'].sum(),
            'average_gpa': round(analysis_df['average_gpa'].mean(), 2),
            'pass_rate': round(analysis_df['pass_rate'].mean(), 2)
        }
        analysis_df = pd.concat([analysis_df, pd.DataFrame([overall_stats])], ignore_index=True)
        
        # Add the analysis sheet to the workbook
        if 'ANALYSIS' in wb.sheetnames:
            del wb['ANALYSIS']
        
        ws = wb.create_sheet('ANALYSIS')
        
        # Write header
        headers = ['SEMESTER', 'TOTAL STUDENTS', 'PASSED ALL', 'CARRYOVER STUDENTS', 
                  'WITHDRAWN STUDENTS', 'AVERAGE GPA', 'PASS RATE (%)']
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)
        
        # Write data
        for row_idx, row_data in analysis_df.iterrows():
            ws.cell(row=row_idx+2, column=1, value=row_data['semester'])
            ws.cell(row=row_idx+2, column=2, value=row_data['total_students'])
            ws.cell(row=row_idx+2, column=3, value=row_data['passed_all'])
            ws.cell(row=row_idx+2, column=4, value=row_data['carryover_students'])
            ws.cell(row=row_idx+2, column=5, value=row_data['withdrawn_students'])
            ws.cell(row=row_idx+2, column=6, value=row_data['average_gpa'])
            ws.cell(row=row_idx+2, column=7, value=row_data['pass_rate'])
        
        # Style the header
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="27ae60", end_color="27ae60", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(mastersheet_path)
        print("✅ Analysis sheet created successfully")
        
        return analysis_df
        
    except Exception as e:
        print(f"❌ Error creating analysis sheet: {e}")
        return None

# ----------------------------
# Upgrade Rule Functions
# ----------------------------

def get_upgrade_threshold_from_user(semester_key, set_name):
    """
    Prompt user to choose upgrade threshold for ND results.
    Returns: (min_threshold, upgraded_count) or (None, 0) if skipped
    """
    print(f"\n🎯 MANAGEMENT THRESHOLD UPGRADE RULE DETECTED")
    print(f"📚 Semester: {semester_key}")
    print(f"📁 Set: {set_name}")
    print("\nSelect minimum score to upgrade (45-49). All scores >= selected value up to 49 will be upgraded to 50.")
    print("Enter 0 to skip upgrade.")
    
    while True:
        try:
            choice = input("\nEnter your choice (0, 45, 46, 47, 48, 49): ").strip()
            
            if not choice:
                print("❌ Please enter a value.")
                continue
                
            if choice == '0':
                print("⏭️ Skipping upgrade for this semester.")
                return None, 0
                
            if choice in ['45', '46', '47', '48', '49']:
                min_threshold = int(choice)
                print(f"✅ Upgrade rule selected: {min_threshold}–49 → 50")
                return min_threshold, 0
            else:
                print("❌ Invalid choice. Please enter 0, 45, 46, 47, 48, or 49.")
                
        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)
        except Exception as e:
            print(f"❌ Error: {e}. Please try again.")

def apply_upgrade_rule(mastersheet, ordered_codes, min_threshold):
    """
    Apply upgrade rule to mastersheet scores.
    Returns: (updated_mastersheet, upgraded_count)
    """
    if min_threshold is None:
        return mastersheet, 0
        
    upgraded_count = 0
    upgraded_students = set()
    
    print(f"🔄 Applying upgrade rule: {min_threshold}–49 → 50")
    
    for code in ordered_codes:
        for idx in mastersheet.index:
            score = mastersheet.at[idx, code]
            if isinstance(score, (int, float)) and min_threshold <= score <= 49:
                exam_no = mastersheet.at[idx, "EXAM NUMBER"]
                original_score = score
                mastersheet.at[idx, code] = 50
                upgraded_count += 1
                upgraded_students.add(exam_no)
                
                # Log first few upgrades for visibility
                if upgraded_count <= 5:
                    print(f"🔼 {exam_no} - {code}: {original_score} → 50")
    
    if upgraded_count > 0:
        print(f"✅ Upgraded {upgraded_count} scores from {min_threshold}–49 to 50")
        print(f"📊 Affected {len(upgraded_students)} students")
    else:
        print(f"ℹ️ No scores found in range {min_threshold}–49 to upgrade")
    
    return mastersheet, upgraded_count

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

def create_zip_folder(source_dir, zip_path):
    """
    Create a ZIP file from a directory.
    Returns True if successful, False otherwise.
    """
    try:
        import zipfile
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    # Create relative path for ZIP
                    arcname = os.path.relpath(file_path, source_dir)
                    zipf.write(file_path, arcname)
        print(f"✅ Successfully created ZIP: {zip_path}")
        return True
    except Exception as e:
        print(f"❌ Failed to create ZIP: {e}")
        return False

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
# Student Tracking Functions
# ----------------------------

def initialize_student_tracker():
    """Initialize the global student tracker."""
    global STUDENT_TRACKER, WITHDRAWN_STUDENTS
    STUDENT_TRACKER = {}
    WITHDRAWN_STUDENTS = {}

def update_student_tracker(
        semester_key,
        exam_numbers,
        withdrawn_students=None):
    """
    Update the student tracker with current semester's students.
    This helps track which students are present in each semester.
    """
    global STUDENT_TRACKER, WITHDRAWN_STUDENTS

    print(f"📊 Updating student tracker for {semester_key}")
    print(f"📝 Current students in this semester: {len(exam_numbers)}")

    # Track withdrawn students
    if withdrawn_students:
        for exam_no in withdrawn_students:
            if exam_no not in WITHDRAWN_STUDENTS:
                WITHDRAWN_STUDENTS[exam_no] = {
                    'withdrawn_semester': semester_key,
                    'withdrawn_date': datetime.now().strftime(TIMESTAMP_FMT),
                    'reappeared_semesters': []
                }
                print(f"🚫 Marked as withdrawn: {exam_no} in {semester_key}")

    for exam_no in exam_numbers:
        if exam_no not in STUDENT_TRACKER:
            STUDENT_TRACKER[exam_no] = {
                'first_seen': semester_key,
                'last_seen': semester_key,
                'semesters_present': [semester_key],
                'status': 'Active',
                'withdrawn': False,
                'withdrawn_semester': None
            }
        else:
            STUDENT_TRACKER[exam_no]['last_seen'] = semester_key
            if semester_key not in STUDENT_TRACKER[exam_no]['semesters_present']:
                STUDENT_TRACKER[exam_no]['semesters_present'].append(
                    semester_key)

            # Check if student was previously withdrawn and has reappeared
            if STUDENT_TRACKER[exam_no]['withdrawn']:
                print(f"⚠️ PREVIOUSLY WITHDRAWN STUDENT REAPPEARED: {exam_no}")
                if exam_no in WITHDRAWN_STUDENTS:
                    if semester_key not in WITHDRAWN_STUDENTS[exam_no]['reappeared_semesters']:
                        WITHDRAWN_STUDENTS[exam_no]['reappeared_semesters'].append(
                            semester_key)

    print(f"📈 Total unique students tracked: {len(STUDENT_TRACKER)}")
    print(f"🚫 Total withdrawn students: {len(WITHDRAWN_STUDENTS)}")

def mark_student_withdrawn(exam_no, semester_key):
    """Mark a student as withdrawn in a specific semester."""
    global STUDENT_TRACKER, WITHDRAWN_STUDENTS

    if exam_no in STUDENT_TRACKER:
        STUDENT_TRACKER[exam_no]['withdrawn'] = True
        STUDENT_TRACKER[exam_no]['withdrawn_semester'] = semester_key
        STUDENT_TRACKER[exam_no]['status'] = 'Withdrawn'

    if exam_no not in WITHDRAWN_STUDENTS:
        WITHDRAWN_STUDENTS[exam_no] = {
            'withdrawn_semester': semester_key,
            'withdrawn_date': datetime.now().strftime(TIMESTAMP_FMT),
            'reappeared_semesters': []
        }

def is_student_withdrawn(exam_no):
    """Check if a student has been withdrawn in any previous semester."""
    return exam_no in WITHDRAWN_STUDENTS

def get_withdrawal_history(exam_no):
    """Get withdrawal history for a student."""
    if exam_no in WITHDRAWN_STUDENTS:
        return WITHDRAWN_STUDENTS[exam_no]
    return None

def filter_out_withdrawn_students(mastersheet, semester_key):
    """
    Filter out students who were withdrawn in previous semesters.
    Returns filtered mastersheet and list of removed students.
    """
    removed_students = []
    filtered_mastersheet = mastersheet.copy()

    exam_col = find_exam_number_column(mastersheet)
    if not exam_col:
        print("❌ Could not find exam number column for filtering withdrawn students")
        return mastersheet, []

    for idx, row in mastersheet.iterrows():
        exam_no = str(row[exam_col]).strip()
        if is_student_withdrawn(exam_no):
            withdrawal_history = get_withdrawal_history(exam_no)
            # Only remove if student was withdrawn in a PREVIOUS semester
            if withdrawal_history and withdrawal_history['withdrawn_semester'] != semester_key:
                removed_students.append(exam_no)
                filtered_mastersheet = filtered_mastersheet[filtered_mastersheet[exam_col].astype(str) != exam_no]

    if removed_students:
        print(
            f"🚫 Removed {len(removed_students)} previously withdrawn students from {semester_key}:")
        for exam_no in removed_students:
            withdrawal_history = get_withdrawal_history(exam_no)
            print(
                f"   - {exam_no} (withdrawn in {withdrawal_history['withdrawn_semester']})")

    return filtered_mastersheet, removed_students

# ----------------------------
# Set Selection Functions
# ----------------------------

def get_available_sets(base_dir):
    """Get all available ND sets (ND-2024, ND-2025, etc.) from the ND folder"""
    # UPDATED: Look in the ND subdirectory
    nd_dir = os.path.join(base_dir, "ND")
    if not os.path.exists(nd_dir):
        print(f"❌ ND directory not found: {nd_dir}")
        return []
        
    sets = []
    for item in os.listdir(nd_dir):
        item_path = os.path.join(nd_dir, item)
        if os.path.isdir(item_path) and item.upper().startswith("ND-"):
            sets.append(item)
    return sorted(sets)

def get_user_set_choice(available_sets):
    """
    Prompt user to choose which set to process.
    Returns the selected set directory name.
    """
    print("\n🎯 AVAILABLE SETS:")
    for i, set_name in enumerate(available_sets, 1):
        print(f"{i}. {set_name}")
    print(f"{len(available_sets) + 1}. Process ALL sets")

    while True:
        try:
            choice = input(
                f"\nEnter your choice (1-{len(available_sets) + 1}): ").strip()
            if not choice:
                print("❌ Please enter a choice.")
                continue

            if choice.isdigit():
                choice_num = int(choice)
                if 1 <= choice_num <= len(available_sets):
                    selected_set = available_sets[choice_num - 1]
                    print(f"✅ Selected set: {selected_set}")
                    return [selected_set]
                elif choice_num == len(available_sets) + 1:
                    print("✅ Selected: ALL sets")
                    return available_sets
                else:
                    print(
                        f"❌ Invalid choice. Please enter a number between 1-{len(available_sets) + 1}.")
            else:
                print("❌ Please enter a valid number.")

        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)
        except Exception as e:
            print(f"❌ Error: {e}. Please try again.")

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
    except BaseException:
        return 'F'

def get_grade_point(score):
    """Convert score to grade point for GPA calculation - NIGERIAN 5.0 SCALE."""
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
    except BaseException:
        return 0.0

# ----------------------------
# Load Course Data
# ----------------------------

def load_course_data():
    """
    Reads course-code-creditUnit.xlsx and returns:
      (semester_course_maps, semester_credit_units,
       semester_lookup, semester_course_titles)
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
        df = pd.read_excel(
            course_file,
            sheet_name=sheet,
            engine='openpyxl',
            header=0)
        df.columns = [str(c).strip() for c in df.columns]
        expected = ['COURSE CODE', 'COURSE TITLE', 'CU']
        if not all(col in df.columns for col in expected):
            print(
                f"Warning: sheet '{sheet}' missing expected columns {expected} — skipped")
            continue
        dfx = df.dropna(subset=['COURSE CODE', 'COURSE TITLE'])
        dfx = dfx[~dfx['COURSE CODE'].astype(
            str).str.contains('TOTAL', case=False, na=False)]
        valid_mask = dfx['CU'].astype(str).str.replace(
            '.', '', regex=False).str.isdigit()
        dfx = dfx[valid_mask]
        if dfx.empty:
            print(
                f"Warning: sheet '{sheet}' has no valid rows after cleaning — skipped")
            continue
        codes = dfx['COURSE CODE'].astype(str).str.strip().tolist()
        titles = dfx['COURSE TITLE'].astype(str).str.strip().tolist()
        cus = dfx['CU'].astype(float).astype(int).tolist()

        # Create enhanced course mapping with normalized titles
        enhanced_course_map = {}
        for title, code in zip(titles, codes):
            normalized_title = normalize_course_name(title)
            enhanced_course_map[normalized_title] = {
                'original_name': title,
                'code': code,
                'normalized': normalized_title
            }
            
        semester_course_maps[sheet] = enhanced_course_map
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
    Returns: (
    semester_key,
    year,
    semester_num,
    level_display,
    semester_display,
     set_code)
    """
    filename_upper = filename.upper()

    # Map filename patterns to actual course sheet names
    if 'FIRST-YEAR-FIRST-SEMESTER' in filename_upper or 'FIRST_YEAR_FIRST_SEMESTER' in filename_upper or 'FIRST SEMESTER' in filename_upper:
        return "ND-FIRST-YEAR-FIRST-SEMESTER", 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
    elif 'FIRST-YEAR-SECOND-SEMESTER' in filename_upper or 'FIRST_YEAR_SECOND_SEMESTER' in filename_upper or 'SECOND SEMESTER' in filename_upper:
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
        print(
            f"⚠️ Could not detect semester from filename: {filename}, defaulting to ND-FIRST-YEAR-FIRST-SEMESTER")
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
    norm_map = {col: re.sub(r'\s+', ' ', str(col).strip().lower())
                for col in df.columns}
    candidates = [re.sub(r'\s+', ' ', c.strip().lower())
                  for c in candidate_names]
    for cand in candidates:
        for col, ncol in norm_map.items():
            if ncol == cand:
                return col
    return None

def find_exam_number_column(df):
    """Find the exam number column in a DataFrame with flexible matching."""
    possible_names = ['EXAM NUMBER', 'Reg No', 'REG NO', 'REGISTRATION NUMBER', 'MAT NO', 'STUDENT ID']
    
    for col in df.columns:
        col_upper = str(col).upper()
        for possible_name in possible_names:
            if possible_name in col_upper:
                return col
    return None

def load_previous_gpas_from_processed_files(
        output_dir, current_semester_key, timestamp):
    """
    Load previous GPA data from previously processed mastersheets in the same run.
    Returns dict: {exam_number: previous_gpa}
    """
    previous_gpas = {}

    print(f"\n🔍 LOADING PREVIOUS GPA for: {current_semester_key}")

    # Determine previous semester based on current
    current_year, current_semester_num, _, _, _ = get_semester_display_info(
        current_semester_key)

    if current_semester_num == 1 and current_year == 1:
        # First semester of first year - no previous GPA
        print("📊 First semester of first year - no previous GPA available")
        return previous_gpas
    elif current_semester_num == 2 and current_year == 1:
        # Second semester of first year - look for first semester of first year
        prev_semester = "ND-FIRST-YEAR-FIRST-SEMESTER"
    elif current_semester_num == 1 and current_year == 2:
        # First semester of second year - look for second semester of first
        # year
        prev_semester = "ND-FIRST-YEAR-SECOND-SEMESTER"
    elif current_semester_num == 2 and current_year == 2:
        # Second semester of second year - look for first semester of second
        # year
        prev_semester = "ND-SECOND-YEAR-FIRST-SEMESTER"
    else:
        print(
            f"⚠️ Unknown semester combination: Year {current_year}, Semester {current_semester_num}")
        return previous_gpas

    print(f"🔍 Looking for previous GPA data from: {prev_semester}")

    # Look for the mastersheet file from the previous semester in the same
    # timestamp directory
    mastersheet_pattern = os.path.join(
        output_dir,
        f"mastersheet_{timestamp}.xlsx")
    print(f"🔍 Checking for mastersheet: {mastersheet_pattern}")

    if os.path.exists(mastersheet_pattern):
        print(f"✅ Found mastersheet: {mastersheet_pattern}")
        try:
            # Read the Excel file properly, skipping the header rows that
            # contain merged cells
            df = pd.read_excel(
                mastersheet_pattern,
                sheet_name=prev_semester,
                header=5)  # Skip first 5 rows

            print(f"📋 Columns in {prev_semester}: {df.columns.tolist()}")

            # Find the actual column names by checking for exam number and GPA
            # columns
            exam_col = None
            gpa_col = None

            for col in df.columns:
                col_str = str(col).upper().strip()
                if 'EXAM' in col_str or 'REG' in col_str or 'NUMBER' in col_str:
                    exam_col = col
                elif 'GPA' in col_str:
                    gpa_col = col

            if exam_col and gpa_col:
                print(
                    f"✅ Found exam column: '{exam_col}', GPA column: '{gpa_col}'")

                gpas_loaded = 0
                for idx, row in df.iterrows():
                    exam_no = str(row[exam_col]).strip()
                    gpa = row[gpa_col]

                    if pd.notna(gpa) and pd.notna(
                            exam_no) and exam_no != 'nan' and exam_no != '':
                        try:
                            previous_gpas[exam_no] = float(gpa)
                            gpas_loaded += 1
                            if gpas_loaded <= 5:  # Show first 5 for debugging
                                print(f"📝 Loaded GPA: {exam_no} → {gpa}")
                        except (ValueError, TypeError):
                            continue

                print(
                    f"✅ Loaded previous GPAs for {gpas_loaded} students from {prev_semester}")

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

def load_all_previous_gpas_for_cgpa(
        output_dir,
        current_semester_key,
        timestamp):
    """
    Load ALL previous GPAs from all completed semesters for CGPA calculation.
    Returns dict: {exam_number: {'gpas': [gpa1, gpa2, ...], 'credits': [credits1, credits2, ...]}}
    """
    print(
        f"\n🔍 LOADING ALL PREVIOUS GPAs for CGPA calculation: {current_semester_key}")

    current_year, current_semester_num, _, _, _ = get_semester_display_info(
        current_semester_key)

    # Determine which semesters to load based on current semester
    semesters_to_load = []

    if current_semester_num == 1 and current_year == 1:
        # First semester - no previous data
        return {}
    elif current_semester_num == 2 and current_year == 1:
        # Second semester of first year - load first semester
        semesters_to_load = ["ND-FIRST-YEAR-FIRST-SEMESTER"]
    elif current_semester_num == 1 and current_year == 2:
        # First semester of second year - load both first year semesters
        semesters_to_load = [
            "ND-FIRST-YEAR-FIRST-SEMESTER",
            "ND-FIRST-YEAR-SECOND-SEMESTER"]
    elif current_semester_num == 2 and current_year == 2:
        # Second semester of second year - load all previous semesters
        semesters_to_load = [
            "ND-FIRST-YEAR-FIRST-SEMESTER",
            "ND-FIRST-YEAR-SECOND-SEMESTER",
            "ND-SECOND-YEAR-FIRST-SEMESTER"
        ]

    print(f"📚 Semesters to load for CGPA: {semesters_to_load}")

    all_student_data = {}
    mastersheet_path = os.path.join(
        output_dir,
        f"mastersheet_{timestamp}.xlsx")

    if not os.path.exists(mastersheet_path):
        print(f"❌ Mastersheet not found: {mastersheet_path}")
        return {}

    for semester in semesters_to_load:
        print(f"📖 Loading data from: {semester}")
        try:
            # Load the semester data, skipping header rows
            df = pd.read_excel(mastersheet_path, sheet_name=semester, header=5)

            # Find columns
            exam_col = None
            gpa_col = None
            credit_col = None

            for col in df.columns:
                col_str = str(col).upper().strip()
                if 'EXAM' in col_str or 'REG' in col_str or 'NUMBER' in col_str:
                    exam_col = col
                elif 'GPA' in col_str:
                    gpa_col = col
                elif 'CU PASSED' in col_str or 'CREDIT' in col_str:
                    credit_col = col

            if exam_col and gpa_col:
                for idx, row in df.iterrows():
                    exam_no = str(row[exam_col]).strip()
                    gpa = row[gpa_col]

                    if pd.notna(gpa) and pd.notna(
                            exam_no) and exam_no != 'nan' and exam_no != '':
                        try:
                            # Get credits completed (use CU Passed if
                            # available, otherwise estimate)
                            credits_completed = 0
                            if credit_col and pd.notna(row[credit_col]):
                                credits_completed = int(row[credit_col])
                            else:
                                # Estimate credits based on typical semester
                                # load
                                if 'FIRST-YEAR-FIRST-SEMESTER' in semester:
                                    credits_completed = 30  # Typical first semester credits
                                elif 'FIRST-YEAR-SECOND-SEMESTER' in semester:
                                    credits_completed = 30  # Typical second semester credits
                                elif 'SECOND-YEAR-FIRST-SEMESTER' in semester:
                                    credits_completed = 30  # Typical third semester credits

                            if exam_no not in all_student_data:
                                all_student_data[exam_no] = {
                                    'gpas': [], 'credits': []}

                            all_student_data[exam_no]['gpas'].append(
                                float(gpa))
                            all_student_data[exam_no]['credits'].append(
                                credits_completed)

                        except (ValueError, TypeError):
                            continue

        except Exception as e:
            print(f"⚠️ Could not load data from {semester}: {str(e)}")

    print(f"📊 Loaded cumulative data for {len(all_student_data)} students")
    return all_student_data

def calculate_cgpa(student_data, current_gpa, current_credits):
    """
    Calculate Cumulative GPA (CGPA) based on all previous semesters and current semester.
    """
    if not student_data:
        return current_gpa

    total_grade_points = 0.0
    total_credits = 0

    # Add previous semesters
    for prev_gpa, prev_credits in zip(
            student_data['gpas'], student_data['credits']):
        total_grade_points += prev_gpa * prev_credits
        total_credits += prev_credits

    # Add current semester
    total_grade_points += current_gpa * current_credits
    total_credits += current_credits

    if total_credits > 0:
        return round(total_grade_points / total_credits, 2)
    else:
        return current_gpa

def get_cumulative_gpa(
        current_gpa,
        previous_gpa,
        current_credits,
        previous_credits):
    """
    Calculate cumulative GPA based on current and previous semester performance.
    """
    if previous_gpa is None:
        return current_gpa

    # For simplicity, we'll assume equal credit weights if not provided
    if current_credits is None or previous_credits is None:
        return round((current_gpa + previous_gpa) / 2, 2)

    total_points = (current_gpa * current_credits) + \
        (previous_gpa * previous_credits)
    total_credits = current_credits + previous_credits
    return round(total_points / total_credits, 2) if total_credits > 0 else 0.0

def determine_student_status(row, total_cu, pass_threshold):
    """
    Determine student status based on performance metrics.
    Returns: 'Pass', 'Carry Over', 'Probation', or 'Withdrawn'
    """
    gpa = row.get("GPA", 0)
    cu_passed = row.get("CU Passed", 0)
    cu_failed = row.get("CU Failed", 0)

    # Calculate percentage of failed credit units
    failed_percentage = (cu_failed / total_cu) * 100 if total_cu > 0 else 0

    # Decision matrix based on the summary criteria
    if cu_failed == 0:
        return "Pass"
    elif gpa >= 2.0 and failed_percentage <= 45:
        return "Carry Over"
    elif gpa < 2.0 and failed_percentage <= 45:
        return "Probation"
    elif failed_percentage > 45:
        return "Withdrawn"
    else:
        return "Carry Over"

def format_failed_courses_remark(failed_courses, max_line_length=60):
    """
    Format failed courses remark with line breaks for long lists.
    Returns list of formatted lines.
    """
    if not failed_courses:
        return [""]

    failed_str = ", ".join(sorted(failed_courses))

    # If the string is short enough, return as single line
    if len(failed_str) <= max_line_length:
        return [failed_str]

    # Split into multiple lines
    lines = []
    current_line = ""

    for course in sorted(failed_courses):
        if not current_line:
            current_line = course
        elif len(current_line) + len(course) + 2 <= max_line_length:  # +2 for ", "
            current_line += ", " + course
        else:
            lines.append(current_line)
            current_line = course

    if current_line:
        lines.append(current_line)

    return lines

def get_user_semester_choice():
    """
    Prompt user to choose which semesters to process.
    Returns list of semester keys to process.
    """
    print("\n🎯 SEMESTER PROCESSING OPTIONS:")
    print("1. Process ALL semesters in order")
    print("2. Process FIRST YEAR - FIRST SEMESTER only")
    print("3. Process FIRST YEAR - SECOND SEMESTER only")
    print("4. Process SECOND YEAR - FIRST SEMESTER only")
    print("5. Process SECOND YEAR - SECOND SEMESTER only")
    print("6. Custom selection")

    while True:
        try:
            choice = input("\nEnter your choice (1-6): ").strip()
            if choice == "1":
                return SEMESTER_ORDER.copy()
            elif choice == "2":
                return ["ND-FIRST-YEAR-FIRST-SEMESTER"]
            elif choice == "3":
                return ["ND-FIRST-YEAR-SECOND-SEMESTER"]
            elif choice == "4":
                return ["ND-SECOND-YEAR-FIRST-SEMESTER"]
            elif choice == "5":
                return ["ND-SECOND-YEAR-SECOND-SEMESTER"]
            elif choice == "6":
                return get_custom_semester_selection()
            else:
                print("❌ Invalid choice. Please enter a number between 1-6.")
        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)
        except Exception as e:
            print(f"❌ Error: {e}. Please try again.")

def get_custom_semester_selection():
    """
    Allow user to select multiple semesters for processing.
    """
    print("\n📚 AVAILABLE SEMESTERS:")
    for i, semester in enumerate(SEMESTER_ORDER, 1):
        year, sem_num, level, sem_display, set_code = get_semester_display_info(
            semester)
        print(f"{i}. {level} - {sem_display}")

    print(f"{len(SEMESTER_ORDER) + 1}. Select all")

    selected = []
    while True:
        try:
            choices = input(
                f"\nEnter semester numbers separated by commas (1-{len(SEMESTER_ORDER) + 1}): ").strip()
            if not choices:
                print("❌ Please enter at least one semester number.")
                continue

            choice_list = [c.strip() for c in choices.split(',')]

            # Check for "select all" option
            if str(len(SEMESTER_ORDER) + 1) in choice_list:
                return SEMESTER_ORDER.copy()

            # Validate and convert choices
            valid_choices = []
            for choice in choice_list:
                if not choice.isdigit():
                    print(f"❌ '{choice}' is not a valid number.")
                    continue

                choice_num = int(choice)
                if 1 <= choice_num <= len(SEMESTER_ORDER):
                    valid_choices.append(choice_num)
                else:
                    print(f"❌ '{choice}' is not a valid semester number.")

            if valid_choices:
                selected_semesters = [SEMESTER_ORDER[i - 1]
                                      for i in valid_choices]
                print(
                    f"✅ Selected semesters: {[get_semester_display_info(sem)[3] for sem in selected_semesters]}")
                return selected_semesters
            else:
                print("❌ No valid semesters selected. Please try again.")

        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)
        except Exception as e:
            print(f"❌ Error: {e}. Please try again.")

# ----------------------------
# PDF Generation - Individual Student Report (Enhanced with Resit Info)
# ----------------------------

def generate_individual_student_pdf(
    mastersheet_df,
    out_pdf_path,
    semester_key,
    logo_path=None,
    prev_mastersheet_df=None,
    filtered_credit_units=None,
    ordered_codes=None,
    course_titles_map=None,
    previous_gpas=None,
    cgpa_data=None,
    total_cu=None,
    pass_threshold=None,
    upgrade_min_threshold=None,
    resit_count=0):
    """
    Create a PDF with one page per student matching the sample format exactly.
    Enhanced to include resit information.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

    doc = SimpleDocTemplate(
        out_pdf_path,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=20,
        bottomMargin=20)
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

    # Style for remarks with smaller font
    remarks_style = ParagraphStyle(
        'RemarksStyle',
        parent=styles['Normal'],
        fontSize=8,
        alignment=TA_LEFT
    )

    elems = []

    for idx, r in mastersheet_df.iterrows():
        # Logo and header
        logo_img = None
        if logo_path and os.path.exists(logo_path):
            try:
                logo_img = Image(
                    logo_path,
                    width=0.8 * inch,
                    height=0.8 * inch)
            except Exception as e:
                print(f"Warning: Could not load logo: {e}")

        # Header table with logo and title
        if logo_img:
            header_data = [[logo_img, Paragraph(
                "FCT COLLEGE OF NURSING SCIENCES", main_header_style)]]
            header_table = Table(
                header_data, colWidths=[
                    1.0 * inch, 5.0 * inch])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            ]))
            elems.append(header_table)
        else:
            elems.append(
                Paragraph(
                    "FCT COLLEGE OF NURSING SCIENCES",
                    main_header_style))

        # Address and contact info
        elems.append(
            Paragraph(
                "P.O.Box 507, Gwagwalada-Abuja, Nigeria",
                header_style))
        elems.append(Paragraph("<b>EXAMINATIONS OFFICE</b>", header_style))
        elems.append(Paragraph("fctsonexamsoffice@gmail.com", header_style))

        elems.append(Spacer(1, 8))
        elems.append(
            Paragraph(
                "STUDENT'S ACADEMIC PROGRESS REPORT",
                title_style))
        elems.append(Paragraph("(THIS IS NOT A TRANSCRIPT)", subtitle_style))

        elems.append(Spacer(1, 8))

        # Student particulars - SEPARATE FROM PASSPORT PHOTO
        exam_no = str(r.get("EXAM NUMBER", r.get("REG. No", "")))
        student_name = str(r.get("NAME", ""))

        # Determine level and semester using the new function
        year, semester_num, level_display, semester_display, set_code = get_semester_display_info(
            semester_key)

        # Create two tables: one for student particulars, one for passport
        # photo
        particulars_data = [
            [Paragraph("<b>STUDENT'S PARTICULARS</b>", styles['Normal'])],
            [Paragraph("<b>NAME:</b>", styles['Normal']), student_name],
            [Paragraph("<b>LEVEL OF<br/>STUDY:</b>", styles['Normal']), level_display,
             Paragraph("<b>SEMESTER:</b>", styles['Normal']), semester_display],
            [Paragraph("<b>REG NO.</b>", styles['Normal']), exam_no,
             Paragraph("<b>SET:</b>", styles['Normal']), set_code],
        ]

        particulars_table = Table(
            particulars_data,
            colWidths=[
                1.2 * inch,
                2.3 * inch,
                0.8 * inch,
                1.5 * inch])
        particulars_table.setStyle(TableStyle([
            ('SPAN', (0, 0), (3, 0)),
            ('SPAN', (1, 1), (3, 1)),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))

        # Passport photo table (separate box)
        passport_data = [
            [Paragraph("Affix Recent<br/>Passport<br/>Photograph",
                       styles['Normal'])]
        ]

        passport_table = Table(
            passport_data,
            colWidths=[
                1.5 * inch],
            rowHeights=[
                1.2 * inch])
        passport_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
        ]))

        # Create a combined table with particulars and passport side by side
        combined_data = [
            [particulars_table, passport_table]
        ]

        combined_table = Table(
            combined_data, colWidths=[
                5.8 * inch, 1.5 * inch])
        combined_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
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
        failed_courses_list = []

        for code in ordered_codes if ordered_codes else []:
            score = r.get(code)
            if pd.isna(score) or score == "":
                continue

            try:
                score_val = float(score)

                # -----------------------
                # FIX: Auto-upgrade borderline scores when threshold upgrade applies
                # This ensures PDF shows the same upgraded scores as Excel
                if upgrade_min_threshold is not None and upgrade_min_threshold <= score_val <= 49:
                    # Use the upgraded score for PDF display
                    score_val = 50.0
                    score_display = "50"
                    print(f"🔼 PDF: Upgraded score for {exam_no} - {code}: {original_score} → 50")
                else:
                    score_display = str(int(round(score_val)))
                # -----------------------

                grade = get_grade(score_val)
                grade_point = get_grade_point(score_val)
            except Exception:
                score_display = str(score)
                grade = "F"
                grade_point = 0.0

            cu = filtered_credit_units.get(code, 0) if filtered_credit_units else 0
            course_title = course_titles_map.get(code, code) if course_titles_map else code

            total_grade_points += grade_point * cu
            total_units += cu

            if score_val >= pass_threshold:
                total_units_passed += cu
            else:
                total_units_failed += cu
                failed_courses_list.append(code)

            course_data.append([
                Paragraph(str(sn), center_align_style),
                Paragraph(code, left_align_style),
                Paragraph(course_title, left_align_style),
                Paragraph(str(cu), center_align_style),
                Paragraph(score_display, center_align_style),
                Paragraph(grade, center_align_style)
            ])
            sn += 1

        course_table = Table(
            course_data,
            colWidths=[
                0.4 * inch,
                0.7 * inch,
                2.8 * inch,
                0.6 * inch,
                0.6 * inch,
                0.6 * inch])
        course_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E0E0E0")),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('ALIGN', (1, 1), (2, -1), 'LEFT'),
        ]))
        elems.append(course_table)
        elems.append(Spacer(1, 14))

        # Calculate current semester GPA
        current_gpa = round(
            total_grade_points / total_units,
            2) if total_units > 0 else 0.0

        # Get previous GPA if available
        exam_no = str(r.get("EXAM NUMBER", "")).strip()
        previous_gpa = previous_gpas.get(
            exam_no, None) if previous_gpas else None

        # Calculate CGPA if available
        cgpa = None
        if cgpa_data and exam_no in cgpa_data:
            cgpa = calculate_cgpa(
                cgpa_data[exam_no],
                current_gpa,
                total_units_passed)

        print(f"📊 PDF GENERATION for {exam_no}:")
        print(f"   Current GPA: {current_gpa}")
        print(f"   Previous GPA available: {previous_gpa is not None}")
        print(f"   CGPA available: {cgpa is not None}")
        if previous_gpa is not None:
            print(f"   Previous GPA value: {previous_gpa}")
        if cgpa is not None:
            print(f"   CGPA value: {cgpa}")

        # Get values from dataframe
        tcpe = round(total_grade_points, 1)
        tcup = total_units_passed
        tcuf = total_units_failed

        # Determine student status based on performance
        student_status = determine_student_status(r, total_cu, pass_threshold)

        # Check if student was previously withdrawn
        withdrawal_history = get_withdrawal_history(exam_no)
        previously_withdrawn = withdrawal_history is not None

        # Format failed courses with line breaks if needed
        failed_courses_formatted = format_failed_courses_remark(
            failed_courses_list)

        # Combine course-specific remarks with overall status
        final_remarks_lines = []

        # Add resit information to remarks if applicable
        if resit_count > 0:
            final_remarks_lines.append(f"Resit Attempts: {resit_count}")

        # For withdrawn students in their withdrawal semester, show appropriate
        # metrics
        if previously_withdrawn and withdrawal_history['withdrawn_semester'] == semester_key:
            # This is the actual withdrawal semester - show normal withdrawal
            # remarks
            if failed_courses_formatted:
                final_remarks_lines.append(
                    f"Failed: {failed_courses_formatted[0]}")
                if len(failed_courses_formatted) > 1:
                    final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("Advised to Withdraw")
            else:
                final_remarks_lines.append("Advised to Withdraw")
        elif previously_withdrawn:
            # Student was withdrawn in a previous semester but appears here -
            # this shouldn't happen due to filtering
            withdrawn_semester = withdrawal_history['withdrawn_semester']
            year, sem_num, level, sem_display, set_code = get_semester_display_info(
                withdrawn_semester)
            final_remarks_lines.append(
                f"STUDENT WAS WITHDRAWN FROM {level} - {sem_display}")
            final_remarks_lines.append(
                "This result should not be processed as student was previously withdrawn")
        elif student_status == "Pass":
            final_remarks_lines.append("Passed")
        elif student_status == "Carry Over":
            if failed_courses_formatted:
                final_remarks_lines.append(
                    f"Failed: {failed_courses_formatted[0]}")
                if len(failed_courses_formatted) > 1:
                    final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("To Carry Over Courses")
            else:
                final_remarks_lines.append("To Carry Over Courses")
        elif student_status == "Probation":
            if failed_courses_formatted:
                final_remarks_lines.append(
                    f"Failed: {failed_courses_formatted[0]}")
                if len(failed_courses_formatted) > 1:
                    final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("Placed on Probation")
            else:
                final_remarks_lines.append("Placed on Probation")
        elif student_status == "Withdrawn":
            if failed_courses_formatted:
                final_remarks_lines.append(
                    f"Failed: {failed_courses_formatted[0]}")
                if len(failed_courses_formatted) > 1:
                    final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("Advised to Withdraw")
            else:
                final_remarks_lines.append("Advised to Withdraw")
        else:
            final_remarks_lines.append(str(r.get("REMARKS", "")))

        final_remarks = "<br/>".join(final_remarks_lines)
        display_gpa = current_gpa
        display_cgpa = cgpa if cgpa is not None else current_gpa

        # Summary section - EXPANDED TO ACCOMMODATE LONG REMARKS
        summary_data = [
            [Paragraph("<b>SUMMARY</b>", styles['Normal']), "", "", ""],
            [Paragraph("<b>TCPE:</b>", styles['Normal']), str(tcpe),
             Paragraph("<b>CURRENT GPA:</b>", styles['Normal']), str(display_gpa)],
        ]

        # Add previous GPA if available (from first year second semester
        # upward)
        if previous_gpa is not None:
            print(f"✅ ADDING PREVIOUS GPA to PDF: {previous_gpa}")
            summary_data.append([
                Paragraph("<b>TCUP:</b>", styles['Normal']), str(tcup),
                Paragraph("<b>PREVIOUS GPA:</b>", styles['Normal']), str(previous_gpa)
            ])
        else:
            summary_data.append([
                Paragraph("<b>TCUP:</b>", styles['Normal']), str(tcup), "", ""
            ])

        # Add CGPA if available (from second semester onward)
        if cgpa is not None:
            print(f"✅ ADDING CGPA to PDF: {cgpa}")
            summary_data.append([
                Paragraph("<b>TCUF:</b>", styles['Normal']), str(tcuf),
                Paragraph("<b>OVERALL GPA:</b>", styles['Normal']), str(display_cgpa)
            ])
        else:
            summary_data.append([
                Paragraph("<b>TCUF:</b>", styles['Normal']), str(tcuf), "", ""
            ])

        # Add resit count to summary
        if resit_count > 0:
            summary_data.append([
                Paragraph("<b>RESIT COUNT:</b>", styles['Normal']), str(resit_count), "", ""
            ])

        # Add remarks with multiple lines if needed
        remarks_paragraph = Paragraph(final_remarks, remarks_style)
        summary_data.append([
            Paragraph("<b>REMARKS:</b>", styles['Normal']),
            remarks_paragraph, "", ""
        ])

        # Calculate row heights based on content
        row_heights = [0.3 * inch] * len(summary_data)  # Default height

        # Adjust height for remarks row based on number of lines
        total_remark_lines = len(final_remarks_lines)
        if total_remark_lines > 1:
            # Add extra height for multiple lines
            row_heights[-1] = max(0.4 * inch,
                                  0.2 * inch * (total_remark_lines + 1))

        summary_table = Table(
            summary_data,
            colWidths=[
                1.5 * inch,
                1.0 * inch,
                1.5 * inch,
                1.0 * inch],
            rowHeights=row_heights)
        summary_table.setStyle(TableStyle([
            ('SPAN', (0, 0), (3, 0)),
            ('SPAN', (1, len(summary_data) - 1), (3, len(summary_data) - 1)),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (3, 0), colors.HexColor("#E0E0E0")),
            ('ALIGN', (0, 0), (3, 0), 'CENTER'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elems.append(summary_table)
        elems.append(Spacer(1, 25))

        # Signature section
        sig_data = [["",
                     ""],
                    ["____________________",
                     "____________________"],
                    [Paragraph("<b>EXAMS SECRETARY</b>",
                               ParagraphStyle('SigStyle',
                                              parent=styles['Normal'],
                                              fontSize=10,
                                              alignment=TA_CENTER)),
                     Paragraph("<b>V.P. ACADEMICS</b>",
                       ParagraphStyle('SigStyle',
                                     parent=styles['Normal'],
                                     fontSize=10,
                                     alignment=TA_CENTER))] ]

        sig_table = Table(sig_data, colWidths=[3.0 * inch, 3.0 * inch])
        sig_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        elems.append(sig_table)

        # Page break for next student
        if idx < len(mastersheet_df) - 1:
            elems.append(PageBreak())

    doc.build(elems)
    print(f"✅ Individual student PDF written: {out_pdf_path}")

# ----------------------------
# Main file processing (Enhanced with Data Transformation)
# ----------------------------

def process_single_file(
        path,
        output_dir,
        ts,
        pass_threshold,
        semester_course_maps,
        semester_credit_units,
        semester_lookup,
        semester_course_titles,
        logo_path,
        semester_key,
        set_name,
        previous_gpas,
        cgpa_data=None,
        upgrade_min_threshold=None):
    """
    Process a single raw file and produce mastersheet Excel and PDFs.
    Enhanced with data transformation for transposed formats.
    """
    fname = os.path.basename(path)
    print(f"🔍 Processing file: {fname} for semester: {semester_key}")

    try:
        xl = pd.ExcelFile(path)
        print(f"✅ Successfully opened Excel file: {fname}")
        print(f"📋 Sheets found: {xl.sheet_names}")
    except Exception as e:
        print(f"❌ Error opening excel {path}: {e}")
        return None

    expected_sheets = ['CA', 'OBJ', 'EXAM']
    dfs = {}
    
    for s in expected_sheets:
        if s in xl.sheet_names:
            try:
                # Try reading with different parameters to handle various Excel formats
                dfs[s] = pd.read_excel(path, sheet_name=s, dtype=str, header=0)
                print(f"✅ Loaded sheet {s} with shape: {dfs[s].shape}")
                print(f"📊 Sheet {s} columns: {dfs[s].columns.tolist()}")
                
                # NEW: Check if data is in transposed format and transform if needed
                if detect_data_format(dfs[s], s):
                    print(f"🔄 Data in {s} sheet is in transposed format, transforming...")
                    dfs[s] = transform_transposed_data(dfs[s], s)
                    print(f"✅ Transformed {s} sheet to wide format")
                    print(f"📊 Transformed shape: {dfs[s].shape}")
                    print(f"📋 Transformed columns: {dfs[s].columns.tolist()}")
                
                # Debug: Show first few rows of data
                if not dfs[s].empty:
                    print(f"🔍 First 3 rows of {s} sheet:")
                    for i in range(min(3, len(dfs[s]))):
                        row_data = {}
                        for col in dfs[s].columns[:5]:  # Show first 5 columns
                            row_data[col] = dfs[s].iloc[i][col]
                        print(f"   Row {i}: {row_data}")
                else:
                    print(f"⚠️ Sheet {s} is empty!")
                    
            except Exception as e:
                print(f"❌ Error reading sheet {s}: {e}")
                # Try alternative reading method
                try:
                    dfs[s] = pd.read_excel(path, sheet_name=s, header=0)
                    print(f"✅ Alternative load successful for sheet {s}")
                except Exception as e2:
                    print(f"❌ Alternative load also failed for sheet {s}: {e2}")
                    dfs[s] = pd.DataFrame()
        else:
            print(f"⚠️ Sheet {s} not found in {fname}")
            dfs[s] = pd.DataFrame()
            
    if not dfs:
        print("❌ No CA/OBJ/EXAM sheets detected — skipping file.")
        return None

    # Use the provided semester key
    sem = semester_key
    year, semester_num, level_display, semester_display, set_code = get_semester_display_info(
        sem)
    print(
        f"📁 Processing: {level_display} - {semester_display} - Set: {set_code}")
    print(f"📊 Using course sheet: {sem}")

    print(f"📊 Previous GPAs provided: {len(previous_gpas)} students")
    print(
        f"📊 CGPA data available for: {len(cgpa_data) if cgpa_data else 0} students")

    # Check if semester exists in course maps
    if sem not in semester_course_maps:
        print(
            f"❌ Semester '{sem}' not found in course data. Available semesters: {list(semester_course_maps.keys())}")
        return None

    course_map = semester_course_maps[sem]
    credit_units = semester_credit_units[sem]
    course_titles = semester_course_titles[sem]

    ordered_titles = list(course_map.keys())
    ordered_codes = [course_map[t]['code']
                     for t in ordered_titles if course_map.get(t)]
    ordered_codes = [c for c in ordered_codes if credit_units.get(c, 0) > 0]
    filtered_credit_units = {c: credit_units[c] for c in ordered_codes}
    total_cu = sum(filtered_credit_units.values())

    print(f"📚 Course codes to process: {ordered_codes}")
    print(f"📊 Total credit units: {total_cu}")

    reg_no_cols = {s: find_column_by_names(df,
                                           ["REG. No",
                                            "Reg No",
                                            "Registration Number",
                                            "Mat No",
                                            "EXAM NUMBER",
                                            "Student ID"]) for s,
                   df in dfs.items()}
    name_cols = {
        s: find_column_by_names(
            df, [
                "NAME", "Full Name", "Candidate Name"]) for s, df in dfs.items()}

    print(f"🔍 Registration columns found: {reg_no_cols}")
    print(f"🔍 Name columns found: {name_cols}")

    merged = None
    for s, df in dfs.items():
        if df.empty:
            print(f"⚠️ Skipping empty sheet: {s}")
            continue
            
        df = df.copy()
        regcol = reg_no_cols.get(s)
        namecol = name_cols.get(s)
        if not regcol:
            regcol = df.columns[0] if len(df.columns) > 0 else None
        if not namecol and len(df.columns) > 1:
            namecol = df.columns[1]

        if regcol is None:
            print(f"❌ Skipping sheet {s}: no reg column found")
            continue

        print(f"📝 Processing sheet {s} with reg column: {regcol}, name column: {namecol}")

        df["REG. No"] = df[regcol].astype(str).str.strip()
        if namecol:
            df["NAME"] = df[namecol].astype(str).str.strip()
        else:
            df["NAME"] = pd.NA

        to_drop = [
            c for c in [
                regcol,
                namecol] if c and c not in [
                "REG. No",
                "NAME"]]
        df.drop(columns=to_drop, errors="ignore", inplace=True)

        # Debug: Show available columns for matching
        print(f"🔍 Available columns in {s} sheet: {df.columns.tolist()}")

        # ENHANCED COURSE MATCHING - Use the new matching algorithm
        for col in [c for c in df.columns if c not in ["REG. No", "NAME"]]:
            matched_course = find_best_course_match(col, course_map)
            if matched_course:
                matched_code = matched_course['code']
                newcol = f"{matched_code}_{s.upper()}"
                df.rename(columns={col: newcol}, inplace=True)
                print(f"✅ Matched column '{col}' to course code '{matched_code}' (original: {matched_course['original_name']})")
            else:
                print(f"❌ No match found for column: '{col}'")

        cur_cols = ["REG. No", "NAME"] + \
            [c for c in df.columns if c.endswith(f"_{s.upper()}")]
        cur = df[cur_cols].copy()
        
        # Debug: Show data before merging
        print(f"📊 Data in {s} sheet - Shape: {cur.shape}")
        if not cur.empty:
            print(f"🔍 First 3 rows of {s} data:")
            for i in range(min(2, len(cur))):
                print(f"   Row {i}: REG. No='{cur.iloc[i]['REG. No']}', NAME='{cur.iloc[i]['NAME']}'")
        
        if merged is None:
            merged = cur
            print(f"✅ Initialized merged dataframe with {s} sheet")
        else:
            print(f"🔗 Merging {s} sheet with existing data")
            before_merge = len(merged)
            merged = merged.merge(
                cur,
                on="REG. No",
                how="outer",
                suffixes=(
                    '',
                    '_dup'))
            after_merge = len(merged)
            print(f"📊 Merge result: {before_merge} -> {after_merge} rows")
            
            if "NAME_dup" in merged.columns:
                merged["NAME"] = merged["NAME"].combine_first(
                    merged["NAME_dup"])
                merged.drop(columns=["NAME_dup"], inplace=True)

    if merged is None or merged.empty:
        print("❌ No data merged from sheets — skipping file.")
        return None

    print(f"✅ Final merged dataframe shape: {merged.shape}")
    print(f"📋 Final merged columns: {merged.columns.tolist()}")

    # CRITICAL FIX: Check if we have actual score data before proceeding
    has_score_data = False
    score_columns = [col for col in merged.columns if any(code in col for code in ordered_codes)]
    print(f"🔍 Checking score columns: {score_columns}")
    
    for col in score_columns:
        if col in merged.columns:
            # Check if column has any non-null, non-zero values
            non_null_count = merged[col].notna().sum()
            if non_null_count > 0:
                # Try to convert to numeric and check for non-zero values
                try:
                    numeric_values = pd.to_numeric(merged[col], errors='coerce')
                    non_zero_count = (numeric_values > 0).sum()
                    if non_zero_count > 0:
                        has_score_data = True
                        print(f"✅ Found score data in column {col}: {non_zero_count} non-zero values")
                        break
                except Exception as e:
                    print(f"⚠️ Error checking column {col}: {e}")
    
    if not has_score_data:
        print(f"❌ CRITICAL: No valid score data found in file {fname}!")
        print(f"🔍 Sample of merged data:")
        print(merged.head(3))
        return None

    mastersheet = merged[["REG. No", "NAME"]].copy()
    mastersheet.rename(columns={"REG. No": "EXAM NUMBER"}, inplace=True)

    print("🎯 Calculating scores for each course...")
    
    for code in ordered_codes:
        ca_col = f"{code}_CA"
        obj_col = f"{code}_OBJ"
        exam_col = f"{code}_EXAM"

        print(f"📊 Processing course {code}:")
        print(f"   CA column: {ca_col} - exists: {ca_col in merged.columns}")
        print(f"   OBJ column: {obj_col} - exists: {obj_col in merged.columns}")
        print(f"   EXAM column: {exam_col} - exists: {exam_col in merged.columns}")

        ca_series = pd.to_numeric(
            merged[ca_col],
            errors="coerce") if ca_col in merged.columns else pd.Series(
            [0] * len(merged),
            index=merged.index)
        obj_series = pd.to_numeric(
            merged[obj_col],
            errors="coerce") if obj_col in merged.columns else pd.Series(
            [0] * len(merged),
            index=merged.index)
        exam_series = pd.to_numeric(
            merged[exam_col],
            errors="coerce") if exam_col in merged.columns else pd.Series(
            [0] * len(merged),
            index=merged.index)

        # Debug: Show score statistics
        print(f"   CA stats: non-null={ca_series.notna().sum()}, non-zero={(ca_series > 0).sum()}")
        print(f"   OBJ stats: non-null={obj_series.notna().sum()}, non-zero={(obj_series > 0).sum()}")
        print(f"   EXAM stats: non-null={exam_series.notna().sum()}, non-zero={(exam_series > 0).sum()}")

        ca_norm = (ca_series / 20) * 100
        obj_norm = (obj_series / 20) * 100
        exam_norm = (exam_series / 80) * 100
        ca_norm = ca_norm.fillna(0).clip(upper=100)
        obj_norm = obj_norm.fillna(0).clip(upper=100)
        exam_norm = exam_norm.fillna(0).clip(upper=100)
        total = (ca_norm * 0.2) + (((obj_norm + exam_norm) / 2) * 0.8)
        mastersheet[code] = total.round(0).clip(upper=100).values

        # Debug: Show final score statistics
        final_scores = mastersheet[code]
        print(f"   Final scores: non-zero={(final_scores > 0).sum()}, mean={final_scores.mean():.2f}")

    # NEW: APPLY FLEXIBLE UPGRADE RULE - Ask user for threshold per semester
    # Only ask in interactive mode
    if should_use_interactive_mode():
        upgrade_min_threshold, upgraded_scores_count = get_upgrade_threshold_from_user(semester_key, set_name)
    else:
        # In non-interactive mode, use the provided threshold or None
        upgraded_scores_count = 0
        if upgrade_min_threshold is not None:
            print(f"🔄 Applying upgrade upgrade from parameters: {upgrade_min_threshold}–49 → 50")

    if upgrade_min_threshold is not None:
        mastersheet, upgraded_scores_count = apply_upgrade_rule(mastersheet, ordered_codes, upgrade_min_threshold)

    for c in ordered_codes:
        if c not in mastersheet.columns:
            mastersheet[c] = 0

    def compute_remarks(row):
        """Compute remarks with expanded failed courses list."""
        fails = [c for c in ordered_codes if float(
            row.get(c, 0) or 0) < pass_threshold]
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
    results = mastersheet.apply(
        calc_tcpe_tcup_tcuf,
        axis=1,
        result_type='expand')
    mastersheet["TCPE"] = results[0].round(1)
    mastersheet["CU Passed"] = results[1]
    mastersheet["CU Failed"] = results[2]

    mastersheet["REMARKS"] = mastersheet.apply(compute_remarks, axis=1)

    total_cu = sum(filtered_credit_units.values()
                   ) if filtered_credit_units else 0

    # Calculate GPA - ALWAYS calculate the actual GPA
    def calculate_gpa(row):
        tcpe = row["TCPE"]
        return round((tcpe / total_cu), 2) if total_cu > 0 else 0.0

    mastersheet["GPA"] = mastersheet.apply(calculate_gpa, axis=1)
    mastersheet["AVERAGE"] = mastersheet[[
        c for c in ordered_codes]].mean(axis=1).round(0)

    # FILTER OUT PREVIOUSLY WITHDRAWN STUDENTS
    mastersheet, removed_students = filter_out_withdrawn_students(
        mastersheet, semester_key)

    # Identify withdrawn students in this semester (after filtering)
    withdrawn_students = []
    for idx, row in mastersheet.iterrows():
        student_status = determine_student_status(
            row, total_cu, pass_threshold)
        if student_status == "Withdrawn":
            exam_no = str(row["EXAM NUMBER"]).strip()
            withdrawn_students.append(exam_no)
            mark_student_withdrawn(exam_no, semester_key)
            print(f"🚫 Student {exam_no} marked as withdrawn in {semester_key}")

    # Update student tracker with current semester's students (after filtering)
    exam_numbers = mastersheet["EXAM NUMBER"].astype(str).str.strip().tolist()
    update_student_tracker(semester_key, exam_numbers, withdrawn_students)

    # Identify and save carryover students after processing
    carryover_students = identify_carryover_students(mastersheet, semester_key, set_name, pass_threshold)
    
    if carryover_students:
        carryover_dir = save_carryover_records(
            carryover_students, output_dir, set_name, semester_key
        )
        print(f"✅ Saved {len(carryover_students)} carryover records to: {carryover_dir}")
        
        # ADD: Log the carryover record file path for debugging
        carryover_file = os.path.join(carryover_dir, f"co_student_{set_name}_{semester_key}_*.json")
        print(f"📁 Carryover file pattern: {carryover_file}")
        
        # Print carryover summary
        total_failed_courses = sum(len(s['failed_courses']) for s in carryover_students)
        print(f"📊 Carryover Summary: {total_failed_courses} failed courses across all students")
        
        # Show most frequently failed courses
        course_fail_count = {}
        for student in carryover_students:
            for course in student['failed_courses']:
                course_code = course['course_code']
                course_fail_count[course_code] = course_fail_count.get(course_code, 0) + 1
        
        if course_fail_count:
            top_failed = sorted(course_fail_count.items(), key=lambda x: x[1], reverse=True)[:5]
            print(f"📚 Most failed courses: {top_failed}")
    else:
        print("✅ No carryover students identified")

    def sort_key(remark):
        if remark == "Passed":
            return (0, "")
        else:
            failed_courses = remark.replace("Failed: ", "").split(", ")
            return (1, len(failed_courses), ",".join(sorted(failed_courses)))
    mastersheet = mastersheet.sort_values(
        by="REMARKS",
        key=lambda x: x.map(sort_key)).reset_index(
        drop=True)

    if "S/N" not in mastersheet.columns:
        mastersheet.insert(0, "S/N", range(1, len(mastersheet) + 1))
    else:
        mastersheet["S/N"] = range(1, len(mastersheet) + 1)
        cols = list(mastersheet.columns)
        if cols[0] != "S/N":
            cols.remove("S/N")
            mastersheet = mastersheet[["S/N"] + cols]

    course_cols = ordered_codes
    out_cols = ["S/N", "EXAM NUMBER", "NAME"] + course_cols + \
        ["REMARKS", "CU Passed", "CU Failed", "TCPE", "GPA", "AVERAGE"]
    for c in out_cols:
        if c not in mastersheet.columns:
            mastersheet[c] = pd.NA
    mastersheet = mastersheet[out_cols]

    # FIXED: Create proper output directory structure - all files go directly to the set output directory
    out_xlsx = os.path.join(output_dir, f"mastersheet_{ts}.xlsx")

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
    title_cell.fill = PatternFill(
        start_color="1E90FF",
        end_color="1E90FF",
        fill_type="solid")
    border = Border(
        left=Side(
            style="medium"), right=Side(
            style="medium"), top=Side(
                style="medium"), bottom=Side(
                    style="medium"))
    title_cell.border = border

    # Use expanded semester name in the subtitle
    expanded_semester_name = f"{level_display} {semester_display}"

    ws.merge_cells("C2:Q2")
    subtitle_cell = ws["C2"]
    subtitle_cell.value = f"{datetime.now().year}/{datetime.now().year + 1} SESSION  NATIONAL DIPLOMA {expanded_semester_name} EXAMINATIONS RESULT — {datetime.now().strftime('%B %d, %Y')}"
    subtitle_cell.font = Font(bold=True, size=12, color="000000")
    subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")

    # FIXED: Remove the upgrade notice from the header to prevent covering course titles
    # The upgrade notice will only appear in the summary section
    start_row = 3

    display_course_titles = []
    for t in ordered_titles:
        course_info = course_map.get(t)
        if course_info and course_info['code'] in ordered_codes:
            display_course_titles.append(course_info['original_name'])

    ws.append([""] * 3 + display_course_titles + [""] * 5)
    for i, cell in enumerate(
            ws[start_row][3:3 + len(display_course_titles)], start=3):
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            text_rotation=45)
        cell.font = Font(bold=True, size=9)
    ws.row_dimensions[start_row].height = 18

    cu_list = [filtered_credit_units.get(c, "") for c in ordered_codes]
    ws.append([""] * 3 + cu_list + [""] * 5)
    for cell in ws[start_row + 1][3:3 + len(cu_list)]:
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            text_rotation=135)
        cell.font = Font(bold=True, size=9)
        cell.fill = PatternFill(
            start_color="D3D3D3",
            end_color="D3D3D3",
            fill_type="solid")

    headers = out_cols
    ws.append(headers)
    for cell in ws[start_row + 2]:
        cell.font = Font(bold=True, size=10, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(
            start_color="4A90E2",
            end_color="4A90E2",
            fill_type="solid")
        cell.border = Border(
            left=Side(
                style="thin"), right=Side(
                style="thin"), top=Side(
                style="thin"), bottom=Side(
                    style="thin"))

    for _, r in mastersheet.iterrows():
        rowvals = [r[col] for col in headers]
        ws.append(rowvals)

    # FIXED: Freeze the column headers (S/N, EXAM NUMBER, NAME, etc.) at row start_row + 3
    # This ensures all column headers remain visible when scrolling
    ws.freeze_panes = ws.cell(row=start_row + 3, column=1)

    thin_border = Border(
        left=Side(
            style="thin"), right=Side(
            style="thin"), top=Side(
                style="thin"), bottom=Side(
                    style="thin"))
    for row in ws.iter_rows(
            min_row=start_row + 3,
            max_row=ws.max_row,
            min_col=1,
            max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

    # Colorize course columns - SPECIAL COLOR FOR UPGRADED SCORES
    upgraded_fill = PatternFill(
        start_color="E6FFCC",
        end_color="E6FFCC",
        fill_type="solid")  # Light green for upgraded scores
    passed_fill = PatternFill(
        start_color="C6EFCE",
        end_color="C6EFCE",
        fill_type="solid")     # Normal green for passed
    failed_fill = PatternFill(
        start_color="FFFFFF",
        end_color="FFFFFF",
        fill_type="solid")     # White for failed

    for idx, code in enumerate(ordered_codes, start=4):
        col_letter = get_column_letter(idx)
        for r_idx in range(start_row + 3, ws.max_row + 1):
            cell = ws.cell(row=r_idx, column=idx)
            try:
                val = float(cell.value) if cell.value not in (None, "") else 0
                if upgrade_min_threshold is not None and upgrade_min_threshold <= val <= 49:
                    # This score was upgraded - use special color
                    cell.fill = upgraded_fill
                    # Dark green for upgraded scores
                    cell.font = Font(color="006600", bold=True)
                elif val >= pass_threshold:
                    cell.fill = passed_fill
                    cell.font = Font(color="006100")
                else:
                    cell.fill = failed_fill
                    cell.font = Font(color="FF0000", bold=True)
            except Exception:
                continue

    # Apply specific column alignments
    left_align_columns = ["CU Passed", "CU Failed", "TCPE", "GPA", "AVERAGE"]

    for col_idx, col_name in enumerate(headers, start=1):
        if col_name in left_align_columns:
            col_letter = get_column_letter(col_idx)
            for row_idx in range(
                    start_row + 3, ws.max_row + 1):  # Start from data rows
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(
                    horizontal="left", vertical="center")

        # Center align S/N column
        elif col_name == "S/N":
            col_letter = get_column_letter(col_idx)
            for row_idx in range(
                    start_row + 3, ws.max_row + 1):  # Start from data rows
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(
                    horizontal="center", vertical="center")

    # Calculate optimal column widths with special handling for REMARKS column
    longest_name_len = max([len(str(x)) for x in mastersheet["NAME"].fillna(
        "")]) if "NAME" in mastersheet.columns else 10
    name_col_width = min(max(longest_name_len + 2, 10), NAME_WIDTH_CAP)

    # Enhanced REMARKS column width calculation
    longest_remark_len = 0
    for remark in mastersheet["REMARKS"].fillna(""):
        remark_str = str(remark)
        # For "Failed" remarks, calculate the length considering all course
        # codes
        if remark_str.startswith("Failed:"):
            # Count the total total characters in failed courses list
            failed_courses = remark_str.replace("Failed: ", "")
            # Estimate the total total characters
            failed_length = len(failed_courses)
            # Add some padding for the "Failed: " prefix and spacing
            total_length = failed_length + 15
        else:
            total_length = len(remark_str)

        if total_length > longest_remark_len:
            longest_remark_len = total_length

    # Set REMARKS column width based on the longest content, with reasonable
    # limits
    # Expanded range for REMARKS
    remarks_col_width = min(max(longest_remark_len + 4, 40), 80)

    # Apply text wrapping and left alignment to REMARKS column
    remarks_col_idx = headers.index("REMARKS") + 1
    for row_idx in range(start_row + 3, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=remarks_col_idx)
        cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True)

    for col_idx, col in enumerate(ws.columns, start=1):
        column_letter = get_column_letter(col_idx)
        if col_idx == 1:  # S/N
            ws.column_dimensions[column_letter].width = 6
        elif column_letter == "B" or headers[col_idx - 1] in ["EXAM NUMBER", "EXAM NO"]:
            ws.column_dimensions[column_letter].width = 18
        elif headers[col_idx - 1] == "NAME":
            ws.column_dimensions[column_letter].width = name_col_width
        elif 4 <= col_idx < 4 + len(ordered_codes):  # course columns
            ws.column_dimensions[column_letter].width = 8
        elif headers[col_idx - 1] in ["REMARKS"]:
            # Dynamic width for REMARKS
            ws.column_dimensions[column_letter].width = remarks_col_width
        else:
            ws.column_dimensions[column_letter].width = 12

    # Fails per course row
    fails_per_course = mastersheet[ordered_codes].apply(
        lambda x: (x < pass_threshold).sum()).tolist()
    footer_vals = [""] * 2 + ["FAILS PER COURSE:"] + \
        fails_per_course + [""] * (len(headers) - 3 - len(ordered_codes))
    ws.append(footer_vals)
    for cell in ws[ws.max_row]:
        if 4 <= cell.column < 4 + len(ordered_codes):
            cell.fill = PatternFill(
                start_color="F0E68C",
                end_color="F0E68C",
                fill_type="solid")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        elif cell.column == 3:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

    # COMPREHENSIVE SUMMARY BLOCK
    total_students = len(mastersheet)
    passed_all = len(mastersheet[mastersheet["REMARKS"] == "Passed"])

    # Calculate students with GPA >= 2.0 but failed some courses
    gpa_above_2_failed = len(mastersheet[
        (mastersheet["GPA"] >= 2.0) &
        (mastersheet["REMARKS"] != "Passed") &
        (mastersheet["CU Passed"] >= 0.45 * total_cu)
    ])

    # Calculate students with GPA < 2.0 but passed at least 45% of credits
    gpa_below_2_failed = len(mastersheet[
        (mastersheet["GPA"] < 2.0) &
        (mastersheet["REMARKS"] != "Passed") &
        (mastersheet["CU Passed"] >= 0.45 * total_cu)
    ])

    # Calculate students who failed more than 45% of credit units
    failed_over_45_percent = len(mastersheet[
        (mastersheet["CU Failed"] > 0.45 * total_cu)
    ])

    # Add withdrawn student tracking to summary
    ws.append([])
    ws.append(["SUMMARY"])
    ws.append(
        [f"A total of {total_students} students registered and sat for the Examination"])
    ws.append(
        [f"A total of {passed_all} students passed in all courses registered and are to proceed to Second Semester, ND I"])
    ws.append([f"A total of {gpa_above_2_failed} students with Grade Point Average (GPA) of 2.00 and above failed various courses, but passed at least 45% of the total registered credit units, and are to carry these courses over to the next session."])
    ws.append([f"A total of {gpa_below_2_failed} students with Grade Point Average (GPA) below 2.00 failed various courses, but passed at least 45% of the total registered credit units, and are placed on Probation, to carry these courses over to the next session."])
    ws.append(
        [f"A total of {failed_over_45_percent} students failed in more than 45% of their registered credit units in various courses and have been advised to withdraw"])

    # FIXED: Keep the upgrade notice only in the summary section, not in the header
    if upgrade_min_threshold is not None:
        ws.append(
            [f"✅ Upgraded all scores between {upgrade_min_threshold}–49 to 50 as per management decision ({upgraded_scores_count} scores upgraded)"])

    # Add removed withdrawn students info
    if removed_students:
        ws.append(
            [f"NOTE: {len(removed_students)} previously withdrawn students were removed from this semester's results as they should not be processed."])

    ws.append(["The above decisions are in line with the provisions of the General Information Section of the NMCN/NBTE Examinations Regulations (Pg 4) adopted by the College."])
    ws.append([])
    ws.append(["________________________", "", "",
              "________________________", "", "", "", "", "", "", "", "", ""])
    ws.append(["Mrs. Abini Hauwa", "", "", "Mrs. Olukemi Ogunleye",
              "", "", "", "", "", "", "", "", ""])
    ws.append(["Head of Exams",
               "",
               "",
               "Chairman, ND/HND Program C'tee",
               "",
               "",
               "",
               "",
               "",
               "",
               "",
               "",
               ""])

    wb.save(out_xlsx)
    print(f"✅ Mastersheet saved: {out_xlsx}")

    # Generate individual student PDF with previous GPAs and CGPA
    safe_sem = re.sub(r'[^\w\-]', '_', sem)
    student_pdf_path = os.path.join(
        output_dir,
        f"mastersheet_students_{ts}_{safe_sem}.pdf")

    print(f"📊 FINAL CHECK before PDF generation:")
    print(f"   Previous GPAs loaded: {len(previous_gpas)}")
    print(
        f"   CGPA data available for: {len(cgpa_data) if cgpa_data else 0} students")
    if previous_gpas:
        sample = list(previous_gpas.items())[:3]
        print(f"   Sample GPAs: {sample}")

    try:
        generate_individual_student_pdf(
            mastersheet,
            student_pdf_path,
            sem,
            logo_path=logo_path_norm,
            prev_mastersheet_df=None,
            filtered_credit_units=filtered_credit_units,
            ordered_codes=ordered_codes,
            course_titles_map=course_titles,
            previous_gpas=previous_gpas,
            cgpa_data=cgpa_data,
            total_cu=total_cu,
            pass_threshold=pass_threshold,
            upgrade_min_threshold=upgrade_min_threshold)  # PASS THE UPGRADE THRESHOLD TO PDF
        print(f"✅ PDF generated successfully for {sem}")
    except Exception as e:
        print(f"❌ Failed to generate student PDF for {sem}: {e}")
        import traceback
        traceback.print_exc()

    return mastersheet

def process_semester_files(
    semester_key,
    raw_files,
    raw_dir,
    output_dir,
    ts,
    pass_threshold,
    semester_course_maps,
    semester_credit_units,
    semester_lookup,
    semester_course_titles,
    logo_path,
    set_name,
    previous_gpas=None,
    upgrade_min_threshold=None):
    """
    Process all files for a specific semester with carryover integration.
    """
    print(f"\n{'='*60}")
    print(f"PROCESSING SEMESTER: {semester_key}")
    print(f"{'='*60}")

    # Filter files for this semester
    normalized_key = semester_key.replace('ND-', '').upper()
    semester_files = [f for f in raw_files 
                      if normalized_key in f.upper().replace('ND-', '')]
    
    if not semester_files:
        print(f"⚠️ No files found for semester {semester_key}")
        print(f"🔍 Available files: {raw_files}")
        return None

    print(
        f"📁 Found {len(semester_files)} files for {semester_key}: {semester_files}")

    # Check for existing carryover files
    existing_carryover_files = check_existing_carryover_files(raw_dir, set_name, semester_key)
    if existing_carryover_files:
        print(f"📋 Found existing carryover files: {existing_carryover_files}")
        print("ℹ️ Carryover processing will be available after regular processing")

    # Process each file for this semester
    mastersheet_result = None
    for rf in semester_files:
        raw_path = os.path.join(raw_dir, rf)
        print(f"\n📄 Processing: {rf}")

        try:
            # Load previous GPAs for this specific semester
            current_previous_gpas = load_previous_gpas_from_processed_files(
                output_dir, semester_key, ts) if previous_gpas is None else previous_gpas

            # Load CGPA data (all previous semesters)
            cgpa_data = load_all_previous_gpas_for_cgpa(
                output_dir, semester_key, ts)

            # Process the file
            result = process_single_file(
                raw_path,
                output_dir,
                ts,
                pass_threshold,
                semester_course_maps,
                semester_credit_units,
                semester_lookup,
                semester_course_titles,
                logo_path,
                semester_key,
                set_name,
                current_previous_gpas,
                cgpa_data,
                upgrade_min_threshold)

            if result is not None:
                print(f"✅ Successfully processed {rf}")
                mastersheet_result = result
            else:
                print(f"❌ Failed to process {rf}")

        except Exception as e:
            print(f"❌ Error processing {rf}: {e}")
            import traceback
            traceback.print_exc()

    # ADD: Verify carryover records were created
    carryover_records_dir = os.path.join(output_dir, "CARRYOVER_RECORDS")
    if os.path.exists(carryover_records_dir):
        json_files = glob.glob(os.path.join(carryover_records_dir, f"co_student_{set_name}_{semester_key}_*.json"))
        if json_files:
            print(f"✅ Carryover records created: {len(json_files)} file(s)")
            print(f"📝 Latest: {sorted(json_files)[-1]}")
        else:
            print(f"⚠️ No carryover records found for {semester_key}")
    
    return mastersheet_result

# ----------------------------
# Main runner (Enhanced)
# ----------------------------

def main():
    print("Starting ND Examination Results Processing with Data Transformation...")
    ts = datetime.now().strftime(TIMESTAMP_FMT)

    # Initialize trackers
    initialize_student_tracker()
    initialize_carryover_tracker()

    # Check if running in web mode
    if is_web_mode():
        uploaded_file_path = get_uploaded_file_path()
        if uploaded_file_path and os.path.exists(uploaded_file_path):
            print("🔧 Running in WEB MODE with uploaded file")
            # This would need to be adapted for your specific uploaded file processing
            print("⚠️ Uploaded file processing for individual files not fully implemented in this version")
            return

    # Get parameters from form
    params = get_form_parameters()
    
    # Use the parameters
    global DEFAULT_PASS_THRESHOLD
    DEFAULT_PASS_THRESHOLD = params['pass_threshold']
    
    base_dir_norm = normalize_path(BASE_DIR)
    print(f"Using base directory: {base_dir_norm}")

    # Check if we should use interactive or non-interactive mode
    if should_use_interactive_mode():
        print("🔧 Running in INTERACTIVE mode (CLI)")
        
        try:
            semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_course_data()
        except Exception as e:
            print(f"❌ Could not load course data: {e}")
            return

        # Get available sets and let user choose
        available_sets = get_available_sets(base_dir_norm)

        if not available_sets:
            print(
                f"No ND-* directories found in {base_dir_norm}. Nothing to process.")
            print(f"Available directories: {os.listdir(base_dir_norm)}")
            return

        print(f"📚 Found {len(available_sets)} available sets: {available_sets}")

        # Let user choose which set(s) to process
        sets_to_process = get_user_set_choice(available_sets)

        print(f"\n🎯 PROCESSING SELECTED SETS: {sets_to_process}")

        for nd_set in sets_to_process:
            print(f"\n{'='*60}")
            print(f"PROCESSING SET: {nd_set}")
            print(f"{'='*60}")

            # Generate a single timestamp for this set processing
            ts = datetime.now().strftime(TIMESTAMP_FMT)
            
            # UPDATED: Raw and clean directories now under ND folder
            raw_dir = normalize_path(os.path.join(base_dir_norm, "ND", nd_set, "RAW_RESULTS"))
            clean_dir = normalize_path(os.path.join(base_dir_norm, "ND", nd_set, "CLEAN_RESULTS"))

            # Create directories if they don't exist
            os.makedirs(raw_dir, exist_ok=True)
            os.makedirs(clean_dir, exist_ok=True)

            # Check if raw directory exists and has files
            if not os.path.exists(raw_dir):
                print(f"⚠️ RAW_RESULTS directory not found: {raw_dir}")
                continue

            raw_files = [
                f for f in os.listdir(raw_dir) if f.lower().endswith(
                    (".xlsx", ".xls")) and not f.startswith("~$")]
            if not raw_files:
                print(f"⚠️ No raw files in {raw_dir}; skipping {nd_set}")
                print(f"   Available files: {os.listdir(raw_dir)}")
                continue

            print(f"📁 Found {len(raw_files)} raw files in {nd_set}: {raw_files}")

            # Create a single timestamped folder for this set
            set_output_dir = os.path.join(clean_dir, f"{nd_set}_RESULT-{ts}")
            os.makedirs(set_output_dir, exist_ok=True)
            print(f"📁 Created set output directory: {set_output_dir}")

            # Get user choice for which semesters to process
            semesters_to_process = get_user_semester_choice()

            print(
                f"\n🎯 PROCESSING SELECTED SEMESTERS for {nd_set}: {[get_semester_display_info(sem)[3] for sem in semesters_to_process]}")

            # Process selected semesters in the correct order
            for semester_key in semesters_to_process:
                if semester_key not in SEMESTER_ORDER:
                    print(f"⚠️ Skipping unknown semester: {semester_key}")
                    continue

                # Check if there are files for this semester
                semester_files_exist = False
                for rf in raw_files:
                    detected_sem, _, _, _, _, _ = detect_semester_from_filename(rf)
                    if detected_sem == semester_key:
                        semester_files_exist = True
                        break

                if semester_files_exist:
                    print(f"\n🎯 Processing {semester_key} in {nd_set}...")
                    process_semester_files(
                        semester_key,
                        raw_files,
                        raw_dir,
                        set_output_dir,
                        ts,
                        DEFAULT_PASS_THRESHOLD,
                        semester_course_maps,
                        semester_credit_units,
                        semester_lookup,
                        semester_course_titles,
                        DEFAULT_LOGO_PATH,
                        nd_set)
                else:
                    print(
                        f"⚠️ No files found for {semester_key} in {nd_set}, skipping...")
            
            # Create ZIP of the entire set results
            try:
                zip_path = os.path.join(clean_dir, f"{nd_set}_RESULT-{ts}.zip")
                zip_success = create_zip_folder(set_output_dir, zip_path)
                
                if zip_success:
                    print(f"✅ ZIP file created: {zip_path}")
                    
                    # Verify file size
                    if os.path.exists(zip_path):
                        zip_size = os.path.getsize(zip_path)
                        zip_size_mb = zip_size / (1024 * 1024)
                        print(f"📦 ZIP file size: {zip_size_mb:.2f} MB")
                    
                else:
                    print(f"❌ Failed to create ZIP file for {nd_set}")
                    
            except Exception as e:
                print(f"⚠️ Failed to create ZIP for {nd_set}: {e}")

        # Print student tracking summary
        print(f"\n📊 STUDENT TRACKING SUMMARY:")
        print(f"Total unique students tracked: {len(STUDENT_TRACKER)}")
        print(f"Total withdrawn students: {len(WITHDRAWN_STUDENTS)}")

        # Print carryover summary
        if CARRYOVER_STUDENTS:
            print(f"\n📋 CARRYOVER STUDENT SUMMARY:")
            print(f"Total carryover students: {len(CARRYOVER_STUDENTS)}")
            
            # Count by semester
            semester_counts = {}
            for student_key, data in CARRYOVER_STUDENTS.items():
                semester = data['semester']
                semester_counts[semester] = semester_counts.get(semester, 0) + 1
            
            for semester, count in semester_counts.items():
                print(f"  {semester}: {count} students")

        # Print withdrawn students who reappeared
        reappeared_count = 0
        for exam_no, data in WITHDRAWN_STUDENTS.items():
            if data['reappeared_semesters']:
                reappeared_count += 1
                print(
                    f"🚨 {exam_no}: Withdrawn in {data['withdrawn_semester']}, reappeared in {data['reappeared_semesters']}")

        if reappeared_count > 0:
            print(
                f"🚨 ALERT: {reappeared_count} previously withdrawn students have reappeared in later semesters!")

        # Analyze student progression
        sem_counts = {}
        for student_data in STUDENT_TRACKER.values():
            sem_count = len(student_data['semesters_present'])
            if sem_count not in sem_counts:
                sem_counts[sem_count] = 0
            sem_counts[sem_count] += 1

        for sem_count, student_count in sorted(sem_counts.items()):
            print(f"Students present in {sem_count} semester(s): {student_count}")

        print("\n✅ ND Examination Results Processing completed successfully.")
    else:
        print("🔧 Running in NON-INTERACTIVE mode (Web)")
        
        # NEW: Check if this is carryover processing mode
        if params.get('process_carryover', False):
            print("🎯 Detected CARRYOVER processing mode - redirecting to integrated_carryover_processor.py")
            
            # Set environment variables for the carryover processor
            os.environ['CARRYOVER_FILE_PATH'] = params['carryover_file_path']
            os.environ['SET_NAME'] = params['selected_set']
            os.environ['SEMESTER_KEY'] = params['selected_semesters'][0] if params['selected_semesters'] else ''
            os.environ['BASE_DIR'] = BASE_DIR

            # Path to the integrated_carryover_processor.py script
            carryover_script_path = os.path.join(os.path.dirname(__file__), 'integrated_carryover_processor.py')
            
            if not os.path.exists(carryover_script_path):
                print(f"❌ Carryover processor script not found: {carryover_script_path}")
                return False

            print(f"🚀 Running carryover processor: {carryover_script_path}")
            
            # Run the carryover processor
            result = subprocess.run([sys.executable, carryover_script_path], capture_output=True, text=True)

            # Print the output and return
            print(result.stdout)
            if result.stderr:
                print(result.stderr)

            return result.returncode == 0
        else:
            # Regular processing mode
            success = process_in_non_interactive_mode(params, base_dir_norm)
            if success:
                print("✅ ND Examination Results Processing completed successfully")
            else:
                print("❌ ND Examination Results Processing failed")
        return

def process_in_non_interactive_mode(params, base_dir_norm):
    """Process exams in non-interactive mode for web interface."""
    print("🔧 Running in NON-INTERACTIVE mode (web interface)")
    
    # Use parameters from environment variables
    selected_set = params['selected_set']
    selected_semesters = params['selected_semesters']
    
    # FIX: Normalize semester names to uppercase for consistent matching
    selected_semesters = [sem.upper() for sem in selected_semesters]
    print(f"🎯 Processing semesters (normalized): {selected_semesters}")
    
    # Get upgrade threshold from environment variable if provided
    upgrade_min_threshold = get_upgrade_threshold_from_env()
    
    # Get available sets
    available_sets = get_available_sets(base_dir_norm)
    
    if not available_sets:
        print("❌ No ND sets found")
        return False
    
    # Remove ND-COURSES from available sets if present
    available_sets = [s for s in available_sets if s != 'ND-COURSES']
    
    if not available_sets:
        print("❌ No valid ND sets found (only ND-COURSES present)")
        return False
    
    # Determine which sets to process
    if selected_set == "all":
        sets_to_process = available_sets
        print(f"🎯 Processing ALL sets: {sets_to_process}")
    else:
        if selected_set in available_sets:
            sets_to_process = [selected_set]
            print(f"🎯 Processing selected set: {selected_set}")
        else:
            print(f"⚠️ Selected set '{selected_set}' not found, processing all sets")
            sets_to_process = available_sets
    
    # Load course data once
    try:
        semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_course_data()
        print(f"✅ Loaded course data for semesters: {list(semester_course_maps.keys())}")
    except Exception as e:
        print(f"❌ Could not load course data: {e}")
        return False
    
    # Initialize carryover tracker
    initialize_carryover_tracker()
    
    # Process each set and semester
    total_processed = 0
    for nd_set in sets_to_process:
        print(f"\n{'='*60}")
        print(f"PROCESSING SET: {nd_set}")
        print(f"{'='*60}")
        
        # Generate a single timestamp for this set processing
        ts = datetime.now().strftime(TIMESTAMP_FMT)
        
        # UPDATED: Raw and clean directories now under ND folder
        raw_dir = normalize_path(os.path.join(base_dir_norm, "ND", nd_set, "RAW_RESULTS"))
        clean_dir = normalize_path(os.path.join(base_dir_norm, "ND", nd_set, "CLEAN_RESULTS"))
        
        # Create directories if they don't exist
        os.makedirs(raw_dir, exist_ok=True)
        os.makedirs(clean_dir, exist_ok=True)
        
        if not os.path.exists(raw_dir):
            print(f"⚠️ RAW_RESULTS directory not found: {raw_dir}")
            continue
        
        raw_files = [f for f in os.listdir(raw_dir) if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
        if not raw_files:
            print(f"⚠️ No raw files in {raw_dir}; skipping {nd_set}")
            continue
        
        print(f"📁 Found {len(raw_files)} raw files in {nd_set}: {raw_files}")
        
        # Create a single timestamped folder for this set
        set_output_dir = os.path.join(clean_dir, f"{nd_set}_RESULT-{ts}")
        os.makedirs(set_output_dir, exist_ok=True)
        print(f"📁 Created set output directory: {set_output_dir}")
        
        # Process selected semesters - FIXED: Use normalized (uppercase) semester names
        for semester_key in selected_semesters:
            # FIX: Check if semester exists in course data (case-sensitive)
            if semester_key not in semester_course_maps:
                print(f"⚠️ Semester '{semester_key}' not found in course data. Available: {list(semester_course_maps.keys())}")
                continue
            
            # Check if there are files for this semester
            semester_files_exist = False
            for rf in raw_files:
                detected_sem, _, _, _, _, _ = detect_semester_from_filename(rf)
                # FIX: Compare in uppercase for case-insensitive matching
                if detected_sem.upper() == semester_key.upper():
                    semester_files_exist = True
                    break
            
            if semester_files_exist:
                print(f"\n🎯 Processing {semester_key} in {nd_set}...")
                try:
                    # Process the semester with the upgrade threshold
                    result = process_semester_files(
                        semester_key,
                        raw_files,
                        raw_dir,
                        set_output_dir,
                        ts,
                        params['pass_threshold'],
                        semester_course_maps,
                        semester_credit_units,
                        semester_lookup,
                        semester_course_titles,
                        DEFAULT_LOGO_PATH,
                        nd_set,
                        previous_gpas=None,
                        upgrade_min_threshold=upgrade_min_threshold
                    )
                    
                    if result is not None:
                        print(f"✅ Successfully processed {semester_key}")
                        total_processed += 1
                    else:
                        print(f"❌ Failed to process {semester_key}")
                        
                except Exception as e:
                    print(f"❌ Error processing {semester_key}: {e}")
                    import traceback
                    traceback.print_exc()
            else:
                print(f"⚠️ No files found for {semester_key} in {nd_set}, skipping...")
        
        # Create ZIP of the entire set results
        try:
            zip_path = os.path.join(clean_dir, f"{nd_set}_RESULT-{ts}.zip")
            zip_success = create_zip_folder(set_output_dir, zip_path)
            
            if zip_success:
                # Verify the ZIP file was created and has content
                if os.path.exists(zip_path):
                    zip_size = os.path.getsize(zip_path)
                    print(f"✅ ZIP file created: {zip_path} ({zip_size} bytes)")
                    
                    # Convert bytes to MB for readability
                    zip_size_mb = zip_size / (1024 * 1024)
                    print(f"📦 ZIP file size: {zip_size_mb:.2f} MB")
                else:
                    print(f"❌ ZIP file was not created: {zip_path}")
            else:
                print(f"❌ Failed to create ZIP file for {nd_set}")
                
        except Exception as e:
            print(f"⚠️ Failed to create ZIP for {nd_set}: {e}")
            import traceback
            traceback.print_exc()
    
    print(f"\n📊 PROCESSING SUMMARY: {total_processed} semester(s) processed")
    
    # Print carryover summary
    if CARRYOVER_STUDENTS:
        print(f"\n📋 CARRYOVER SUMMARY:")
        print(f"   Total carryover students: {len(CARRYOVER_STUDENTS)}")
        
        # Count by semester
        semester_counts = {}
        for student_key, data in CARRYOVER_STUDENTS.items():
            semester = data['semester']
            semester_counts[semester] = semester_counts.get(semester, 0) + 1
        
        for semester, count in semester_counts.items():
            print(f"   {semester}: {count} students")
    
    return total_processed > 0

if __name__ == "__main__":
    try:
        main()
        print("✅ ND Examination Results Processing completed successfully")
    except Exception as e:
        print(f"❌ Error during processing: {e}")
        import traceback
        traceback.print_exc()