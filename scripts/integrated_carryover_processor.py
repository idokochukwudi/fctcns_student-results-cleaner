#!/usr/bin/env python3
"""
integrated_carryover_processor.py
Script to process carryover/resit results, update mastersheet, and generate reports.
ENHANCED DEBUGGING VERSION - Enhanced debugging and matching logic for carryover processing
"""

from openpyxl.cell import MergedCell
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
import glob
import json
import traceback
import shutil
import tempfile
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ----------------------------
# Configuration
# ----------------------------

def get_base_directory():
    """Get base directory - compatible with both local and Railway environments"""
    base_dir_env = os.getenv('BASE_DIR')
    if base_dir_env and os.path.exists(base_dir_env):
        print(f"✅ Using BASE_DIR from environment: {base_dir_env}")
        return base_dir_env
    
    # Check if we're running on Railway
    def is_running_on_railway():
        return any(key in os.environ for key in [
            'RAILWAY_ENVIRONMENT', 'RAILWAY_STATIC_URL', 
            'RAILWAY_PROJECT_ID', 'RAILWAY_SERVICE_NAME'
        ])
    
    if is_running_on_railway():
        # Create the directory structure on Railway
        railway_base = '/app/EXAMS_INTERNAL'
        os.makedirs(railway_base, exist_ok=True)
        os.makedirs(os.path.join(railway_base, 'ND', 'ND-COURSES'), exist_ok=True)
        print(f"✅ Using Railway base directory: {railway_base}")
        return railway_base
    
    # Local development fallbacks - UPDATED PATHS
    local_paths = [
        os.path.join(os.path.expanduser('~'), 'student_result_cleaner'),
        os.path.dirname(__file__),
        os.getcwd()
    ]
    
    for local_path in local_paths:
        exam_internal_path = os.path.join(local_path, 'EXAMS_INTERNAL')
        if os.path.exists(exam_internal_path):
            print(f"✅ Using base directory: {exam_internal_path}")
            return exam_internal_path
    
    # Final fallback - create in current directory
    fallback_path = os.path.join(os.getcwd(), 'EXAMS_INTERNAL')
    os.makedirs(fallback_path, exist_ok=True)
    os.makedirs(os.path.join(fallback_path, 'ND', 'ND-COURSES'), exist_ok=True)
    print(f"✅ Created fallback base directory: {fallback_path}")
    return fallback_path

BASE_DIR = get_base_directory()
ND_COURSES_DIR = os.path.join(BASE_DIR, "ND", "ND-COURSES")

# FIX 2: Align timestamp format with app.py (day-month-year)
TIMESTAMP_FMT = "%d-%m-%Y_%H%M%S"  # Changed from "%Y%m%d_%H%M%S"

# FIX: Ensure consistent uppercase semester naming to match main processor
SEMESTER_ORDER = [
    "ND-FIRST-YEAR-FIRST-SEMESTER",
    "ND-FIRST-YEAR-SECOND-SEMESTER", 
    "ND-SECOND-YEAR-FIRST-SEMESTER",
    "ND-SECOND-YEAR-SECOND-SEMESTER"
]
DEFAULT_PASS_THRESHOLD = 50.0
DEFAULT_LOGO_PATH = os.path.join(os.path.dirname(__file__), "logo.png")
CARRYOVER_STUDENTS = {}
STUDENT_TRACKER = {}
WITHDRAWN_STUDENTS = {}

def is_running_on_railway():
    """Check if running on Railway."""
    return any(key in os.environ for key in [
        'RAILWAY_ENVIRONMENT', 
        'RAILWAY_STATIC_URL', 
        'RAILWAY_PROJECT_ID',
        'RAILWAY_SERVICE_NAME'
    ])

def is_web_mode():
    """Check if running in web mode (file upload)."""
    return os.getenv('WEB_MODE') == 'true'

def get_uploaded_file_path():
    """Get path of uploaded file in web mode."""
    return os.getenv('UPLOADED_FILE_PATH')

def get_form_parameters():
    """Get parameters from environment variables."""
    selected_semesters = os.getenv('SELECTED_SEMESTERS', '').split(',') if os.getenv('SELECTED_SEMESTERS') else []
    selected_semesters = [s.strip().upper() for s in selected_semesters if s.strip().upper() in SEMESTER_ORDER]
    return {
        'selected_set': os.getenv('SELECTED_SET', 'all'),
        'selected_semesters': selected_semesters,
        'pass_threshold': float(os.getenv('PASS_THRESHOLD', '50.0')),
        'generate_pdf': os.getenv('GENERATE_PDF', 'True').lower() == 'true',
        'track_withdrawn': os.getenv('TRACK_WITHDRAWN', 'True').lower() == 'true',
        'process_resit': os.getenv('PROCESS_RESIT', 'False').lower() == 'true',
        'resit_file_path': os.getenv('RESIT_FILE_PATH', ''),
        'resit_output_dir': os.path.join(BASE_DIR, "ND", "RESIT_OUTPUT")
    }

# ----------------------------
# Enhanced Debugging Functions (NEW)
# ----------------------------

def debug_carryover_matching(resit_file_path, mastersheet_path, semester_key, set_name, pass_threshold, course_map, output_dir):
    """Enhanced debugging to identify why no scores are being updated."""
    print(f"\n🔍 DEEP DEBUGGING FOR CARRYOVER MISMATCH")
    print("=" * 60)
    
    try:
        # Load files
        resit_df = pd.read_excel(resit_file_path, header=0)
        mastersheet_df = pd.read_excel(mastersheet_path, sheet_name=semester_key, header=5)
        
        # Find exam columns
        resit_exam_col = find_exam_number_column(resit_df)
        mastersheet_exam_col = find_exam_number_column(mastersheet_df) or 'EXAM NUMBER'
        
        # Normalize course map
        course_map_normalized = {}
        if semester_key in course_map:
            course_map_normalized = {normalize_course_name(k): v for k, v in course_map[semester_key].items()}
        
        # Get carryover students for this semester
        carryover_in_semester = {k: v for k, v in CARRYOVER_STUDENTS.items() if semester_key in k}
        print(f"📋 Carryover students in {semester_key}: {len(carryover_in_semester)}")
        
        debug_log = []
        debug_log.append("DEEP DEBUGGING ANALYSIS")
        debug_log.append("=" * 50)
        debug_log.append(f"Semester: {semester_key}")
        debug_log.append(f"Pass threshold: {pass_threshold}")
        debug_log.append(f"Carryover students: {len(carryover_in_semester)}")
        debug_log.append("")
        
        # Analyze each carryover student
        for student_key, carryover_data in carryover_in_semester.items():
            exam_no = carryover_data['exam_number']
            debug_log.append(f"STUDENT: {exam_no} - {carryover_data['name']}")
            debug_log.append(f"Failed courses: {[c['course_code'] for c in carryover_data['failed_courses']]}")
            
            # Find student in resit file
            resit_student_mask = resit_df[resit_exam_col].astype(str).str.strip().str.upper() == exam_no.upper()
            if not resit_student_mask.any():
                debug_log.append(f"  ❌ Not found in resit file")
                debug_log.append("")
                continue
            
            # Find student in mastersheet
            mastersheet_student_mask = mastersheet_df[mastersheet_exam_col].astype(str).str.strip().str.upper() == exam_no.upper()
            if not mastersheet_student_mask.any():
                debug_log.append(f"  ❌ Not found in mastersheet")
                debug_log.append("")
                continue
            
            resit_student = resit_df[resit_student_mask].iloc[0]
            mastersheet_student = mastersheet_df[mastersheet_student_mask].iloc[0]
            
            debug_log.append(f"  ✅ Found in both files")
            
            # Check each failed course
            for failed_course in carryover_data['failed_courses']:
                course_code = failed_course['course_code']
                original_score = failed_course['original_score']
                
                debug_log.append(f"  📊 Failed course: {course_code} (Original: {original_score})")
                
                # Check if course exists in mastersheet
                if course_code not in mastersheet_df.columns:
                    debug_log.append(f"    ❌ Course code not in mastersheet columns")
                    continue
                
                current_mastersheet_score = mastersheet_student.get(course_code)
                debug_log.append(f"    Current mastersheet score: {current_mastersheet_score}")
                
                # Try to find matching resit course
                resit_course_found = False
                for resit_col in resit_df.columns:
                    if resit_col == resit_exam_col or resit_col == 'NAME' or 'Unnamed' in str(resit_col):
                        continue
                    
                    # Try to match this resit column to the failed course
                    course_info = find_best_course_match(resit_col, course_map_normalized)
                    if course_info and course_info['code'] == course_code:
                        resit_score = resit_student.get(resit_col)
                        debug_log.append(f"    🔍 Found resit column: '{resit_col}' -> Score: {resit_score}")
                        
                        if pd.isna(resit_score) or resit_score == '':
                            debug_log.append(f"      ❌ Resit score is empty/NaN")
                        else:
                            try:
                                resit_score_val = float(resit_score)
                                debug_log.append(f"      📈 Resit score: {resit_score_val}")
                                
                                # Check update conditions
                                condition1 = original_score < pass_threshold
                                condition2 = resit_score_val >= pass_threshold
                                condition3 = current_mastersheet_score < pass_threshold
                                
                                debug_log.append(f"      📋 Conditions:")
                                debug_log.append(f"        Original ({original_score}) < {pass_threshold}: {condition1}")
                                debug_log.append(f"        Resit ({resit_score_val}) >= {pass_threshold}: {condition2}")
                                debug_log.append(f"        Current ({current_mastersheet_score}) < {pass_threshold}: {condition3}")
                                
                                if condition1 and condition2 and condition3:
                                    debug_log.append(f"      ✅ WOULD UPDATE: {original_score} → {resit_score_val}")
                                else:
                                    debug_log.append(f"      ❌ NO UPDATE - Conditions not met")
                                    
                            except (ValueError, TypeError) as e:
                                debug_log.append(f"      ❌ Invalid resit score: {resit_score} - {e}")
                        
                        resit_course_found = True
                        break
                
                if not resit_course_found:
                    debug_log.append(f"    ❌ No matching resit column found for {course_code}")
                    # Show available resit columns for debugging
                    resit_cols = [col for col in resit_df.columns if col != resit_exam_col and col != 'NAME' and 'Unnamed' not in str(col)]
                    debug_log.append(f"    Available resit columns: {resit_cols[:10]}...")  # First 10 only
            
            debug_log.append("")
        
        # Save debug log
        debug_file = os.path.join(output_dir, f"deep_debug_carryover_{set_name}_{semester_key}_{datetime.now().strftime(TIMESTAMP_FMT)}.txt")
        with open(debug_file, 'w') as f:
            f.write("\n".join(debug_log))
        
        print(f"📝 Deep debug analysis saved: {debug_file}")
        print(f"🔍 Check this file to understand why no updates are happening")
        
        return debug_file
        
    except Exception as e:
        print(f"❌ Error in deep debugging: {e}")
        traceback.print_exc()
        return None

def analyze_resit_file_content(resit_file_path, semester_key):
    """Analyze what's actually in the resit file."""
    print(f"\n📊 RESIT FILE CONTENT ANALYSIS")
    print("=" * 40)
    
    try:
        resit_df = pd.read_excel(resit_file_path, header=0)
        exam_col = find_exam_number_column(resit_df)
        
        print(f"File: {resit_file_path}")
        print(f"Semester: {semester_key}")
        print(f"Shape: {resit_df.shape}")
        print(f"Columns: {list(resit_df.columns)}")
        print(f"Number of students: {len(resit_df)}")
        
        # Show first few students and their courses
        print(f"\n📋 FIRST 3 STUDENTS IN RESIT FILE:")
        for i in range(min(3, len(resit_df))):
            student = resit_df.iloc[i]
            exam_no = student[exam_col] if exam_col in student else "N/A"
            print(f"  {i+1}. {exam_no}")
            
            # Show courses with scores
            course_cols = [col for col in resit_df.columns if col != exam_col and col != 'NAME' and 'Unnamed' not in str(col)]
            for course in course_cols[:5]:  # First 5 courses only
                score = student.get(course)
                if pd.notna(score) and score != '':
                    print(f"     {course}: {score}")
        
        return True
    except Exception as e:
        print(f"❌ Error analyzing resit file: {e}")
        return False

# ----------------------------
# Helper Functions
# ----------------------------

def normalize_path(path):
    """Normalize file path for cross-platform compatibility."""
    return os.path.normpath(path)

def normalize_for_matching(text):
    """Normalize text for matching (for semester names and filenames)."""
    if not isinstance(text, str):
        return ""
    # Convert to lowercase and strip
    normalized = text.lower().strip()
    # Remove punctuation and extra spaces
    normalized = re.sub(r'[^\w\s]', '', normalized)
    normalized = re.sub(r'\s+', '-', normalized)
    return normalized

def get_grade(score):
    """Determine grade based on score."""
    try:
        score = float(score)
        if score >= 70: return "A"
        elif score >= 60: return "B"
        elif score >= 50: return "C"
        elif score >= 45: return "D"
        elif score >= 40: return "E"
        else: return "F"
    except (ValueError, TypeError):
        return "F"

def get_grade_point(score):
    """Determine grade point based on score."""
    try:
        score = float(score)
        if score >= 70: return 4.0
        elif score >= 60: return 3.0
        elif score >= 50: return 2.0
        elif score >= 45: return 1.0
        elif score >= 40: return 0.5
        else: return 0.0
    except (ValueError, TypeError):
        return 0.0

def get_withdrawal_history(exam_no):
    """Retrieve withdrawal history for a student."""
    return WITHDRAWN_STUDENTS.get(exam_no)

def initialize_student_tracker():
    """Initialize the global student tracker."""
    global STUDENT_TRACKER
    STUDENT_TRACKER = {}

def initialize_carryover_tracker():
    """Initialize the global carryover tracker."""
    global CARRYOVER_STUDENTS
    CARRYOVER_STUDENTS = {}

def get_available_sets(base_dir):
    """Get available ND sets from the base directory."""
    nd_dir = os.path.join(base_dir, "ND")
    if not os.path.exists(nd_dir):
        return []
    return [d for d in os.listdir(nd_dir) if os.path.isdir(os.path.join(nd_dir, d)) and d.startswith("ND-")]

def get_user_set_choice(available_sets):
    """Prompt user to choose sets to process."""
    print("\n📚 AVAILABLE SETS:")
    for i, set_name in enumerate(available_sets, 1):
        print(f"{i}. {set_name}")
    print(f"{len(available_sets) + 1}. Process all sets")

    while True:
        try:
            choice = input("\nEnter set number(s) separated by commas (or 'all' for all sets): ").strip()
            if choice.lower() == 'all' or choice == str(len(available_sets) + 1):
                return available_sets
            choices = [c.strip() for c in choice.split(',')]
            valid_choices = []
            for c in choices:
                if c.isdigit() and 1 <= int(c) <= len(available_sets):
                    valid_choices.append(available_sets[int(c) - 1])
                else:
                    print(f"❌ Invalid choice: {c}")
            if valid_choices:
                return valid_choices
            print("❌ No valid sets selected. Please try again.")
        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)

def get_user_semester_choice():
    """Prompt user to choose semesters to process."""
    print("\n🎯 SEMESTER PROCESSING OPTIONS:")
    for i, sem in enumerate(SEMESTER_ORDER, 1):
        year, sem_num, level, sem_display, set_code = get_semester_display_info(sem)
        print(f"{i}. {level} - {sem_display}")
    print(f"{len(SEMESTER_ORDER) + 1}. Process all semesters")

    while True:
        try:
            choice = input("\nEnter semester number(s) separated by commas: ").strip()
            if choice == str(len(SEMESTER_ORDER) + 1):
                return SEMESTER_ORDER
            choices = [c.strip() for c in choice.split(',')]
            valid_choices = []
            for c in choices:
                if c.isdigit() and 1 <= int(c) <= len(SEMESTER_ORDER):
                    valid_choices.append(SEMESTER_ORDER[int(c) - 1])
                else:
                    print(f"❌ Invalid choice: {c}")
            if valid_choices:
                return valid_choices
            print("❌ No valid semesters selected. Please try again.")
        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)

def create_zip_folder(source_dir, zip_path):
    """Create a ZIP file from a directory."""
    try:
        temp_dir = tempfile.mkdtemp()
        temp_zip = os.path.join(temp_dir, os.path.basename(zip_path))
        shutil.make_archive(temp_zip.replace('.zip', ''), 'zip', source_dir)
        shutil.move(temp_zip, zip_path)
        shutil.rmtree(temp_dir)
        return True
    except Exception as e:
        print(f"❌ Error creating ZIP: {e}")
        return False

# ----------------------------
# DATA TRANSFORMATION FUNCTIONS
# ----------------------------

def transform_transposed_data(df, sheet_type):
    """Transform transposed data format to wide format."""
    print(f"🔄 Transforming {sheet_type} sheet from transposed to wide format...")
    
    reg_col = find_column_by_names(df, ["REG. no", "Reg No", "Registration Number", "Exam No"])
    name_col = find_column_by_names(df, ["NAME", "Full Name", "Candidate Name"])
    
    if not reg_col:
        print("❌ Could not find registration column for transformation")
        return df
    
    course_columns = [col for col in df.columns 
                     if col not in [reg_col, name_col] and col not in ['', None] and not str(col).startswith('Unnamed')]
    
    print(f"📊 Found {len(course_columns)} course columns: {course_columns}")
    
    transformed_data = []
    student_dict = {}
    
    for idx, row in df.iterrows():
        exam_no = str(row[reg_col]).strip()
        student_name = str(row[name_col]).strip() if name_col and pd.notna(row.get(name_col)) else ""
        
        if exam_no not in student_dict:
            student_dict[exam_no] = {
                'EXAM NUMBER': exam_no,
                'NAME': student_name
            }
        
        for course_col in course_columns:
            score = row.get(course_col)
            if pd.notna(score) and score != "" and score != " ":
                column_name = course_col  # Use original column name
                student_dict[exam_no][column_name] = score
    
    transformed_data = list(student_dict.values())
    
    if transformed_data:
        transformed_df = pd.DataFrame(transformed_data)
        print(f"✅ Transformed data: {len(transformed_df)} students, {len(transformed_df.columns)} columns")
        return transformed_df
    else:
        print("❌ No data after transformation")
        return df

def detect_data_format(df, sheet_type):
    """Detect if data is in transposed format."""
    reg_col = find_exam_number_column(df)
    
    if not reg_col:
        return False
    
    student_counts = df[reg_col].value_counts()
    max_occurrences = student_counts.max()
    
    if max_occurrences > 1:
        print(f"📊 Data format detection for {sheet_type}:")
        print(f"   Total students: {len(student_counts)}")
        print(f"   Max occurrences per student: {max_occurrences}")
        print(f"   Students with multiple entries: {(student_counts > 1).sum()}")
        return True
    
    return False

# ----------------------------
# Enhanced Course Name Matching Functions (UPDATED)
# ----------------------------

def normalize_course_name(name):
    """Normalize course title for matching with improved logic."""
    if not isinstance(name, str):
        return ""
    
    # Convert to lowercase and strip
    normalized = name.lower().strip()
    
    # Remove punctuation and extra spaces
    normalized = re.sub(r'[^\w\s]', '', normalized)
    normalized = re.sub(r'\s+', '-', normalized)
    
    # Common substitutions and corrections
    substitutions = {
        'coomunication': 'communication',
        'communciation': 'communication',
        'nsg': 'nursing',
        'foundation': 'foundations',
        'of of': 'of',
        'emergency care': 'emergency',
        'nursing/ emergency': 'nursing emergency',
        'care i': 'care',
        'foundations of nursing': 'foundations nursing',
        'foundation of nsg': 'foundations nursing',
        'foundation of nursing': 'foundations nursing',
        'principles practice': 'principles and practice',
        'principles of nursing': 'principles and practice of nursing',
        '_ca': '',  # Remove _CA suffix from resit files
        'eed216': 'eed216',  # Keep as is for matching
        'nur221': 'nur221',
        'nur222': 'nur222',
        'nur223': 'nur223',
        'nur224': 'nur224',
        'nur225': 'nur225',
        'nur227': 'nur227',
        'nus221': 'nus221',
        'nus222': 'nus222',
    }
    
    for old, new in substitutions.items():
        normalized = normalized.replace(old, new)
        
    return normalized.strip()

def find_best_course_match(column_name, course_map):
    """Find the best matching course with enhanced debugging."""
    if not isinstance(column_name, str):
        return None
        
    normalized_column = normalize_course_name(column_name)
    print(f"  🔍 Matching: '{column_name}' -> normalized: '{normalized_column}'")
    
    # Exact match
    if normalized_column in course_map:
        print(f"  ✅ Exact match found: '{normalized_column}'")
        return course_map[normalized_column]
    
    # Contains match
    for course_norm, course_info in course_map.items():
        if course_norm in normalized_column or normalized_column in course_norm:
            print(f"  ✅ Contains match: '{normalized_column}' contains '{course_norm}'")
            return course_info
    
    # Word-based matching
    column_words = set(normalized_column.split('-'))
    best_match = None
    best_score = 0
    
    for course_norm, course_info in course_map.items():
        course_words = set(course_norm.split('-'))
        common_words = column_words.intersection(course_words)
        
        if common_words:
            score = len(common_words)
            # Weight important words more heavily
            key_words = ['nursing', 'foundation', 'emergency', 'care', 'health', 
                        'communication', 'anatomy', 'physiology', 'practice', 'principles']
            for word in key_words:
                if word in column_words and word in course_words:
                    score += 2
            
            if score > best_score:
                best_score = score
                best_match = course_info
    
    if best_match and best_score >= 2:
        print(f"  ✅ Word-based match (score: {best_score}): '{normalized_column}' -> '{best_match['code']}'")
        return best_match
    
    # Direct code matching for resit courses
    if len(column_name.strip()) <= 10:  # Likely a course code
        for norm_name, course_info in course_map.items():
            if course_info['code'].upper() == column_name.upper():
                print(f"  ✅ Direct code match: '{column_name}' -> '{course_info['code']}'")
                return course_info
    
    # Fuzzy matching as last resort
    best_match = None
    best_ratio = 0
    
    for course_norm, course_info in course_map.items():
        ratio = difflib.SequenceMatcher(None, normalized_column, course_norm).ratio()
        if ratio > best_ratio and ratio > 0.6:
            best_ratio = ratio
            best_match = course_info
    
    if best_match:
        print(f"  ✅ Fuzzy match (ratio: {best_ratio:.2f}): '{normalized_column}' -> '{best_match['code']}'")
    
    return best_match

def find_course_by_alternative_methods(column_name, course_map_normalized, semester_key):
    """Alternative course matching methods when primary method fails."""
    print(f"  🔍 Trying alternative matching for: '{column_name}'")
    
    # Method 1: Direct code matching
    if len(column_name.strip()) <= 10:  # Likely a course code
        for norm_name, course_info in course_map_normalized.items():
            if course_info['code'].upper() == column_name.upper():
                print(f"  ✅ Matched by direct code: '{column_name}' -> '{course_info['code']}'")
                return course_info
    
    # Method 2: Partial word matching
    column_words = set(column_name.lower().split())
    best_match = None
    best_score = 0
    
    for norm_name, course_info in course_map_normalized.items():
        course_words = set(norm_name.split('-'))
        common_words = column_words.intersection(course_words)
        
        if common_words:
            score = len(common_words)
            # Bonus for key words
            key_words = ['nursing', 'foundation', 'emergency', 'care', 'health', 'anatomy', 'physiology']
            for word in key_words:
                if word in column_words and word in course_words:
                    score += 2
            
            if score > best_score:
                best_score = score
                best_match = course_info
    
    if best_match and best_score >= 2:
        print(f"  ✅ Matched by partial words (score: {best_score}): '{column_name}' -> '{best_match['code']}'")
        return best_match
    
    return None

# ----------------------------
# Enhanced Debugging Functions (UPDATED)
# ----------------------------

def analyze_course_matching_debug(resit_df, mastersheet_df, course_map_normalized, semester_key):
    """Comprehensive analysis of course matching between files."""
    print(f"\n🔍 COMPREHENSIVE COURSE MATCHING ANALYSIS")
    print("=" * 60)
    
    # Get exam column from resit file
    exam_col = find_exam_number_column(resit_df)
    print(f"Resit file exam column: '{exam_col}'")
    
    # Analyze resit file columns
    print(f"\n📊 RESIT FILE COLUMNS ANALYSIS:")
    resit_courses = [col for col in resit_df.columns if col != exam_col and col != 'NAME' and not str(col).startswith('Unnamed')]
    print(f"Found {len(resit_courses)} course columns in resit file:")
    for i, course in enumerate(resit_courses):
        print(f"  {i+1}. '{course}'")
    
    # Analyze mastersheet columns
    print(f"\n📊 MASTERSHEET COLUMNS ANALYSIS:")
    mastersheet_courses = [col for col in mastersheet_df.columns 
                          if col not in ['S/N', 'EXAM NUMBER', 'NAME', 'REMARKS', 'CU Passed', 'CU Failed', 'TCPE', 'GPA', 'CGPA']]
    print(f"Found {len(mastersheet_courses)} course columns in mastersheet:")
    for i, course in enumerate(mastersheet_courses):
        print(f"  {i+1}. '{course}'")
    
    # Analyze course map
    print(f"\n📊 COURSE MAP ANALYSIS for {semester_key}:")
    print(f"Course map has {len(course_map_normalized)} entries")
    for i, (norm_name, course_info) in enumerate(list(course_map_normalized.items())):
        print(f"  {i+1}. Normalized: '{norm_name}' -> Code: '{course_info['code']}'")
    
    # Test matching for each resit course
    print(f"\n🔍 COURSE MATCHING TEST:")
    match_results = []
    for resit_course in resit_courses:
        course_info = find_best_course_match(resit_course, course_map_normalized)
        if course_info:
            status = "✅ MATCHED"
            match_results.append((resit_course, course_info['code'], status))
        else:
            status = "❌ NO MATCH"
            match_results.append((resit_course, "N/A", status))
    
    for resit_course, matched_code, status in match_results:
        print(f"  {status}: '{resit_course}' -> '{matched_code}'")
    
    matched_count = len([r for r in match_results if "MATCH" in r[2]])
    print(f"\n📊 MATCHING SUMMARY: {matched_count}/{len(resit_courses)} courses matched")
    
    return matched_count

def create_course_mapping_file(resit_file_path, mastersheet_path, semester_key, output_dir):
    """Create a manual course mapping file to help with matching issues."""
    print(f"\n📋 CREATING COURSE MAPPING ASSISTANT...")
    
    try:
        resit_df = pd.read_excel(resit_file_path, header=0)
        mastersheet_df = pd.read_excel(mastersheet_path, sheet_name=semester_key, header=5)
        
        resit_exam_col = find_exam_number_column(resit_df)
        resit_courses = [col for col in resit_df.columns 
                        if col != resit_exam_col and col != 'NAME' and not str(col).startswith('Unnamed')]
        
        mastersheet_courses = [col for col in mastersheet_df.columns 
                              if col not in ['S/N', 'EXAM NUMBER', 'NAME', 'REMARKS', 'CU Passed', 'CU Failed', 'TCPE', 'GPA', 'CGPA']]
        
        # Create mapping template
        mapping_data = []
        for resit_course in resit_courses:
            mapping_data.append({
                'Resit_Course_Name': resit_course,
                'Mastersheet_Course_Code': '',
                'Suggested_Match': '',
                'Match_Confidence': '',
                'Notes': ''
            })
        
        mapping_df = pd.DataFrame(mapping_data)
        mapping_file = os.path.join(output_dir, f"course_mapping_template_{semester_key}_{datetime.now().strftime(TIMESTAMP_FMT)}.xlsx")
        mapping_df.to_excel(mapping_file, index=False)
        
        print(f"✅ Course mapping template created: {mapping_file}")
        print(f"📝 Please manually map {len(resit_courses)} resit courses to {len(mastersheet_courses)} mastersheet courses")
        print(f"💡 Suggested matches will be provided in the template")
        
        return mapping_file
        
    except Exception as e:
        print(f"❌ Error creating mapping file: {e}")
        return None

def quick_diagnostic(resit_file_path, mastersheet_path, semester_key):
    """Run a quick diagnostic on the files."""
    print(f"\n🔍 QUICK DIAGNOSTIC")
    print("=" * 40)
    
    try:
        resit_df = pd.read_excel(resit_file_path, header=0)
        mastersheet_df = pd.read_excel(mastersheet_path, sheet_name=semester_key, header=5)
        
        print(f"📊 FILE ANALYSIS:")
        print(f"Resit file: {len(resit_df)} rows, {len(resit_df.columns)} columns")
        print(f"Mastersheet: {len(mastersheet_df)} students, {len(mastersheet_df.columns)} columns")
        
        resit_exam_col = find_exam_number_column(resit_df)
        print(f"Resit exam column: '{resit_exam_col}'")
        
        # Sample data
        print(f"\n📋 RESIT SAMPLE (first 3 rows):")
        print(resit_df.head(3))
        
        print(f"\n📋 MASTERSHEET SAMPLE (first 3 students):")
        print(mastersheet_df[['EXAM NUMBER', 'NAME']].head(3))
        
        # Course analysis
        resit_courses = [col for col in resit_df.columns if col != resit_exam_col and col != 'NAME' and not str(col).startswith('Unnamed')]
        mastersheet_courses = [col for col in mastersheet_df.columns if col not in ['S/N', 'EXAM NUMBER', 'NAME', 'REMARKS', 'CU Passed', 'CU Failed', 'TCPE', 'GPA', 'CGPA']]
        
        print(f"\n🎯 COURSE ANALYSIS:")
        print(f"Resit courses: {len(resit_courses)}")
        print(f"Mastersheet courses: {len(mastersheet_courses)}")
        print(f"First 5 resit courses: {resit_courses[:5]}")
        print(f"First 5 mastersheet courses: {mastersheet_courses[:5]}")
        
        return True
        
    except Exception as e:
        print(f"❌ Diagnostic failed: {e}")
        return False

# ----------------------------
# Carryover Management Functions
# ----------------------------

def identify_carryover_students(mastersheet_df, semester_key, set_name, pass_threshold=50.0):
    """Identify students with carryover courses."""
    carryover_students = []
    
    course_columns = [col for col in mastersheet_df.columns 
                     if col not in ['S/N', 'EXAM NUMBER', 'NAME', 'REMARKS', 
                                   'CU Passed', 'CU Failed', 'TCPE', 'GPA', 'CGPA']]
    
    for idx, student in mastersheet_df.iterrows():
        failed_courses = []
        exam_no = str(student['EXAM NUMBER']).strip()
        student_name = str(student.get('NAME', '')).strip()
        
        for course in course_columns:
            score = student.get(course, 0)
            try:
                score_val = float(score) if pd.notna(score) else 0
                if score_val < pass_threshold:
                    failed_courses.append({
                        'course_code': course,
                        'original_score': score_val,
                        'resit_attempts': 0,  # Added from fix 1
                        'best_score': score_val,  # Added from fix 1
                        'status': 'Failed',  # Added from fix 1
                        'semester': semester_key,
                        'set': set_name,
                        'identified_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
            }
            carryover_students.append(carryover_data)
            student_key = f"{exam_no}_{semester_key}"
            CARRYOVER_STUDENTS[student_key] = carryover_data
    
    # Fix 2: Add logging
    print(f"✅ Identified {len(carryover_students)} carryover students for {semester_key}")
    for student in carryover_students[:3]:  # Show first 3
        print(f"   - {student['exam_number']}: {len(student['failed_courses'])} failed courses")

    return carryover_students

def save_carryover_records(carryover_students, output_dir, set_name, semester_key):
    """Save carryover student records."""
    if not carryover_students:
        print("ℹ️ No carryover students to save")
        return None
    
    carryover_dir = os.path.join(output_dir, "CARRYOVER_RECORDS")
    os.makedirs(carryover_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime(TIMESTAMP_FMT)
    filename = f"co_student_{set_name}_{semester_key}_{timestamp}"
    
    excel_file = os.path.join(carryover_dir, f"{filename}.xlsx")
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
                'IDENTIFIED DATE': student['identified_date']
            })
    
    if records_data:
        df = pd.DataFrame(records_data)
        df.to_excel(excel_file, index=False)
        print(f"✅ Carryover records saved: {excel_file}")
    
    json_file = os.path.join(carryover_dir, f"{filename}.json")
    with open(json_file, 'w') as f:
        json.dump(carryover_students, f, indent=2)
    
    print(f"📁 Carryover records saved in: {carryover_dir}")
    return carryover_dir

def load_carryover_records(output_dir, set_name, semester_key):
    """Load carryover records."""
    carryover_dir = os.path.join(output_dir, "CARRYOVER_RECORDS")
    if not os.path.exists(carryover_dir):
        return None
    
    pattern = f"co_student_{set_name}_{semester_key}_*.json"
    matching_files = glob.glob(os.path.join(carryover_dir, pattern))
    
    if not matching_files:
        return None
    
    latest_file = sorted(matching_files)[-1]
    try:
        with open(latest_file, 'r') as f:
            carryover_data = json.load(f)
        print(f"✅ Loaded carryover records: {latest_file}")
        return carryover_data
    except Exception as e:
        print(f"❌ Error loading carryover records: {e}")
        return None

# ----------------------------
# Enhanced Resit Processing with Better Debugging (UPDATED)
# ----------------------------

def process_resit_file(resit_file_path, mastersheet_path, semester_key, set_name, pass_threshold, course_map, credit_units, output_dir, course_titles_map=None):
    """Enhanced resit processing with comprehensive debugging and improved matching."""
    print(f"\n🔄 ENHANCED RESIT PROCESSING WITH DEBUGGING")
    print("=" * 60)

    # Fix 1: Create CARRYOVER/RESIT output directory structure at the beginning
    timestamp = datetime.now().strftime(TIMESTAMP_FMT)
    carryover_output_dir = os.path.join(output_dir, f"CARRYOVER_{semester_key}_{timestamp}")
    os.makedirs(carryover_output_dir, exist_ok=True)

    if not os.path.exists(resit_file_path):
        print(f"❌ Resit file not found: {resit_file_path}")
        return False, 0, 0
    
    # FIX 3: Skip mastersheet check in resit mode - use the carryover_mastersheet_path after copy
    if not os.path.exists(mastersheet_path):
        print(f"❌ Mastersheet not found: {mastersheet_path}")
        return False, 0, 0
    
    try:
        # Fix 1: Copy original mastersheet to CARRYOVER folder with new name
        carryover_mastersheet_path = os.path.join(carryover_output_dir, f"CARRYOVER_mastersheet_{semester_key}_{timestamp}.xlsx")
        # Copy the original mastersheet to the new location (before updates)
        shutil.copy(mastersheet_path, carryover_mastersheet_path)
        print(f"✅ Copied original mastersheet to: {carryover_mastersheet_path}")

        # Read files with comprehensive analysis
        print(f"\n📖 LOADING AND ANALYZING FILES...")
        
        # Read resit file
        resit_df = pd.read_excel(resit_file_path, header=0)
        print(f"✅ Resit file loaded: {len(resit_df)} rows")
        print(f"📋 Resit file columns: {list(resit_df.columns)}")
        
        # Check for transposed data
        if detect_data_format(resit_df, "resit"):
            print("🔄 Resit file appears to be in transposed format - transforming...")
            original_rows = len(resit_df)
            resit_df = transform_transposed_data(resit_df, "resit")
            print(f"✅ Transformed: {original_rows} -> {len(resit_df)} rows")
        
        # Read mastersheet - FIX 3: Use the copied carryover mastersheet for resit mode
        mastersheet_df = pd.read_excel(carryover_mastersheet_path, sheet_name=semester_key, header=5)
        print(f"✅ Mastersheet loaded: {len(mastersheet_df)} students")
        
        # Add check after loading (Fix 3)
        if 'EXAM NUMBER' not in mastersheet_df.columns:
            print("❌ Mastersheet missing 'EXAM NUMBER' column")
            print(f"Mastersheet columns: {mastersheet_df.columns.tolist()}")
            return False, 0, 0
        
        # Find exam number columns
        resit_exam_col = find_exam_number_column(resit_df)
        mastersheet_exam_col = find_exam_number_column(mastersheet_df) or 'EXAM NUMBER'
        
        print(f"📝 Exam columns - Resit: '{resit_exam_col}', Mastersheet: '{mastersheet_exam_col}'")
        
        if not resit_exam_col:
            print("❌ Cannot find exam number column in resit file")
            print("Available columns:", list(resit_df.columns))
            return False, 0, 0
        
        # Normalize course map
        course_map_normalized = {}
        if semester_key in course_map:
            course_map_normalized = {normalize_course_name(k): v for k, v in course_map[semester_key].items()}
            print(f"✅ Course map normalized: {len(course_map_normalized)} courses")
        else:
            print(f"❌ No course map for semester: {semester_key}")
            return False, 0, 0
        
        # Comprehensive course matching analysis
        matched_count = analyze_course_matching_debug(resit_df, mastersheet_df, course_map_normalized, semester_key)
        
        if matched_count == 0:
            print(f"\n🚨 CRITICAL: No course matches found!")
            print("This usually happens when:")
            print("1. Resit file contains courses from a different semester")
            print("2. Course naming conventions don't match")
            print("3. The semester being processed doesn't match the resit file")
            
            # Show sample data to help debugging
            print(f"\n🔍 SAMPLE DATA FROM RESIT FILE:")
            print(resit_df.head(3))
            
            print(f"\n🔍 SAMPLE DATA FROM MASTERSHEET:")
            print(mastersheet_df[['EXAM NUMBER', 'NAME']].head(3))
            
            return False, 0, 0
        
        print(f"\n🎯 PROCEEDING WITH PROCESSING - {matched_count} potential course matches found")
        
        # Run deep debugging to understand why no updates might be happening
        print(f"\n🔍 RUNNING DEEP CARRYOVER ANALYSIS...")
        debug_file = debug_carryover_matching(resit_file_path, mastersheet_path, semester_key, set_name, pass_threshold, course_map, output_dir)
        
        # Process each student with enhanced debugging
        updated_count = 0
        students_updated = set()
        detailed_log = []
        
        print(f"\n👨‍🎓 PROCESSING STUDENTS...")
        
        for idx, row in resit_df.iterrows():
            exam_no = str(row[resit_exam_col]).strip().upper()
            if not exam_no or exam_no in ['NAN', 'NONE', '']:
                continue
            
            # Find student in mastersheet
            student_mask = mastersheet_df[mastersheet_exam_col].astype(str).str.strip().str.upper() == exam_no
            if not student_mask.any():
                detailed_log.append(f"❌ Student {exam_no} not found in mastersheet")
                continue
            
            student_data = mastersheet_df[student_mask].iloc[0]
            student_updated = False
            
            # Process each course column
            for col in resit_df.columns:
                if col == resit_exam_col or col == 'NAME' or 'Unnamed' in str(col):
                    continue
                    
                score = row.get(col)
                if pd.isna(score) or score == '' or score == ' ':
                    continue
                
                try:
                    score_val = float(score)
                    
                    # Find course match
                    course_info = find_best_course_match(col, course_map_normalized)
                    
                    if not course_info:
                        detailed_log.append(f"⚠️ {exam_no}: No match for '{col}'")
                        continue
                    
                    course_code = course_info['code']
                    
                    # Check if course exists in mastersheet
                    if course_code not in mastersheet_df.columns:
                        detailed_log.append(f"⚠️ {exam_no}: Course code '{course_code}' not in mastersheet")
                        continue
                    
                    # Get original score from mastersheet
                    original_score = student_data.get(course_code)
                    if pd.isna(original_score):
                        detailed_log.append(f"⚠️ {exam_no}: No original score for {course_code}")
                        continue
                    
                    try:
                        original_score_val = float(original_score)
                    except (ValueError, TypeError):
                        detailed_log.append(f"⚠️ {exam_no}: Invalid original score {original_score} for {course_code}")
                        continue
                    
                    # Check update condition
                    if original_score_val < pass_threshold and score_val >= pass_threshold:
                        print(f"✅ UPDATE: {exam_no} - {course_code}: {original_score_val} → {score_val}")
                        
                        # Update the mastersheet
                        mastersheet_df.loc[student_mask, course_code] = score_val
                        updated_count += 1
                        students_updated.add(exam_no)
                        student_updated = True
                        
                        detailed_log.append(f"✅ {exam_no}: Updated {course_code} from {original_score_val} to {score_val}")
                        
                        # Update carryover records
                        student_key = f"{exam_no}_{semester_key}"
                        if student_key in CARRYOVER_STUDENTS:
                            original_failed_count = len(CARRYOVER_STUDENTS[student_key]['failed_courses'])
                            CARRYOVER_STUDENTS[student_key]['failed_courses'] = [
                                c for c in CARRYOVER_STUDENTS[student_key]['failed_courses']
                                if c['course_code'] != course_code
                            ]
                            new_failed_count = len(CARRYOVER_STUDENTS[student_key]['failed_courses'])
                            print(f"  📊 Carryover records updated: {original_failed_count} → {new_failed_count} failed courses")
                            
                            # If no more failed courses, remove student from carryover
                            if not CARRYOVER_STUDENTS[student_key]['failed_courses']:
                                del CARRYOVER_STUDENTS[student_key]
                                print(f"  ✅ Removed {exam_no} from carryover records (all courses passed)")
                    
                    elif original_score_val >= pass_threshold:
                        detailed_log.append(f"ℹ️ {exam_no}: {course_code} already passed ({original_score_val})")
                    else:
                        detailed_log.append(f"ℹ️ {exam_no}: {course_code} still failed ({score_val})")
                        
                except Exception as e:
                    detailed_log.append(f"❌ {exam_no}: Error processing {col} - {e}")
                    continue
            
            if student_updated:
                print(f"✅ Student {exam_no} had courses updated")
        
        # Save detailed processing log
        log_file = os.path.join(output_dir, f"detailed_processing_log_{set_name}_{semester_key}_{datetime.now().strftime(TIMESTAMP_FMT)}.txt")
        with open(log_file, 'w') as f:
            f.write("\n".join(detailed_log))
        print(f"📝 Detailed log saved: {log_file}")
        
        print(f"\n📊 FINAL PROCESSING SUMMARY:")
        print(f"   Scores updated: {updated_count}")
        print(f"   Students updated: {len(students_updated)}")
        print(f"   Detailed log: {log_file}")
        
        # Add verify headers (Fix 5)
        print(f"Mastersheet columns after update: {mastersheet_df.columns.tolist()}")
        
        if updated_count > 0:
            print(f"\n🔄 UPDATING MASTERSHEET AND RECALCULATING GPAs...")
            
            # Recalculate metrics for updated students
            for exam_no in students_updated:
                student_mask = mastersheet_df[mastersheet_exam_col].astype(str).str.strip().str.upper() == exam_no
                
                total_grade_points = 0.0
                total_units = 0
                total_units_passed = 0
                total_units_failed = 0
                
                for code in credit_units.get(semester_key, {}):
                    if code not in mastersheet_df.columns:
                        continue
                    score = mastersheet_df.loc[student_mask, code].iloc[0]
                    if pd.isna(score):
                        continue
                    try:
                        score_val = float(score)
                        cu = credit_units[semester_key].get(code, 0)
                        grade_point = get_grade_point(score_val)
                        total_grade_points += grade_point * cu
                        total_units += cu
                        if score_val >= pass_threshold:
                            total_units_passed += cu
                        else:
                            total_units_failed += cu
                    except (ValueError, TypeError):
                        continue
                
                # Update metrics
                current_gpa = round(total_grade_points / total_units, 2) if total_units > 0 else 0.0
                mastersheet_df.loc[student_mask, 'CU Passed'] = total_units_passed
                mastersheet_df.loc[student_mask, 'CU Failed'] = total_units_failed
                mastersheet_df.loc[student_mask, 'TCPE'] = round(total_grade_points, 1)
                mastersheet_df.loc[student_mask, 'GPA'] = current_gpa
                mastersheet_df.loc[student_mask, 'REMARKS'] = determine_student_status(
                    {'GPA': current_gpa, 'CU Failed': total_units_failed}, total_units, pass_threshold
                )
                
                print(f"✅ Recalculated: {exam_no} - GPA: {current_gpa}")
            
            # Save updated mastersheet to the copied path (to avoid overwriting original)
            wb = load_workbook(carryover_mastersheet_path)
            if semester_key in wb.sheetnames:
                wb.remove(wb[semester_key])
            ws = wb.create_sheet(semester_key)
            
            # Write headers
            for col_idx, col_name in enumerate(mastersheet_df.columns, 1):
                ws.cell(row=6, column=col_idx).value = col_name
            
            # Write data
            for row_idx, row_data in enumerate(mastersheet_df.itertuples(index=False), 7):
                for col_idx, value in enumerate(row_data, 1):
                    ws.cell(row=row_idx, column=col_idx).value = value
            
            wb.save(carryover_mastersheet_path)
            print(f"💾 Updated mastersheet saved: {carryover_mastersheet_path}")

            # Fix 1: Generate individual student PDFs for updated students
            previous_gpas = load_previous_gpas_from_processed_files(output_dir, semester_key, timestamp)
            cgpa_data = load_all_previous_gpas_for_cgpa(output_dir, semester_key, timestamp)
            ordered_codes = list(credit_units.get(semester_key, {}).keys())
            filtered_credit_units = credit_units.get(semester_key, {})
            filtered_course_titles = course_titles_map.get(semester_key, {}) if course_titles_map else {}
            for exam_no in students_updated:
                student_df = mastersheet_df[mastersheet_df['EXAM NUMBER'] == exam_no]
                # FIX 1: Sanitize exam_no for filename
                exam_no_safe = exam_no.replace('/', '_').replace('\\', '_')  # Replace slashes
                pdf_path = os.path.join(carryover_output_dir, f"CARRYOVER_student_{exam_no_safe}_{semester_key}_{timestamp}.pdf")
                generate_individual_student_pdf(
                    student_df, pdf_path, semester_key, logo_path=DEFAULT_LOGO_PATH,
                    filtered_credit_units=filtered_credit_units,
                    ordered_codes=ordered_codes,
                    course_titles_map=filtered_course_titles,
                    pass_threshold=pass_threshold,
                    previous_gpas=previous_gpas,
                    cgpa_data=cgpa_data
                )
                print(f"✅ Generated PDF for {exam_no}: {pdf_path}")

        # Fix 1: Before return, copy/rename all generated files to have CARRYOVER_ prefix (if not already)
        # Assuming PDFs and mastersheet already have prefix, and logs can be renamed if needed
        for file in os.listdir(output_dir):
            if file.startswith("detailed_processing_log") or file.startswith("deep_debug_carryover"):
                src = os.path.join(output_dir, file)
                dst = os.path.join(carryover_output_dir, f"CARRYOVER_{file}")
                shutil.move(src, dst)
                print(f"✅ Moved/renamed: {dst}")

        return True, updated_count, len(students_updated)
        
    except Exception as e:
        print(f"❌ Error in enhanced processing: {e}")
        traceback.print_exc()
        return False, 0, 0

# ----------------------------
# GPA and CGPA Calculations
# ----------------------------

def load_previous_gpas_from_processed_files(output_dir, current_semester_key, timestamp):
    """Load previous GPA data from mastersheets."""
    previous_gpas = {}
    current_year, current_semester_num, _, _, _ = get_semester_display_info(current_semester_key)

    if current_semester_num == 1 and current_year == 1:
        print("📊 First semester of first year - no previous GPA available")
        return previous_gpas
    
    prev_semester = {
        (1, 2): "ND-FIRST-YEAR-FIRST-SEMESTER",
        (2, 1): "ND-FIRST-YEAR-SECOND-SEMESTER",
        (2, 2): "ND-SECOND-YEAR-FIRST-SEMESTER"
    }.get((current_year, current_semester_num))

    if not prev_semester:
        print(f"⚠️ Unknown semester combination: Year {current_year}, Semester {current_semester_num}")
        return previous_gpas

    mastersheet_path = os.path.join(output_dir, f"mastersheet_{timestamp}.xlsx")
    if not os.path.exists(mastersheet_path):
        print(f"❌ Mastersheet not found: {mastersheet_path}")
        return previous_gpas

    try:
        df = pd.read_excel(mastersheet_path, sheet_name=prev_semester, header=5)
        exam_col = find_exam_number_column(df)
        gpa_col = None
        for col in df.columns:
            if 'GPA' in str(col).upper():
                gpa_col = col
                break
        
        if exam_col and gpa_col:
            for idx, row in df.iterrows():
                exam_no = str(row[exam_col]).strip()
                gpa = row[gpa_col]
                if pd.notna(gpa) and pd.notna(exam_no) and exam_no != 'nan':
                    try:
                        previous_gpas[exam_no] = float(gpa)
                    except (ValueError, TypeError):
                        continue
            print(f"✅ Loaded previous GPAs for {len(previous_gpas)} students from {prev_semester}")
    except Exception as e:
        print(f"⚠️ Could not load data from {prev_semester}: {e}")
    
    return previous_gpas

def load_all_previous_gpas_for_cgpa(output_dir, current_semester_key, timestamp):
    """Load all previous GPAs for CGPA calculation."""
    all_student_data = {}
    current_year, current_semester_num, _, _, _ = get_semester_display_info(current_semester_key)

    semesters_to_load = []
    if current_semester_num == 1 and current_year == 1:
        return {}
    elif current_semester_num == 2 and current_year == 1:
        semesters_to_load = ["ND-FIRST-YEAR-FIRST-SEMESTER"]
    elif current_semester_num == 1 and current_year == 2:
        semesters_to_load = ["ND-FIRST-YEAR-FIRST-SEMESTER", "ND-FIRST-YEAR-SECOND-SEMESTER"]
    elif current_semester_num == 2 and current_year == 2:
        semesters_to_load = ["ND-FIRST-YEAR-FIRST-SEMESTER", "ND-FIRST-YEAR-SECOND-SEMESTER", 
                            "ND-SECOND-YEAR-FIRST-SEMESTER"]

    mastersheet_path = os.path.join(output_dir, f"mastersheet_{timestamp}.xlsx")
    if not os.path.exists(mastersheet_path):
        print(f"❌ Mastersheet not found: {mastersheet_path}")
        return {}

    for semester in semesters_to_load:
        try:
            df = pd.read_excel(mastersheet_path, sheet_name=semester, header=5)
            exam_col = find_exam_number_column(df)
            gpa_col = None
            credit_col = None
            for col in df.columns:
                col_str = str(col).upper()
                if 'GPA' in col_str:
                    gpa_col = col
                if 'CU PASSED' in col_str:
                    credit_col = col
            
            if exam_col and gpa_col:
                for idx, row in df.iterrows():
                    exam_no = str(row[exam_col]).strip()
                    gpa = row[gpa_col]
                    credits = int(row[credit_col]) if credit_col and pd.notna(row[credit_col]) else 30
                    
                    if pd.notna(gpa) and pd.notna(exam_no) and exam_no != 'nan':
                        try:
                            if exam_no not in all_student_data:
                                all_student_data[exam_no] = {'gpas': [], 'credits': []}
                            all_student_data[exam_no]['gpas'].append(float(gpa))
                            all_student_data[exam_no]['credits'].append(credits)
                        except (ValueError, TypeError):
                            continue
        except Exception as e:
            print(f"⚠️ Could not load data from {semester}: {e}")
    
    print(f"📊 Loaded cumulative data for {len(all_student_data)} students")
    return all_student_data

def calculate_cgpa(student_data, current_gpa, current_credits):
    """Calculate Cumulative GPA."""
    if not student_data:
        return current_gpa

    total_grade_points = 0.0
    total_credits = 0

    for prev_gpa, prev_credits in zip(student_data['gpas'], student_data['credits']):
        total_grade_points += prev_gpa * prev_credits
        total_credits += prev_credits

    total_grade_points += current_gpa * current_credits
    total_credits += current_credits

    return round(total_grade_points / total_credits, 2) if total_credits > 0 else current_gpa

# ----------------------------
# Student Status and Remarks
# ----------------------------

def determine_student_status(row, total_cu, pass_threshold):
    """Determine student status."""
    gpa = row.get("GPA", 0)
    cu_failed = row.get("CU Failed", 0)
    failed_percentage = (cu_failed / total_cu) * 100 if total_cu > 0 else 0

    if cu_failed == 0:
        return "Pass"
    elif gpa >= 2.0 and failed_percentage <= 45:
        return "Carry Over"
    elif gpa < 2.0 and failed_percentage <= 45:
        return "Probation"
    elif failed_percentage > 45:
        return "Withdrawn"
    return "Carry Over"

def format_failed_courses_remark(failed_courses, max_line_length=60):
    """Format failed courses remark."""
    if not failed_courses:
        return [""]

    failed_str = ", ".join(sorted(failed_courses))
    if len(failed_str) <= max_line_length:
        return [failed_str]

    lines = []
    current_line = ""
    for course in sorted(failed_courses):
        if not current_line:
            current_line = course
        elif len(current_line) + len(course) + 2 <= max_line_length:
            current_line += ", " + course
        else:
            lines.append(current_line)
            current_line = course
    if current_line:
        lines.append(current_line)
    return lines

# ----------------------------
# Load Course Data
# ----------------------------

def load_course_data():
    """Load course data from course-code-creditUnit.xlsx."""
    course_file = os.path.join(ND_COURSES_DIR, "course-code-creditUnit.xlsx")
    print(f"Loading course data from: {course_file}")
    if not os.path.exists(course_file):
        raise FileNotFoundError(f"Course file not found: {course_file}")

    xl = pd.ExcelFile(course_file)
    semester_course_maps = {}
    semester_credit_units = {}
    semester_lookup = {}
    semester_course_titles = {}

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
        cus = dfx['CU'].astype(float).tolist()

        # FIXED: Correct indentation for this block
        enhanced_course_map = {}
        for code, title in zip(codes, titles):
            normalized_title = normalize_course_name(title)
            enhanced_course_map[normalized_title] = {
                'original_name': title,
                'code': code,
                'normalized': normalized_title
            }
            
        semester_course_maps[sheet] = enhanced_course_map
        semester_credit_units[sheet] = dict(zip(codes, cus))
        semester_course_titles[sheet] = dict(zip(codes, titles))

        norm = normalize_for_matching(sheet)
        semester_lookup[norm] = sheet
        norm_no_nd = norm.replace('nd-', '').replace('nd ', '')
        semester_lookup[norm_no_nd] = sheet
        norm_hyphen = norm.replace('-', ' ')
        semester_lookup[norm_hyphen] = sheet
        norm_space = norm.replace(' ', '-')
        semester_lookup[norm_space] = sheet

    if not semester_course_maps:
        raise ValueError("No course data loaded from course workbook")
    print(f"Loaded course sheets: {list(semester_course_maps.keys())}")
    return semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles

# ----------------------------
# Semester Detection and Matching
# ----------------------------

def detect_semester_from_filename(filename):
    """Detect semester from filename, handling carryover/resit patterns robustly."""
    filename_upper = filename.upper()
    
    # Strip extension, carryover/resit prefix, and set suffix
    filename_upper = re.sub(r'\.(XLSX|XLS)$', '', filename_upper)
    filename_upper = re.sub(r'^(CARRYOVER-|RESIT-)', '', filename_upper)
    filename_upper = re.sub(r'-ND-\d{4}$', '', filename_upper)
    
    # Try to match full semester pattern first
    match = re.search(r'((?:ND-|BN-|BM-)?(?:FIRST|SECOND|THIRD)[-_]YEAR[-_](?:FIRST|SECOND)[-_]SEMESTER)', filename_upper)
    if match:
        detected_sem = match.group(1).replace('_', '-')
        # If no prefix, add default 'ND-'
        if not re.match(r'^[A-Z]+-', detected_sem):
            detected_sem = 'ND-' + detected_sem
        
        # Extract year and semester numbers
        if 'FIRST-YEAR-FIRST-SEMESTER' in detected_sem:
            return detected_sem, 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
        elif 'FIRST-YEAR-SECOND-SEMESTER' in detected_sem:
            return detected_sem, 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI"
        elif 'SECOND-YEAR-FIRST-SEMESTER' in detected_sem:
            return detected_sem, 2, 1, "YEAR TWO", "FIRST SEMESTER", "NDII"
        elif 'SECOND-YEAR-SECOND-SEMESTER' in detected_sem:
            return detected_sem, 2, 2, "YEAR TWO", "SECOND SEMESTER", "NDII"
        elif 'THIRD-YEAR-FIRST-SEMESTER' in detected_sem:
            return detected_sem, 3, 1, "YEAR THREE", "FIRST SEMESTER", "NDIII"
        elif 'THIRD-YEAR-SECOND-SEMESTER' in detected_sem:
            return detected_sem, 3, 2, "YEAR THREE", "SECOND SEMESTER", "NDIII"
    
    # Fallback to simple pattern matching
    if 'FIRST' in filename_upper and 'SECOND' not in filename_upper:
        return "ND-FIRST-YEAR-FIRST-SEMESTER", 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
    elif 'SECOND' in filename_upper:
        return "ND-FIRST-YEAR-SECOND-SEMESTER", 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI"
    
    # Default fallback
    print(f"⚠️ Could not detect semester from filename: {filename}, defaulting to ND-FIRST-YEAR-FIRST-SEMESTER")
    return "ND-FIRST-YEAR-FIRST-SEMESTER", 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"

def get_semester_display_info(semester_key):
    """Get display information for a semester key."""
    semester_lower = semester_key.lower()
    if 'first-year-first-semester' in semester_lower:
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
    elif 'first-year-second-semester' in semester_lower:
        return 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI"
    elif 'second-year-first-semester' in semester_lower:
        return 2, 1, "YEAR TWO", "FIRST SEMESTER", "NDII"
    elif 'second-year-second-semester' in semester_lower:
        return 2, 2, "YEAR TWO", "SECOND SEMESTER", "NDII"
    else:
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"

def match_semester_from_filename(fname, semester_lookup):
    """Match semester using lookup table."""
    fn = normalize_for_matching(fname)
    for norm, sheet in semester_lookup.items():
        if norm in fn:
            return sheet
    keys = list(semester_lookup.keys())
    best = difflib.get_close_matches(fn, keys, n=1, cutoff=0.55)
    if best:
        return semester_lookup[best[0]]
    sem, _, _, _, _, _ = detect_semester_from_filename(fname)
    return sem

def find_column_by_names(df, candidate_names):
    """Find column by possible names."""
    norm_map = {col: re.sub(r'\s+', ' ', str(col).strip().lower()) for col in df.columns}
    candidates = [re.sub(r'\s+', ' ', c.strip().lower()) for c in candidate_names]
    for cand in candidates:
        for col, ncol in norm_map.items():
            if ncol == cand:
                return col
    return None

def find_exam_number_column(df):
    """Find the exam number column in a DataFrame."""
    possible_names = ['EXAM NUMBER', 'REG. No', 'REG NO', 'REGISTRATION NUMBER', 'MAT NO', 'STUDENT ID']
    for col in df.columns:
        col_upper = str(col).upper()
        for possible_name in possible_names:
            if possible_name in col_upper:
                return col
    # Robust finder (Fix 5)
    for col in df.columns:
        col_upper = col.upper().strip()
        if 'EXAM' in col_upper and 'NUMBER' in col_upper:
            return col
    return None

# ----------------------------
# PDF Generation
# ----------------------------

def generate_individual_student_pdf(mastersheet_df, out_pdf_path, semester_key, logo_path=None,
                                  filtered_credit_units=None, ordered_codes=None, course_titles_map=None,
                                  pass_threshold=None, previous_gpas=None, cgpa_data=None):
    """Generate individual student PDF reports."""
    doc = SimpleDocTemplate(out_pdf_path, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()

    header_style = ParagraphStyle('CustomHeader', parent=styles['Normal'], fontSize=10, alignment=TA_CENTER, spaceAfter=2)
    main_header_style = ParagraphStyle('MainHeader', parent=styles['Normal'], fontSize=16, alignment=TA_CENTER, 
                                      fontName='Helvetica-Bold', spaceAfter=6, textColor=colors.HexColor("#800080"))
    title_style = ParagraphStyle('CustomTitle', parent=styles['Normal'], fontSize=12, alignment=TA_CENTER, 
                                fontName='Helvetica-Bold', spaceAfter=4)
    subtitle_style = ParagraphStyle('SubtitleStyle', parent=styles['Normal'], fontSize=10, alignment=TA_CENTER, 
                                   spaceAfter=10, textColor=colors.red)
    left_align_style = ParagraphStyle('LeftAlign', parent=styles['Normal'], fontSize=9, alignment=TA_LEFT, leftIndent=4)
    center_align_style = ParagraphStyle('CenterAlign', parent=styles['Normal'], fontSize=9, alignment=TA_CENTER)
    remarks_style = ParagraphStyle('RemarksStyle', parent=styles['Normal'], fontSize=8, alignment=TA_LEFT)

    elems = []

    for idx, r in mastersheet_df.iterrows():
        logo_img = None
        if logo_path and os.path.exists(logo_path):
            try:
                logo_img = Image(logo_path, width=0.8 * inch, height=0.8 * inch)
            except Exception as e:
                print(f"Warning: Could not load logo: {e}")

        if logo_img:
            header_data = [[logo_img, Paragraph("FCT COLLEGE OF NURSING SCIENCES", main_header_style)]]
            header_table = Table(header_data, colWidths=[1.0 * inch, 5.0 * inch])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            ]))
            elems.append(header_table)
        else:
            elems.append(Paragraph("FCT COLLEGE OF NURSING SCIENCES", main_header_style))

        elems.append(Paragraph("P.O.Box 507, Gwagwalada-Abuja, Nigeria", header_style))
        elems.append(Paragraph("<b>EXAMINATIONS OFFICE</b>", header_style))
        elems.append(Paragraph("fctsonexamsoffice@gmail.com", header_style))
        elems.append(Spacer(1, 8))
        elems.append(Paragraph("STUDENT'S ACADEMIC PROGRESS REPORT", title_style))
        elems.append(Paragraph("(THIS IS NOT A TRANSCRIPT)", subtitle_style))
        elems.append(Spacer(1, 8))

        exam_no = str(r.get("EXAM NUMBER", "")).strip()
        student_name = str(r.get("NAME", "")).strip()
        year, semester_num, level_display, semester_display, set_code = get_semester_display_info(semester_key)

        particulars_data = [
            [Paragraph("<b>STUDENT'S PARTICULARS</b>", styles['Normal'])],
            [Paragraph("<b>NAME:</b>", styles['Normal']), student_name],
            [Paragraph("<b>LEVEL OF<br/>STUDY:</b>", styles['Normal']), level_display,
             Paragraph("<b>SEMESTER:</b>", styles['Normal']), semester_display],
            [Paragraph("<b>REG NO.</b>", styles['Normal']), exam_no,
             Paragraph("<b>SET:</b>", styles['Normal']), set_code],
        ]

        particulars_table = Table(particulars_data, colWidths=[1.2 * inch, 2.3 * inch, 0.8 * inch, 1.5 * inch])
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

        passport_data = [[Paragraph("Affix Recent<br/>Passport<br/>Photograph", styles['Normal'])]]
        passport_table = Table(passport_data, colWidths=[1.5 * inch], rowHeights=[1.2 * inch])
        passport_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
        ]))

        combined_data = [[particulars_table, passport_table]]
        combined_table = Table(combined_data, colWidths=[5.8 * inch, 1.5 * inch])
        combined_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ]))
        elems.append(combined_table)
        elems.append(Spacer(1, 12))

        elems.append(Paragraph("<b>SEMESTER RESULT</b>", title_style))
        elems.append(Spacer(1, 6))

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
            if pd.isna(score):
                continue
            try:
                score_val = float(score)
                score_display = str(int(round(score_val)))
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

        course_table = Table(course_data, colWidths=[0.4 * inch, 0.7 * inch, 2.8 * inch, 0.6 * inch, 0.6 * inch, 0.6 * inch])
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

        current_gpa = round(total_grade_points / total_units, 2) if total_units > 0 else 0.0
        exam_no = str(r.get("EXAM NUMBER", "")).strip()
        previous_gpa = previous_gpas.get(exam_no, None) if previous_gpas else None
        cgpa = calculate_cgpa(cgpa_data.get(exam_no, {}), current_gpa, total_units) if cgpa_data else current_gpa

        tcpe = round(total_grade_points, 1)
        tcup = total_units_passed
        tcuf = total_units_failed
        student_status = determine_student_status(r, total_units, pass_threshold)
        withdrawal_history = get_withdrawal_history(exam_no)
        previously_withdrawn = withdrawal_history is not None

        failed_courses_formatted = format_failed_courses_remark(failed_courses_list)
        final_remarks_lines = []

        if previously_withdrawn and withdrawal_history['withdrawn_semester'] == semester_key:
            if failed_courses_formatted:
                final_remarks_lines.append(f"Failed: {failed_courses_formatted[0]}")
                final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("Advised to Withdraw")
            else:
                final_remarks_lines.append("Advised to Withdraw")
        elif previously_withdrawn:
            year, sem_num, level, sem_display, set_code = get_semester_display_info(withdrawal_history['withdrawn_semester'])
            final_remarks_lines.append(f"STUDENT WAS WITHDRAWN FROM {level} - {sem_display}")
        elif student_status == "Pass":
            final_remarks_lines.append("Passed")
        elif student_status == "Carry Over":
            if failed_courses_formatted:
                final_remarks_lines.append(f"Failed: {failed_courses_formatted[0]}")
                final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("To Carry Over Courses")
            else:
                final_remarks_lines.append("To Carry Over Courses")
        elif student_status == "Probation":
            if failed_courses_formatted:
                final_remarks_lines.append(f"Failed: {failed_courses_formatted[0]}")
                final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("Placed on Probation")
            else:
                final_remarks_lines.append("Placed on Probation")
        elif student_status == "Withdrawn":
            if failed_courses_formatted:
                final_remarks_lines.append(f"Failed: {failed_courses_formatted[0]}")
                final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("Advised to Withdraw")
            else:
                final_remarks_lines.append("Advised to Withdraw")

        final_remarks = "<br/>".join(final_remarks_lines)
        display_gpa = current_gpa
        display_cgpa = cgpa

        summary_data = [
            [Paragraph("<b>SUMMARY</b>", styles['Normal']), "", "", ""],
            [Paragraph("<b>TCPE:</b>", styles['Normal']), str(tcpe),
             Paragraph("<b>CURRENT GPA:</b>", styles['Normal']), str(display_gpa)],
        ]

        if previous_gpa is not None:
            summary_data.append([
                Paragraph("<b>TCUP:</b>", styles['Normal']), str(tcup),
                Paragraph("<b>PREVIOUS GPA:</b>", styles['Normal']), str(previous_gpa)
            ])
        else:
            summary_data.append([Paragraph("<b>TCUP:</b>", styles['Normal']), str(tcup), "", ""])

        if cgpa is not None:
            summary_data.append([
                Paragraph("<b>TCUF:</b>", styles['Normal']), str(tcuf),
                Paragraph("<b>OVERALL GPA:</b>", styles['Normal']), str(display_cgpa)
            ])
        else:
            summary_data.append([Paragraph("<b>TCUF:</b>", styles['Normal']), str(tcuf), "", ""])

        remarks_paragraph = Paragraph(final_remarks, remarks_style)
        summary_data.append([Paragraph("<b>REMARKS:</b>", styles['Normal']), remarks_paragraph, "", ""])

        row_heights = [0.3 * inch] * len(summary_data)
        total_remark_lines = len(final_remarks_lines)
        if total_remark_lines > 1:
            row_heights[-1] = max(0.4 * inch, 0.2 * inch * (total_remark_lines + 1))

        summary_table = Table(summary_data, colWidths=[1.5 * inch, 1.0 * inch, 1.5 * inch, 1.0 * inch], rowHeights=row_heights)
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

        sig_data = [["", ""], ["____________________", "____________________"],
                    [Paragraph("<b>EXAMS SECRETARY</b>", ParagraphStyle('SigStyle', parent=styles['Normal'], fontSize=10, alignment=TA_CENTER)),
                     Paragraph("<b>V.P. ACADEMICS</b>", ParagraphStyle('SigStyle', parent=styles['Normal'], fontSize=10, alignment=TA_CENTER))]]
        sig_table = Table(sig_data, colWidths=[3.0 * inch, 3.0 * inch])
        sig_table.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'TOP')]))
        elems.append(sig_table)

        if idx < len(mastersheet_df) - 1:
            elems.append(PageBreak())

    doc.build(elems)
    print(f"✅ Individual student PDF written: {out_pdf_path}")

# ----------------------------
# FIX 6: Generate New Carryover Mastersheet
# ----------------------------

def generate_resit_carryover_mastersheet(updated_mastersheet_df, output_dir, timestamp, semester_key, set_name, pass_threshold, filtered_credit_units, ordered_codes, course_titles_map):
    """Generate updated carryover mastersheet in original format with summary block."""
    wb = Workbook()
    ws = wb.active
    ws.title = semester_key
    
    # Add logo if available
    logo_path = DEFAULT_LOGO_PATH
    if logo_path and os.path.exists(logo_path):
        img = openpyxl.drawing.image.Image(logo_path)
        ws.add_image(img, 'A1')
    
    # Title and subtitle
    ws.merge_cells('A3:J3')
    title_cell = ws['A3']
    title_cell.value = "FCT COLLEGE OF NURSING SCIENCES"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center')
    
    ws.merge_cells('A4:J4')
    subtitle_cell = ws['A4']
    subtitle_cell.value = "CARRYOVER SUMMARY - " + semester_key
    subtitle_cell.font = Font(size=12)
    subtitle_cell.alignment = Alignment(horizontal='center')
    
    # Course titles and CU
    course_row = 5
    cu_row = 6
    ws.cell(row=course_row, column=3, value="COURSE TITLES")
    ws.cell(row=cu_row, column=3, value="CU")
    col = 4
    for code in ordered_codes:
        ws.cell(row=course_row, column=col, value=course_titles_map.get(code, code))
        ws.cell(row=cu_row, column=col, value=filtered_credit_units.get(code, 0))
        col += 1
    
    # Headers
    headers = ['S/N', 'EXAM NUMBER', 'NAME'] + ordered_codes + ['REMARKS', 'CU Passed', 'CU Failed', 'TCPE', 'GPA', 'CGPA']
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    # Append data
    for idx, row in updated_mastersheet_df.iterrows():
        row_num = idx + 8
        ws.cell(row=row_num, column=1, value=idx + 1)
        ws.cell(row=row_num, column=2, value=row['EXAM NUMBER'])
        ws.cell(row=row_num, column=3, value=row['NAME'])
        col = 4
        for code in ordered_codes:
            ws.cell(row=row_num, column=col, value=row.get(code, ''))
            col += 1
        ws.cell(row=row_num, column=col, value=row['REMARKS'])
        ws.cell(row=row_num, column=col + 1, value=row['CU Passed'])
        ws.cell(row=row_num, column=col + 2, value=row['CU Failed'])
        ws.cell(row=row_num, column=col + 3, value=row['TCPE'])
        ws.cell(row=row_num, column=col + 4, value=row['GPA'])
        ws.cell(row=row_num, column=col + 5, value=row['CGPA'])
    
    # Fails per course
    fails_row = len(updated_mastersheet_df) + 10
    ws.cell(row=fails_row, column=3, value="Fails per course")
    col = 4
    for code in ordered_codes:
        fails = updated_mastersheet_df[code][updated_mastersheet_df[code] < pass_threshold].count()
        ws.cell(row=fails_row, column=col, value=fails)
        col += 1
    
    # Comprehensive summary block
    summary_row = fails_row + 2
    ws.cell(row=summary_row, column=3, value="SUMMARY")
    ws.cell(row=summary_row + 1, column=3, value="Total Students")
    ws.cell(row=summary_row + 1, column=4, value=len(updated_mastersheet_df))
    ws.cell(row=summary_row + 2, column=3, value="Passed All")
    passed_all = (updated_mastersheet_df['CU Failed'] == 0).sum()
    ws.cell(row=summary_row + 2, column=4, value=passed_all)
    ws.cell(row=summary_row + 3, column=3, value="Carryover")
    carryover = (updated_mastersheet_df['REMARKS'] == "Carry Over").sum()
    ws.cell(row=summary_row + 3, column=4, value=carryover)
    ws.cell(row=summary_row + 4, column=3, value="Probation")
    probation = (updated_mastersheet_df['REMARKS'] == "Probation").sum()
    ws.cell(row=summary_row + 4, column=4, value=probation)
    ws.cell(row=summary_row + 5, column=3, value="Withdrawn")
    withdrawn = (updated_mastersheet_df['REMARKS'] == "Withdrawn").sum()
    ws.cell(row=summary_row + 5, column=4, value=withdrawn)
    
    # Styles and borders (simplified)
    thin = Side(border_style="thin", color="000000")
    for row in ws.iter_rows(min_row=3, max_row=summary_row + 5, min_col=3, max_col=col):
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
    
    timestamp_fmt = datetime.now().strftime(TIMESTAMP_FMT)
    output_file = os.path.join(output_dir, f"resit_carryover_mastersheet_{set_name}_{semester_key}_{timestamp_fmt}.xlsx")
    wb.save(output_file)
    print(f"✅ Resit carryover mastersheet saved: {output_file}")
    return output_file

# ----------------------------
# Enhanced Semester Processing
# ----------------------------

def process_semester_files_enhanced(semester_key, raw_files, raw_dir, output_dir, timestamp, pass_threshold, 
                                    semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles, logo_path, set_name, process_resit=False, resit_file_path=None):
    """Enhanced semester processing with better resit handling and CARRYOVER subdirectory support."""
    print(f"\n📂 Processing semester {semester_key} for set {set_name}")
    print(f"📋 Raw files provided: {raw_files}")
    print(f"🔄 Resit processing enabled: {process_resit}, Resit file: {resit_file_path}")

    mastersheet_path = os.path.join(output_dir, f"mastersheet_{timestamp}.xlsx")
    wb = Workbook()
    ws = wb.create_sheet(semester_key)
    
    try:
        mastersheet_df = None
        course_codes = list(semester_credit_units.get(semester_key, {}).keys())
        print(f"📚 Course codes for {semester_key}: {course_codes}")
        
        # Fix 1: Update file filtering
        normalized_key = semester_key.replace('ND-', '').upper()
        
        files = [f for f in raw_files 
                 if normalized_key in f.upper().replace('ND-', '')]
        
        # Check CARRYOVER subdirectory for additional files
        carryover_subdir = os.path.join(raw_dir, "CARRYOVER")
        carryover_files = []
        if os.path.exists(carryover_subdir):
            resit_files = [f for f in os.listdir(carryover_subdir) 
                           if f.lower().endswith(('.xlsx', '.xls')) 
                           and normalized_key in f.upper().replace('ND-', '')]
            carryover_files = [os.path.join('CARRYOVER', f) for f in resit_files]
        
        # Fix 1: Do NOT add carryover_files when process_resit=True
        if not process_resit:
            all_files = files + carryover_files
        else:
            all_files = files  # Only regular files (but in resit mode, we'll skip this loop later via main() fix)
        
        print(f"📑 All files to process: {all_files}")
        
        processed_files = []
        for rf in all_files:
            detected_sem, _, _, _, _, _ = detect_semester_from_filename(rf)
            print(f"🔍 File {rf}: Detected semester {detected_sem}")
            
            # Normalize for comparison (ignore prefixes like 'ND-', 'BN-', 'BM-')
            normalized_detected = re.sub(r'^[A-Z]+-', '', detected_sem) if detected_sem else None
            normalized_key = re.sub(r'^[A-Z]+-', '', semester_key)
            
            if normalized_detected != normalized_key:
                print(f"⏭️ Skipping {rf}: Semester mismatch (detected {detected_sem}, expected {semester_key})")
                continue
            
            file_path = os.path.join(raw_dir, rf)
            if not os.path.exists(file_path):
                print(f"⚠️ File not found: {file_path}")
                continue
            
            df = pd.read_excel(file_path, header=0)
            print(f"✅ Read file {rf}: {len(df)} rows")
            
            if detect_data_format(df, semester_key):
                print(f"🔄 Transforming transposed data for {rf}")
                df = transform_transposed_data(df, semester_key)
            
            if mastersheet_df is None:
                mastersheet_df = df.copy()
                exam_col = find_exam_number_column(df)
                mastersheet_df['EXAM NUMBER'] = mastersheet_df[exam_col]
                mastersheet_df['NAME'] = mastersheet_df.get('NAME', '')
            else:
                mastersheet_df = mastersheet_df.merge(df, on=['EXAM NUMBER', 'NAME'], how='outer')
            
            processed_files.append(rf)
            print(f"✓ Processed file: {rf}")
        
        if not processed_files:
            print(f"⚠️ No files processed for {semester_key}. Available files: {all_files}")
            return
        
        if mastersheet_df is None:
            print(f"⚠️ No data processed for {semester_key}")
            return
        
        print(f"✅ Processed {len(processed_files)} files: {processed_files}")
        
        # Initialize columns
        mastersheet_df['CU Passed'] = 0
        mastersheet_df['CU Failed'] = 0
        mastersheet_df['TCPE'] = 0.0
        mastersheet_df['GPA'] = 0.0
        mastersheet_df['CGPA'] = 0.0
        mastersheet_df['REMARKS'] = ''
        
        # Calculate metrics
        params = get_form_parameters()
        previous_gpas = load_previous_gpas_from_processed_files(output_dir, semester_key, timestamp)
        cgpa_data = load_all_previous_gpas_for_cgpa(output_dir, semester_key, timestamp)
        
        for idx, row in mastersheet_df.iterrows():
            exam_no = str(row['EXAM NUMBER']).strip()
            total_grade_points = 0.0
            total_units = 0
            total_units_passed = 0
            total_units_failed = 0
            failed_courses = []
            
            for code in course_codes:
                score = row.get(code)
                if pd.isna(score):
                    continue
                try:
                    score_val = float(score)
                    cu = semester_credit_units[semester_key].get(code, 0)
                    total_grade_points += get_grade_point(score_val) * cu
                    total_units += cu
                    if score_val >= pass_threshold:
                        total_units_passed += cu
                    else:
                        total_units_failed += cu
                        failed_courses.append(code)
                except (ValueError, TypeError):
                    continue
            
            current_gpa = round(total_grade_points / total_units, 2) if total_units > 0 else 0.0
            cgpa = calculate_cgpa(cgpa_data.get(exam_no, {}), current_gpa, total_units)
            
            mastersheet_df.at[idx, 'CU Passed'] = total_units_passed
            mastersheet_df.at[idx, 'CU Failed'] = total_units_failed
            mastersheet_df.at[idx, 'TCPE'] = round(total_grade_points, 1)
            mastersheet_df.at[idx, 'GPA'] = current_gpa
            mastersheet_df.at[idx, 'CGPA'] = cgpa
            mastersheet_df.at[idx, 'REMARKS'] = determine_student_status(
                {'GPA': current_gpa, 'CU Failed': total_units_failed}, total_units, pass_threshold)
            
            if mastersheet_df.at[idx, 'REMARKS'] == "Withdrawn" and params['track_withdrawn']:
                WITHDRAWN_STUDENTS[exam_no] = {
                    'withdrawn_semester': semester_key,
                    'reappeared_semesters': [],
                    'withdrawn_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
        
        # Identify carryover students (Fix 2: Ensured after regular processing)
        carryover_students = identify_carryover_students(mastersheet_df, semester_key, set_name, pass_threshold)
        save_carryover_records(carryover_students, output_dir, set_name, semester_key)
        
        # Save mastersheet
        for idx, row in mastersheet_df.iterrows():
            row_idx = idx + 6
            for col_name in mastersheet_df.columns:
                col_idx = mastersheet_df.columns.get_loc(col_name) + 1
                ws.cell(row=row_idx, column=col_idx).value = row[col_name]
        
        wb.save(mastersheet_path)
        print(f"✅ Mastersheet saved: {mastersheet_path}")
        
        # Generate PDFs
        if params['generate_pdf']:
            pdf_path = os.path.join(output_dir, f"student_reports_{set_name}_{semester_key}_{timestamp}.pdf")
            generate_individual_student_pdf(
                mastersheet_df, pdf_path, semester_key, logo_path=logo_path,
                filtered_credit_units=semester_credit_units.get(semester_key, {}),
                ordered_codes=course_codes,
                course_titles_map=semester_course_titles.get(semester_key, {}),
                pass_threshold=pass_threshold,
                previous_gpas=previous_gpas,
                cgpa_data=cgpa_data
            )

        # Fix 3: Check for resit file in CARRYOVER subdirectory
        carryover_subdir = os.path.join(raw_dir, "CARRYOVER")
        if os.path.exists(carryover_subdir):
            resit_files = [f for f in os.listdir(carryover_subdir) 
                           if f.lower().endswith(('.xlsx', '.xls')) and semester_key in f.upper()]
            
            if resit_files:
                print(f"🔍 Found {len(resit_files)} resit files for {semester_key}")
                for resit_file in resit_files:
                    resit_path = os.path.join(carryover_subdir, resit_file)
                    print(f"📄 Processing resit: {resit_path}")
                    success, updated_scores, updated_students = process_resit_file(
                        resit_path, mastersheet_path, semester_key, set_name, 
                        pass_threshold, semester_course_maps, semester_credit_units, output_dir, semester_course_titles
                    )
                    if success:
                        print(f"✅ Updated {updated_scores} scores for {len(updated_students)} students")
                    else:
                        print(f"❌ Failed to process resit file: {resit_file}")
        
        # Enhanced resit processing with debugging
        if process_resit and resit_file_path and os.path.exists(resit_file_path):
            print(f"\n🎯 PROCESSING RESIT FOR {semester_key}")
            
            # First analyze the resit file content
            analyze_resit_file_content(resit_file_path, semester_key)
            
            # Then create a course mapping analysis
            mapping_file = create_course_mapping_file(resit_file_path, mastersheet_path, semester_key, output_dir)
            
            # Check if there are any carryover students for this semester
            semester_carryovers = {k: v for k, v in CARRYOVER_STUDENTS.items() if semester_key in k}
            print(f"📋 Carryover students in {semester_key}: {len(semester_carryovers)}")
            
            if len(semester_carryovers) == 0:
                print(f"❌ No carryover students found for this semester!")
                print(f"💡 Solution: Process the semester where the students actually failed first")
                print(f"💡 Or check if students have passing scores in all courses")
            else:
                print(f"✅ Found {len(semester_carryovers)} carryover students")
                for student_key, data in list(semester_carryovers.items())[:3]:  # Show first 3
                    print(f"   - {data['exam_number']}: {[c['course_code'] for c in data['failed_courses']]}")
            
            # Then process with enhanced debugging
            success, updated_scores, updated_students = process_resit_file(
                resit_file_path, mastersheet_path, semester_key, set_name, pass_threshold,
                semester_course_maps, semester_credit_units, output_dir, semester_course_titles
            )
            
            if success:
                if updated_scores > 0:
                    print(f"✅ Carryover processing completed! Updated {updated_scores} scores for {updated_students} students")
                    # FIX 5: Regenerate PDFs with updated mastersheet_df
                    pdf_path = os.path.join(output_dir, f"mastersheet_students_resit_{timestamp}_{set_name}_{semester_key}.pdf")
                    generate_individual_student_pdf(
                        mastersheet_df, pdf_path, semester_key, logo_path=logo_path,
                        filtered_credit_units=semester_credit_units.get(semester_key, {}),
                        ordered_codes=course_codes,
                        course_titles_map=semester_course_titles.get(semester_key, {}),
                        pass_threshold=pass_threshold,
                        previous_gpas=previous_gpas,
                        cgpa_data=cgpa_data
                    )
                    
                    # FIX 6: Generate new carryover mastersheet
                    generate_resit_carryover_mastersheet(mastersheet_df, output_dir, timestamp, semester_key, set_name, pass_threshold, 
                                                         semester_credit_units.get(semester_key, {}), course_codes, semester_course_titles.get(semester_key, {}))
                else:
                    print(f"ℹ️ No scores updated. Check the detailed log for reasons.")
                    
                    # Provide troubleshooting advice
                    print(f"\n🔧 TROUBLESHOOTING SUGGESTIONS:")
                    print(f"1. Check course mapping template: {mapping_file}")
                    print(f"2. Verify exam number formats match between files")
                    print(f"3. Check if resit scores are actually passing (≥ {pass_threshold})")
                    print(f"4. Verify original scores were actually failing (< {pass_threshold})")
                    print(f"5. Check data format (transposed vs wide)")
                    print(f"6. Check if resit file contains courses for the right semester")
            else:
                print(f"❌ Resit processing failed. Check logs above.")
    
    except Exception as e:
        print(f"❌ Error processing semester {semester_key}: {e}")
        traceback.print_exc()

# ----------------------------
# Non-Interactive Mode
# ----------------------------

def process_in_non_interactive_mode(params, base_dir):
    """Process results in non-interactive mode."""
    try:
        semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_course_data()
    except Exception as e:
        print(f"❌ Could not load course data: {e}")
        return False

    selected_set = params['selected_set']
    selected_semesters = params['selected_semesters']
    process_resit = params['process_resit']
    resit_file_path = params['resit_file_path']
    resit_output_dir = params['resit_output_dir']
    pass_threshold = params['pass_threshold']

    sets_to_process = [selected_set] if selected_set != 'all' else get_available_sets(base_dir)
    
    for nd_set in sets_to_process:
        raw_dir = os.path.join(base_dir, "ND", nd_set, "RAW_RESULTS")
        clean_dir = os.path.join(base_dir, "ND", nd_set, "CLEAN_RESULTS")
        os.makedirs(clean_dir, exist_ok=True)

        # Fix 4: Skip regular processing in resit mode
        if process_resit:
            print(f"🔄 Resit mode for set {nd_set}")
            timestamp = datetime.now().strftime(TIMESTAMP_FMT)
            # FIX 2: Use environment output_dir and align timestamp
            output_dir = os.environ.get('RESIT_OUTPUT_DIR')
            if not output_dir:
                set_output_dir = os.path.join(clean_dir, f"{nd_set}_CARRYOVER-{timestamp}")  # Use CARRYOVER- instead of RESIT-
            else:
                set_output_dir = output_dir
            os.makedirs(set_output_dir, exist_ok=True)
            print(f"✅ Using resit output dir: {set_output_dir}")
            
            semesters_to_process = selected_semesters if selected_semesters else SEMESTER_ORDER
            for semester_key in semesters_to_process:
                # Find latest regular mastersheet
                timestamp_folders = [f for f in os.listdir(clean_dir) 
                                    if f.startswith(f"{nd_set}_RESULT-") and os.path.isdir(os.path.join(clean_dir, f))]
                if timestamp_folders:
                    latest_folder = sorted(timestamp_folders)[-1]
                    regular_output_dir = os.path.join(clean_dir, latest_folder)
                    mastersheet_files = sorted(glob.glob(os.path.join(regular_output_dir, "mastersheet_*.xlsx")))
                    if mastersheet_files:
                        mastersheet_path = mastersheet_files[-1]
                        print(f"✅ Using existing mastersheet for resit: {mastersheet_path}")
                        
                        # Get resit files from CARRYOVER subdir
                        carryover_subdir = os.path.join(raw_dir, "CARRYOVER")
                        normalized_key = semester_key.replace('ND-', '').upper()
                        resit_files = [f for f in os.listdir(carryover_subdir) 
                                       if f.lower().endswith(('.xlsx', '.xls')) 
                                       and normalized_key in f.upper().replace('ND-', '')]
                        
                        if resit_files:
                            print(f"🔍 Found {len(resit_files)} resit files")
                            for resit_file in resit_files:
                                resit_path = os.path.join(carryover_subdir, resit_file)
                                print(f"📄 Processing resit: {resit_path}")
                                success, updated_scores, num_updated_students = process_resit_file(
                                    resit_path, mastersheet_path, semester_key, nd_set, 
                                    pass_threshold, semester_course_maps, semester_credit_units, set_output_dir, semester_course_titles
                                )
                                if success:
                                    # FIX: updated_students is now an integer (count), not a collection
                                    print(f"✅ Updated {updated_scores} scores for {num_updated_students} students")
                                else:
                                    print(f"❌ Failed to process resit file: {resit_file}")
                        else:
                            print("❌ No resit files found matching semester")
                    else:
                        print("❌ No mastersheet found in latest regular folder")
                else:
                    print("❌ No regular result folders found - Run regular processing first")
                    
            # FIX 4: Adjust ZIP creation for resit mode
            if os.path.exists(set_output_dir):
                # FIX 4: Use CARRYOVER prefix for ZIP in resit mode
                zip_filename = f"CARRYOVER_{nd_set}_{timestamp}.zip"
                zip_path = os.path.join(clean_dir, zip_filename)
                if create_zip_folder(set_output_dir, zip_path):
                    print(f"✅ ZIP file created: {zip_path}")
        else:
            # Normal regular processing
            timestamp = datetime.now().strftime(TIMESTAMP_FMT)
            set_output_dir = os.path.join(clean_dir, f"{nd_set}_RESULT-{timestamp}")
            os.makedirs(set_output_dir, exist_ok=True)

            # Fix 4: Better environment variable handling (adapted for non-interactive)
            if params.get('process_resit') or params.get('process_carryover'):
                resit_file_path = params.get('resit_file_path') or params.get('carryover_file_path')
                if resit_file_path and os.path.exists(resit_file_path):
                    print(f"✅ Resit file found: {resit_file_path}")
                    # Ensure it's copied to the correct CARRYOVER subdirectory
                    carryover_subdir = os.path.join(raw_dir, "CARRYOVER")
                    os.makedirs(carryover_subdir, exist_ok=True)
                    resit_filename = os.path.basename(resit_file_path)
                    target_path = os.path.join(carryover_subdir, resit_filename)
                    if not os.path.exists(target_path):
                        shutil.copy(resit_file_path, target_path)
                        print(f"✅ Copied resit file to {target_path}")
                else:
                    print(f"❌ Resit file not found or not specified")

            raw_files = [f for f in os.listdir(raw_dir) if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
            semesters_to_process = selected_semesters if selected_semesters else SEMESTER_ORDER

            for semester_key in semesters_to_process:
                if semester_key not in SEMESTER_ORDER:
                    print(f"⚠️ Invalid semester: {semester_key}")
                    continue
                
                process_semester_files_enhanced(
                    semester_key, raw_files, raw_dir, set_output_dir, timestamp, pass_threshold,
                    semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles, DEFAULT_LOGO_PATH, nd_set, process_resit, resit_file_path
                )
            
            # FIX 4: Regular ZIP naming
            zip_path = os.path.join(clean_dir, f"{nd_set}_RESULT-{timestamp}.zip")
            if create_zip_folder(set_output_dir, zip_path):
                print(f"✅ ZIP file created: {zip_path}")
    
    return True

# ----------------------------
# Main Runner
# ----------------------------

def main():
    """Main entry point for carryover processor."""
    print("🎯 ENHANCED CARRYOVER PROCESSOR WITH DEBUGGING")
    print("=" * 50)
    
    initialize_student_tracker()
    initialize_carryover_tracker()
    
    base_dir_norm = normalize_path(BASE_DIR)
    os.makedirs(base_dir_norm, exist_ok=True)
    os.makedirs(ND_COURSES_DIR, exist_ok=True)
    
    params = get_form_parameters()
    global DEFAULT_PASS_THRESHOLD
    DEFAULT_PASS_THRESHOLD = params['pass_threshold']

    # Fix 4: Better environment variable handling
    process_resit = params.get('process_resit', False)
    resit_file_path = params.get('resit_file_path', '')

    if is_web_mode() or any([params['selected_set'], params['selected_semesters']]):
        print("🔧 Running in NON-INTERACTIVE mode (Web)")
        success = process_in_non_interactive_mode(params, base_dir_norm)
        if success:
            print("✅ ND Examination Results Processing completed successfully")
            sys.exit(0)
        else:
            print("❌ ND Examination Results Processing failed")
            sys.exit(1)
    else:
        print("🔧 Running in INTERACTIVE mode (CLI)")
        try:
            semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_course_data()
        except Exception as e:
            print(f"❌ Could not load course data: {e}")
            sys.exit(1)

        available_sets = get_available_sets(base_dir_norm)
        if not available_sets:
            print(f"No ND-* directories found in {base_dir_norm}. Nothing to process.")
            sys.exit(1)

        sets_to_process = get_user_set_choice(available_sets)
        timestamp = datetime.now().strftime(TIMESTAMP_FMT)
        
        for nd_set in sets_to_process:
            raw_dir = os.path.join(base_dir_norm, "ND", nd_set, "RAW_RESULTS")
            clean_dir = os.path.join(base_dir_norm, "ND", nd_set, "CLEAN_RESULTS")
            os.makedirs(clean_dir, exist_ok=True)
            
            # Fix 4: Skip regular processing in resit mode
            if process_resit:
                print(f"🔄 Resit mode for set {nd_set}")
                timestamp = datetime.now().strftime(TIMESTAMP_FMT)
                # FIX 2: Use environment output_dir and align timestamp
                output_dir = os.environ.get('RESIT_OUTPUT_DIR')
                if not output_dir:
                    set_output_dir = os.path.join(clean_dir, f"{nd_set}_CARRYOVER-{timestamp}")  # Use CARRYOVER- instead of RESIT-
                else:
                    set_output_dir = output_dir
                os.makedirs(set_output_dir, exist_ok=True)
                print(f"✅ Using resit output dir: {set_output_dir}")
                
                semesters_to_process = get_user_semester_choice()
                for semester_key in semesters_to_process:
                    # Find latest regular mastersheet
                    timestamp_folders = [f for f in os.listdir(clean_dir) 
                                        if f.startswith(f"{nd_set}_RESULT-") and os.path.isdir(os.path.join(clean_dir, f))]
                    if timestamp_folders:
                        latest_folder = sorted(timestamp_folders)[-1]
                        regular_output_dir = os.path.join(clean_dir, latest_folder)
                        mastersheet_files = sorted(glob.glob(os.path.join(regular_output_dir, "mastersheet_*.xlsx")))
                        if mastersheet_files:
                            mastersheet_path = mastersheet_files[-1]
                            print(f"✅ Using existing mastersheet for resit: {mastersheet_path}")
                            
                            # Get resit files from CARRYOVER subdir
                            carryover_subdir = os.path.join(raw_dir, "CARRYOVER")
                            normalized_key = semester_key.replace('ND-', '').upper()
                            resit_files = [f for f in os.listdir(carryover_subdir) 
                                           if f.lower().endswith(('.xlsx', '.xls')) 
                                           and normalized_key in f.upper().replace('ND-', '')]
                            
                            if resit_files:
                                print(f"🔍 Found {len(resit_files)} resit files")
                                for resit_file in resit_files:
                                    resit_path = os.path.join(carryover_subdir, resit_file)
                                    print(f"📄 Processing resit: {resit_path}")
                                    success, updated_scores, num_updated_students = process_resit_file(
                                        resit_path, mastersheet_path, semester_key, nd_set, 
                                        params['pass_threshold'], semester_course_maps, semester_credit_units, set_output_dir, semester_course_titles
                                    )
                                    if success:
                                        # FIX: updated_students is now an integer (count), not a collection
                                        print(f"✅ Updated {updated_scores} scores for {num_updated_students} students")
                                    else:
                                        print(f"❌ Failed to process resit file: {resit_file}")
                            else:
                                print("❌ No resit files found matching semester")
                        else:
                            print("❌ No mastersheet found in latest regular folder")
                    else:
                        print("❌ No regular result folders found - Run regular processing first")
                        
                # FIX 4: Adjust ZIP creation for resit mode
                if os.path.exists(set_output_dir):
                    # FIX 4: Use CARRYOVER prefix for ZIP in resit mode
                    zip_filename = f"CARRYOVER_{nd_set}_{timestamp}.zip"
                    zip_path = os.path.join(clean_dir, zip_filename)
                    if create_zip_folder(set_output_dir, zip_path):
                        print(f"✅ ZIP file created: {zip_path}")
            else:
                # Normal regular processing
                set_output_dir = os.path.join(clean_dir, f"{nd_set}_RESULT-{timestamp}")
                os.makedirs(set_output_dir, exist_ok=True)

                # Fix 4: Copy resit file to CARRYOVER subdir if specified
                if process_resit:
                    if resit_file_path and os.path.exists(resit_file_path):
                        carryover_subdir = os.path.join(raw_dir, "CARRYOVER")
                        os.makedirs(carryover_subdir, exist_ok=True)
                        resit_filename = os.path.basename(resit_file_path)
                        target_path = os.path.join(carryover_subdir, resit_filename)
                        if not os.path.exists(target_path):
                            shutil.copy(resit_file_path, target_path)
                            print(f"✅ Copied resit file to {target_path}")
                
                raw_files = [f for f in os.listdir(raw_dir) if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
                if not raw_files:
                    print(f"⚠️ No raw files in {raw_dir}; skipping {nd_set}")
                    continue
                
                semesters_to_process = get_user_semester_choice()
                
                for semester_key in semesters_to_process:
                    process_semester_files_enhanced(
                        semester_key, raw_files, raw_dir, set_output_dir, timestamp, params['pass_threshold'],
                        semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles,
                        DEFAULT_LOGO_PATH, nd_set, process_resit, params['resit_file_path']
                    )
                
                # FIX 4: Regular ZIP naming
                zip_path = os.path.join(clean_dir, f"{nd_set}_RESULT-{timestamp}.zip")
                if create_zip_folder(set_output_dir, zip_path):
                    print(f"✅ ZIP file created: {zip_path}")

        print(f"\n📊 STUDENT TRACKING SUMMARY:")
        print(f"Total unique students tracked: {len(STUDENT_TRACKER)}")
        print(f"Total withdrawn students: {len(WITHDRAWN_STUDENTS)}")
        
        if CARRYOVER_STUDENTS:
            print(f"\n📋 CARRYOVER STUDENT SUMMARY:")
            print(f"Total carryover students: {len(CARRYOVER_STUDENTS)}")
            semester_counts = {}
            for student_key, data in CARRYOVER_STUDENTS.items():
                semester = data['semester']
                semester_counts[semester] = semester_counts.get(semester, 0) + 1
            for semester, count in semester_counts.items():
                print(f"  {semester}: {count} students")

        reappeared_count = 0
        for exam_no, data in WITHDRAWN_STUDENTS.items():
            if data['reappeared_semesters']:
                reappeared_count += 1
                print(f"🚨 {exam_no}: Withdrawn in {data['withdrawn_semester']}, reappeared in {data['reappeared_semesters']}")
        if reappeared_count > 0:
            print(f"🚨 ALERT: {reappeared_count} previously withdrawn students have reappeared in later semesters!")

if __name__ == "__main__":
    main()