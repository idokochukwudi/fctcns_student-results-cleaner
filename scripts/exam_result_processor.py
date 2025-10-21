#!/usr/bin/env python3
"""
exam_result_processor.py
Complete script with integrated carryover student management and resit processing.
FIXED VERSION - Proper course mapping and mastersheet formatting
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
    # Check Railway environment first
    railway_base = os.getenv('BASE_DIR')
    if railway_base and os.path.exists(railway_base):
        return railway_base
    
    # Check if we're running on Railway but BASE_DIR doesn't exist
    if is_running_on_railway():
        # Create the directory structure on Railway
        railway_base = '/app/EXAMS_INTERNAL'
        os.makedirs(railway_base, exist_ok=True)
        os.makedirs(os.path.join(railway_base, 'ND', 'ND-COURSES'), exist_ok=True)
        return railway_base
    
    # Local development fallback - updated to match your structure
    local_path = os.path.join(os.path.expanduser('~'), 'student_result_cleaner', 'EXAMS_INTERNAL')
    if os.path.exists(local_path):
        return local_path
    
    # Final fallback - current directory
    return os.path.join(os.path.dirname(__file__), 'EXAMS_INTERNAL')

BASE_DIR = get_base_directory()
# UPDATED: ND directories now under ND folder
ND_BASE_DIR = os.path.join(BASE_DIR, "ND")
ND_COURSES_DIR = os.path.join(ND_BASE_DIR, "ND-COURSES")

# Ensure directories exist
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(ND_BASE_DIR, exist_ok=True)
os.makedirs(ND_COURSES_DIR, exist_ok=True)

# Global variables for threshold upgrade
THRESHOLD_UPGRADED = False
ORIGINAL_THRESHOLD = 50.0
UPGRADE_MIN = None
UPGRADE_MAX = 49

# Global carryover tracker
CARRYOVER_STUDENTS = {}
RESIT_HISTORY = {}

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
    process_resit = os.getenv('PROCESS_RESIT', 'False').lower() == 'true'
    resit_file_path = os.getenv('RESIT_FILE_PATH', '')
    
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
    
    print(f"🎯 FORM PARAMETERS:")
    print(f"   Selected Set: {selected_set}")
    print(f"   Processing Mode: {processing_mode}")
    print(f"   Selected Semesters: {selected_semesters}")
    print(f"   Pass Threshold: {pass_threshold}")
    print(f"   Generate PDF: {generate_pdf}")
    print(f"   Track Withdrawn: {track_withdrawn}")
    print(f"   Process Resit: {process_resit}")
    print(f"   Resit File Path: {resit_file_path}")
    
    return {
        'selected_set': selected_set,
        'processing_mode': processing_mode,
        'selected_semesters': selected_semesters,
        'pass_threshold': pass_threshold,
        'generate_pdf': generate_pdf,
        'track_withdrawn': track_withdrawn,
        'process_resit': process_resit,
        'resit_file_path': resit_file_path
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

DEFAULT_LOGO_PATH = os.path.normpath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "launcher",
        "static",
        "logo.png"))

NAME_WIDTH_CAP = 40

# Define semester processing order
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
# Carryover Management Functions
# ----------------------------

def initialize_carryover_tracker():
    """Initialize the global carryover tracker."""
    global CARRYOVER_STUDENTS, RESIT_HISTORY
    CARRYOVER_STUDENTS = {}
    RESIT_HISTORY = {}

def identify_carryover_students(mastersheet_df, semester_key, set_name, pass_threshold=50.0):
    """
    Identify students with carryover courses from current semester processing.
    Returns list of carryover students with their failed courses.
    """
    carryover_students = []
    
    # Get course columns (excluding student info columns)
    course_columns = [col for col in mastersheet_df.columns 
                     if col not in ['S/N', 'EXAMS NUMBER', 'NAME', 'REMARKS', 
                                   'CU Passed', 'CU Failed', 'TCPE', 'GPA', 'AVERAGE']]
    
    for idx, student in mastersheet_df.iterrows():
        failed_courses = []
        exam_no = str(student['EXAMS NUMBER'])
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
    
    # Prepare data for Excel
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
    
    # Save as JSON for easy processing
    json_file = os.path.join(carryover_dir, f"{filename}.json")
    with open(json_file, 'w') as f:
        json.dump(carryover_students, f, indent=2)
    
    print(f"📁 Carryover records saved in: {carryover_dir}")
    return carryover_dir

def load_carryover_records(output_dir, set_name, semester_key):
    """
    Load carryover records from the clean results folder.
    """
    carryover_dir = os.path.join(output_dir, "CARRYOVER_RECORDS")
    if not os.path.exists(carryover_dir):
        return None
    
    # Look for the most recent carryover file for this set and semester
    pattern = f"co_student_{set_name}_{semester_key}_*.json"
    matching_files = []
    
    for file in os.listdir(carryover_dir):
        if file.startswith(f"co_student_{set_name}_{semester_key}_") and file.endswith(".json"):
            matching_files.append(file)
    
    if not matching_files:
        return None
    
    # Get the most recent file
    latest_file = sorted(matching_files)[-1]
    json_file = os.path.join(carryover_dir, latest_file)
    
    try:
        with open(json_file, 'r') as f:
            carryover_data = json.load(f)
        print(f"✅ Loaded carryover records: {json_file}")
        return carryover_data
    except Exception as e:
        print(f"❌ Error loading carryover records: {e}")
        return None

def process_resit_results(resit_file_path, output_dir, set_name, semester_key, pass_threshold=50.0):
    """
    Process resit results and update student records.
    Returns updated mastersheet DataFrame.
    """
    print(f"🔄 Processing resit results for {set_name} - {semester_key}")
    
    # Load existing carryover records
    carryover_records = load_carryover_records(output_dir, set_name, semester_key)
    if not carryover_records:
        print(f"❌ No carryover records found for {set_name} - {semester_key}")
        return None
    
    # Load resit results
    try:
        resit_df = pd.read_excel(resit_file_path)
        print(f"✅ Loaded resit file: {resit_file_path}")
        print(f"📊 Resit file columns: {resit_df.columns.tolist()}")
    except Exception as e:
        print(f"❌ Error loading resit file: {e}")
        return None
    
    # Load the current mastersheet
    mastersheet_path = os.path.join(output_dir, f"mastersheet_*.xlsx")
    mastersheet_files = glob.glob(mastersheet_path)
    
    if not mastersheet_files:
        print(f"❌ No mastersheet found in {output_dir}")
        return None
    
    latest_mastersheet = sorted(mastersheet_files)[-1]
    try:
        mastersheet_df = pd.read_excel(latest_mastersheet, sheet_name=semester_key)
        print(f"✅ Loaded mastersheet: {latest_mastersheet}")
    except Exception as e:
        print(f"❌ Error loading mastersheet: {e}")
        return None
    
    # Process resit results and update records
    updated_students = []
    update_count = 0
    
    for student in carryover_records:
        exam_no = student['exam_number']
        student_updated = False
        
        # Find student in resit results
        student_resits = resit_df[resit_df['EXAMS NUMBER'] == exam_no]
        
        if not student_resits.empty:
            for course in student['failed_courses']:
                course_code = course['course_code']
                
                # Check if course exists in resit results
                if course_code in student_resits.columns:
                    resit_score = student_resits[course_code].iloc[0]
                    
                    if pd.notna(resit_score):
                        try:
                            resit_score_val = float(resit_score)
                            
                            # Update course record
                            course['resit_attempts'] += 1
                            course['best_score'] = max(course['best_score'], resit_score_val)
                            
                            if resit_score_val >= pass_threshold:
                                course['status'] = 'Passed_Resit'
                                print(f"✅ {exam_no} passed {course_code} with {resit_score_val}")
                            else:
                                course['status'] = 'Failed_Resit'
                                print(f"❌ {exam_no} failed {course_code} resit with {resit_score_val}")
                            
                            # Update mastersheet
                            student_mask = mastersheet_df['EXAMS NUMBER'] == exam_no
                            if student_mask.any():
                                mastersheet_df.loc[student_mask, course_code] = resit_score_val
                                update_count += 1
                            
                            student_updated = True
                            student['total_resit_attempts'] += 1
                            
                        except (ValueError, TypeError) as e:
                            print(f"⚠️ Invalid score for {exam_no} - {course_code}: {resit_score}")
                            continue
        
        if student_updated:
            updated_students.append(student)
    
    if updated_students:
        # Save updated carryover records
        updated_carryover_dir = save_carryover_records(updated_students, output_dir, set_name, semester_key)
        
        # Recompute GPA and metrics for updated students
        mastersheet_df = recompute_student_metrics(mastersheet_df, semester_key, set_name, pass_threshold)
        
        # Save updated mastersheet
        updated_mastersheet_path = latest_mastersheet.replace('.xlsx', '_RESIT_UPDATED.xlsx')
        try:
            with pd.ExcelWriter(updated_mastersheet_path, engine='openpyxl') as writer:
                mastersheet_df.to_excel(writer, sheet_name=semester_key, index=False)
            print(f"✅ Updated mastersheet saved: {updated_mastersheet_path}")
        except Exception as e:
            print(f"❌ Error saving updated mastersheet: {e}")
        
        print(f"✅ Updated {update_count} scores for {len(updated_students)} students")
        
        # Update resit history
        update_resit_history(updated_students, set_name, semester_key)
        
        return mastersheet_df
    else:
        print("⚠️ No students updated from resit results")
        return None

def recompute_student_metrics(mastersheet_df, semester_key, set_name, pass_threshold):
    """
    Recompute GPA and other metrics after resit updates.
    """
    print("🔄 Recomputing student metrics after resit updates...")
    
    # Get course data for this semester
    try:
        semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_course_data()
        course_map = semester_course_maps[semester_key]
        credit_units = semester_credit_units[semester_key]
        
        ordered_titles = list(course_map.keys())
        ordered_codes = [course_map[t] for t in ordered_titles if course_map.get(t)]
        ordered_codes = [c for c in ordered_codes if credit_units.get(c, 0) > 0]
        filtered_credit_units = {c: credit_units[c] for c in ordered_codes}
        total_cu = sum(filtered_credit_units.values())
        
    except Exception as e:
        print(f"⚠️ Error loading course data for recomputation: {e}")
        return mastersheet_df
    
    # Recompute TCPE, TCUP, TCUF, GPA for each student
    for idx, student in mastersheet_df.iterrows():
        tcpe = 0.0
        tcup = 0
        tcuf = 0
        
        for code in ordered_codes:
            if code in mastersheet_df.columns:
                score = student.get(code, 0)
                try:
                    score_val = float(score) if pd.notna(score) else 0
                    cu = filtered_credit_units.get(code, 0)
                    gp = get_grade_point(score_val)
                    
                    # TCPE: Grade Point × Credit Units
                    tcpe += gp * cu
                    
                    # TCUP/TCUF: Count credit units based on pass/fail
                    if score_val >= pass_threshold:
                        tcup += cu
                    else:
                        tcuf += cu
                        
                except (ValueError, TypeError):
                    continue
        
        # Update student metrics
        mastersheet_df.at[idx, "TCPE"] = round(tcpe, 1)
        mastersheet_df.at[idx, "CU Passed"] = tcup
        mastersheet_df.at[idx, "CU Failed"] = tcuf
        
        # Calculate GPA
        gpa = round((tcpe / total_cu), 2) if total_cu > 0 else 0.0
        mastersheet_df.at[idx, "GPA"] = gpa
        
        # Update average
        valid_scores = [student.get(code, 0) for code in ordered_codes if code in mastersheet_df.columns]
        valid_scores = [s for s in valid_scores if pd.notna(s)]
        if valid_scores:
            mastersheet_df.at[idx, "AVERAGE"] = round(sum(valid_scores) / len(valid_scores), 0)
        
        # Update remarks
        fails = [code for code in ordered_codes if float(student.get(code, 0) or 0) < pass_threshold]
        if not fails:
            mastersheet_df.at[idx, "REMARKS"] = "Passed"
        else:
            failed_courses_str = ", ".join(sorted(fails))
            mastersheet_df.at[idx, "REMARKS"] = f"Failed: {failed_courses_str}"
    
    # Add resit count column if not exists
    if "RESIT COUNT" not in mastersheet_df.columns:
        mastersheet_df["RESIT COUNT"] = 0
    
    # Update resit counts based on carryover records
    carryover_records = load_carryover_records(os.path.dirname(mastersheet_df), set_name, semester_key)
    if carryover_records:
        for student in carryover_records:
            exam_no = student['exam_number']
            total_resits = student['total_resit_attempts']
            
            student_mask = mastersheet_df['EXAMS NUMBER'] == exam_no
            if student_mask.any():
                mastersheet_df.loc[student_mask, "RESIT COUNT"] = total_resits
    
    return mastersheet_df

def update_resit_history(updated_students, set_name, semester_key):
    """
    Update global resit history tracker.
    """
    global RESIT_HISTORY
    
    for student in updated_students:
        exam_no = student['exam_number']
        student_key = f"{exam_no}_{set_name}_{semester_key}"
        
        if student_key not in RESIT_HISTORY:
            RESIT_HISTORY[student_key] = []
        
        resit_record = {
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'courses_attempted': [c['course_code'] for c in student['failed_courses'] if c['resit_attempts'] > 0],
            'total_attempts': student['total_resit_attempts'],
            'results': {c['course_code']: c['best_score'] for c in student['failed_courses'] if c['resit_attempts'] > 0}
        }
        
        RESIT_HISTORY[student_key].append(resit_record)

def get_resit_summary(set_name, semester_key):
    """
    Get summary of resit activity for a set and semester.
    """
    summary = {
        'total_students_with_resits': 0,
        'total_resit_attempts': 0,
        'courses_with_resits': {},
        'successful_resits': 0,
        'failed_resits': 0
    }
    
    for student_key, history in RESIT_HISTORY.items():
        if set_name in student_key and semester_key in student_key:
            summary['total_students_with_resits'] += 1
            for record in history:
                summary['total_resit_attempts'] += record['total_attempts']
                
                for course, score in record['results'].items():
                    if course not in summary['courses_with_resits']:
                        summary['courses_with_resits'][course] = 0
                    summary['courses_with_resits'][course] += 1
                    
                    if score >= DEFAULT_PASS_THRESHOLD:
                        summary['successful_resits'] += 1
                    else:
                        summary['failed_resits'] += 1
    
    return summary

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
                exam_no = mastersheet.at[idx, "EXAMS NUMBER"]
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

def normalize_course_name(name):
    """Simple normalization for course title matching."""
    return re.sub(
        r'\s+',
        ' ',
        str(name).strip().lower()).replace(
        'coomunication',
        'communication')

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

    for idx, row in mastersheet.iterrows():
        exam_no = str(row["EXAMS NUMBER"]).strip()
        if is_student_withdrawn(exam_no):
            withdrawal_history = get_withdrawal_history(exam_no)
            # Only remove if student was withdrawn in a PREVIOUS semester
            if withdrawal_history and withdrawal_history['withdrawn_semester'] != semester_key:
                removed_students.append(exam_no)
                filtered_mastersheet = filtered_mastersheet[filtered_mastersheet["EXAMS NUMBER"] != exam_no]

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
        exam_no = str(r.get("EXAMS NUMBER", r.get("REG. No", "")))
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
                    print(f"🔼 PDF: Upgraded score for {exam_no} - {code}: → 50")
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
        exam_no = str(r.get("EXAMS NUMBER", "")).strip()
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
# Main file processing (Enhanced with Carryover)
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
    Enhanced with resit count tracking.
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
    ordered_codes = [course_map[t]
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
                                            "Exam No",
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

        for col in [c for c in df.columns if c not in ["REG. No", "NAME"]]:
            norm = normalize_course_name(col)
            matched_code = None
            for title, code in zip(
                ordered_titles, [
                    course_map[t] for t in ordered_titles]):
                if normalize_course_name(title) == norm:
                    matched_code = code
                    break
            if matched_code:
                newcol = f"{matched_code}_{s.upper()}"
                df.rename(columns={col: newcol}, inplace=True)
                print(f"✅ Matched column '{col}' to course code '{matched_code}'")

        cur_cols = ["REG. No", "NAME"] + \
            [c for c in df.columns if c.endswith(f"_{s.upper()}")]
        cur = df[cur_cols].copy()
        
        # Debug: Show data before merging
        print(f"📊 Data in {s} sheet - Shape: {cur.shape}")
        if not cur.empty:
            print(f"🔍 First 2 rows of {s} data:")
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
    mastersheet.rename(columns={"REG. No": "EXAMS NUMBER"}, inplace=True)

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
            print(f"🔄 Applying upgrade rule from parameters: {upgrade_min_threshold}–49 → 50")
    
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
            exam_no = str(row["EXAMS NUMBER"]).strip()
            withdrawn_students.append(exam_no)
            mark_student_withdrawn(exam_no, semester_key)
            print(f"🚫 Student {exam_no} marked as withdrawn in {semester_key}")

    # Update student tracker with current semester's students (after filtering)
    exam_numbers = mastersheet["EXAMS NUMBER"].astype(str).str.strip().tolist()
    update_student_tracker(semester_key, exam_numbers, withdrawn_students)

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
    out_cols = ["S/N", "EXAMS NUMBER", "NAME"] + course_cols + \
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
    for t, c in zip(ordered_titles, [course_map[t] for t in ordered_titles]):
        if c in ordered_codes:
            display_course_titles.append(t)

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

    # FIXED: Freeze the column headers (S/N, EXAMS NUMBER, NAME, etc.) at row start_row + 3
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
            # Count the total characters in failed courses list
            failed_courses = remark_str.replace("Failed: ", "")
            # Estimate width needed for the course codes
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
        elif column_letter == "B" or headers[col_idx - 1] in ["EXAMS NUMBER", "EXAM NO"]:
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
    semester_files = []
    for rf in raw_files:
        detected_sem, _, _, _, _, _ = detect_semester_from_filename(rf)
        if detected_sem == semester_key:
            semester_files.append(rf)

    if not semester_files:
        print(f"⚠️ No files found for semester {semester_key}")
        return None

    print(
        f"📁 Found {len(semester_files)} files for {semester_key}: {semester_files}")

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
                
                # Identify and save carryover students after processing
                carryover_students = identify_carryover_students(
                    result, semester_key, set_name, pass_threshold
                )
                
                if carryover_students:
                    carryover_dir = save_carryover_records(
                        carryover_students, output_dir, set_name, semester_key
                    )
                    print(f"✅ Identified {len(carryover_students)} carryover students")
                    
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
            else:
                print(f"❌ Failed to process {rf}")

        except Exception as e:
            print(f"❌ Error processing {rf}: {e}")
            import traceback
            traceback.print_exc()
    
    return mastersheet_result

# ----------------------------
# Resit Processing Mode
# ----------------------------

def process_resit_mode(params, base_dir_norm):
    """
    Process resit results in non-interactive mode for web interface.
    """
    print("🔄 Running in RESIT PROCESSING mode")
    
    resit_file_path = params['resit_file_path']
    selected_set = params['selected_set']
    selected_semesters = params['selected_semesters']
    
    if not resit_file_path or not os.path.exists(resit_file_path):
        print(f"❌ Resit file not found: {resit_file_path}")
        return False
    
    print(f"📁 Processing resit file: {resit_file_path}")
    print(f"🎯 Target set: {selected_set}")
    print(f"📚 Target semesters: {selected_semesters}")
    
    success_count = 0
    
    for semester_key in selected_semesters:
        # Find the clean directory for this set and semester
        clean_dir = normalize_path(os.path.join(base_dir_norm, "ND", selected_set, "CLEAN_RESULTS"))
        
        if not os.path.exists(clean_dir):
            print(f"❌ Clean results directory not found: {clean_dir}")
            continue
        
        # Look for the most recent timestamped folder
        timestamp_folders = [f for f in os.listdir(clean_dir) 
                           if f.startswith(f"{selected_set}_RESULT-") and os.path.isdir(os.path.join(clean_dir, f))]
        
        if not timestamp_folders:
            print(f"❌ No result folders found in {clean_dir}")
            continue
        
        latest_folder = sorted(timestamp_folders)[-1]
        output_dir = os.path.join(clean_dir, latest_folder)
        
        print(f"🔍 Processing resit for {semester_key} in {output_dir}")
        
        # Process resit results
        updated_mastersheet = process_resit_results(
            resit_file_path, output_dir, selected_set, semester_key, params['pass_threshold']
        )
        
        if updated_mastersheet is not None:
            print(f"✅ Successfully processed resit for {semester_key}")
            success_count += 1
            
            # Generate updated PDFs
            try:
                safe_sem = re.sub(r'[^\w\-]', '_', semester_key)
                student_pdf_path = os.path.join(
                    output_dir,
                    f"mastersheet_students_RESIT_UPDATED_{safe_sem}.pdf")
                
                # Get course data for PDF generation
                semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_course_data()
                course_map = semester_course_maps[semester_key]
                credit_units = semester_credit_units[semester_key]
                
                ordered_titles = list(course_map.keys())
                ordered_codes = [course_map[t] for t in ordered_titles if course_map.get(t)]
                ordered_codes = [c for c in ordered_codes if credit_units.get(c, 0) > 0]
                filtered_credit_units = {c: credit_units[c] for c in ordered_codes}
                
                # Generate PDF with resit information
                generate_individual_student_pdf(
                    updated_mastersheet,
                    student_pdf_path,
                    semester_key,
                    logo_path=DEFAULT_LOGO_PATH,
                    filtered_credit_units=filtered_credit_units,
                    ordered_codes=ordered_codes,
                    course_titles_map=semester_course_titles[semester_key],
                    previous_gpas=None,
                    cgpa_data=None,
                    total_cu=sum(filtered_credit_units.values()),
                    pass_threshold=params['pass_threshold'],
                    upgrade_min_threshold=None
                )
                
                print(f"✅ Updated PDF generated: {student_pdf_path}")
                
            except Exception as e:
                print(f"⚠️ Error generating updated PDF: {e}")
        else:
            print(f"❌ Failed to process resit for {semester_key}")
    
    print(f"📊 RESIT PROCESSING SUMMARY: {success_count}/{len(selected_semesters)} semesters updated")
    return success_count > 0

# ----------------------------
# Non-interactive mode processing (Enhanced)
# ----------------------------

def create_zip_folder(source_dir, zip_path):
    """Create a ZIP file from a directory"""
    import zipfile
    print(f"📦 Creating ZIP file from: {source_dir}")
    print(f"📦 ZIP destination: {zip_path}")
    
    # Check if source directory exists and has files
    if not os.path.exists(source_dir):
        print(f"❌ Source directory does not exist: {source_dir}")
        return False
    
    files_in_dir = []
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            files_in_dir.append(os.path.join(root, file))
    
    if not files_in_dir:
        print(f"❌ No files found in source directory: {source_dir}")
        return False
    
    print(f"📁 Found {len(files_in_dir)} files to include in ZIP")
    
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in files_in_dir:
                # Calculate relative path for proper ZIP structure
                arcname = os.path.relpath(file_path, source_dir)
                zipf.write(file_path, arcname)
                print(f"📄 Added to ZIP: {arcname}")
        
        # Verify the ZIP was created and has content
        zip_size = os.path.getsize(zip_path)
        print(f"✅ ZIP created successfully: {zip_path} ({zip_size} bytes)")
        
        # List contents of ZIP for verification
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            file_list = zipf.namelist()
            print(f"📋 ZIP contains {len(file_list)} files")
            for f in file_list[:5]:  # Show first 5 files
                print(f"   - {f}")
            if len(file_list) > 5:
                print(f"   ... and {len(file_list) - 5} more files")
        
        return True
        
    except Exception as e:
        print(f"❌ Error creating ZIP file: {e}")
        import traceback
        traceback.print_exc()
        return False

def process_in_non_interactive_mode(params, base_dir_norm):
    """Process exams in non-interactive mode for web interface."""
    print("🔧 Running in NON-INTERACTIVE mode (web interface)")
    
    # Check if this is resit processing mode
    if params.get('process_resit', False):
        return process_resit_mode(params, base_dir_norm)
    
    # Use parameters from environment variables
    selected_set = params['selected_set']
    selected_semesters = params['selected_semesters']
    
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
        
        # Process selected semesters
        for semester_key in selected_semesters:
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
        
        # Create ZIP of the entire set results - FIXED: Ensure proper file creation
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

# ----------------------------
# Main runner (Enhanced)
# ----------------------------

def main():
    print("Starting ND Examination Results Processing with Carryover Management...")
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
            
            # Create ZIP of the entire set results - FIXED: Use the enhanced create_zip_folder function
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
        success = process_in_non_interactive_mode(params, base_dir_norm)
        if success:
            print("✅ ND Examination Results Processing completed successfully")
        else:
            print("❌ ND Examination Results Processing failed")
        return

if __name__ == "__main__":
    try:
        main()
        print("✅ ND Examination Results Processing completed successfully")
    except Exception as e:
        print(f"❌ Error during processing: {e}")
        import traceback
        traceback.print_exc()