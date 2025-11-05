#!/usr/bin/env python3
"""
exam_processor_bn.py - Enhanced BN Examination Processor
Complete script with flexible threshold upgrade rule for BN results.
Enhanced with transposed data transformation, carryover management,
CGPA tracking, analysis sheet, and comprehensive BN course matching.
Web-compatible version with file upload support.
Uniform logic borrowed from ND script.
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
import zipfile
import time
import traceback
import logging
# PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
# ----------------------------
# Logging Configuration
# ----------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)
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
        os.makedirs(os.path.join(railway_base, 'BN', 'BN-COURSES'), exist_ok=True)
        return railway_base
  
    # Local development fallback - updated to match new structure
    local_path = os.path.join(os.path.expanduser('~'), 'student_result_cleaner', 'EXAMS_INTERNAL')
    if os.path.exists(local_path):
        return local_path
  
    # Final fallback - current directory
    return os.path.join(os.path.dirname(__file__), 'EXAMS_INTERNAL')
BASE_DIR = get_base_directory()
# UPDATED: BN directories now under BN folder
BN_BASE_DIR = os.path.join(BASE_DIR, "BN")
BN_COURSES_DIR = os.path.join(BN_BASE_DIR, "BN-COURSES")
# Ensure directories exist
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(BN_BASE_DIR, exist_ok=True)
os.makedirs(BN_COURSES_DIR, exist_ok=True)
# Define semester processing order for BN
BN_SEMESTER_ORDER = [
    "N-FIRST-YEAR-FIRST-SEMESTER",
    "N-FIRST-YEAR-SECOND-SEMESTER",
    "N-SECOND-YEAR-FIRST-SEMESTER",
    "N-SECOND-YEAR-SECOND-SEMESTER",
    "N-THIRD-YEAR-FIRST-SEMESTER",
    "N-THIRD-YEAR-SECOND-SEMESTER"
]
# Global variables for threshold upgrade
THRESHOLD_UPGRADED = False
ORIGINAL_THRESHOLD = 50.0
UPGRADE_MIN = 0 # Defaults to 0 (disabled) - no upgrades unless explicitly set
UPGRADE_MAX = 49
# Global student tracker
STUDENT_TRACKER = {}
WITHDRAWN_STUDENTS = {}
CARRYOVER_STUDENTS = {} # New global carryover tracker
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
def validate_raw_file(file_path):
    """Validate raw Excel file before processing."""
    logger.info(f"🔍 Validating file: {file_path}")
  
    if not os.path.exists(file_path):
        logger.error(f"❌ File does not exist: {file_path}")
        return False, "File does not exist"
  
    if os.path.getsize(file_path) == 0:
        logger.error(f"❌ File is empty: {file_path}")
        return False, "File is empty"
  
    try:
        xl = pd.ExcelFile(file_path)
        if not xl.sheet_names:
            logger.error(f"❌ No sheets in Excel file: {file_path}")
            return False, "No sheets in Excel file"
      
        expected_sheets = ['CA', 'OBJ', 'EXAM']
        found_sheets = [s for s in expected_sheets if s in xl.sheet_names]
      
        if not found_sheets:
            logger.error(f"❌ No expected sheets found in {file_path}. Has: {xl.sheet_names}")
            return False, f"No expected sheets found. Has: {xl.sheet_names}"
      
        logger.info(f"✅ Valid file with sheets: {found_sheets}")
        return True, f"Valid file with sheets: {found_sheets}"
      
    except Exception as e:
        logger.error(f"❌ Cannot open file {file_path}: {e}")
        return False, f"Cannot open file: {e}"
def process_uploaded_file(uploaded_file_path, base_dir_norm):
    """
    Process uploaded file in web mode.
    This function handles the single uploaded file for web processing.
    """
    logger.info("🔧 Processing uploaded file in web mode")
  
    # Validate the uploaded file first
    is_valid, validation_msg = validate_raw_file(uploaded_file_path)
    if not is_valid:
        logger.error(f"❌ Uploaded file validation failed: {validation_msg}")
        return False
  
    # Extract set name from filename or use default
    filename = os.path.basename(uploaded_file_path)
    set_name = "BN-UPLOADED"
  
    # Create temporary directory structure
    temp_dir = tempfile.mkdtemp()
    raw_dir = os.path.join(temp_dir, set_name, "RAW_RESULTS")
    clean_dir = os.path.join(temp_dir, set_name, "CLEAN_RESULTS")
    os.makedirs(raw_dir, exist_ok=True)
    os.makedirs(clean_dir, exist_ok=True)
  
    # Copy uploaded file to raw directory
    dest_path = os.path.join(raw_dir, filename)
    shutil.copy2(uploaded_file_path, dest_path)
  
    # Get parameters from environment
    params = get_form_parameters()
    ts = datetime.now().strftime(TIMESTAMP_FMT)
  
    try:
        # Load course data
        semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_bn_course_data()
      
        # Process the single file
        raw_files = [filename]
      
        # Detect semester from filename
        semester_key = detect_bn_semester_from_filename(filename)
      
        logger.info(f"🎯 Detected semester: {semester_key}")
        logger.info(f"📁 Processing uploaded file: {filename}")
      
        # Process the file
        result = process_bn_single_file(
            dest_path,
            raw_dir,
            clean_dir,
            ts,
            params['pass_threshold'],
            semester_course_maps,
            semester_credit_units,
            semester_lookup,
            semester_course_titles,
            DEFAULT_LOGO_PATH,
            semester_key,
            set_name,
            previous_gpas=None,
            upgrade_min_threshold=get_upgrade_threshold_from_env()
        )
      
        if result is not None:
            logger.info("✅ Successfully processed uploaded file")
            return True
        else:
            logger.error("❌ Failed to process uploaded file")
            return False
          
    except Exception as e:
        logger.error(f"❌ Error processing uploaded file: {e}")
        traceback.print_exc()
        return False
    finally:
        # Clean up temporary directory
        shutil.rmtree(temp_dir, ignore_errors=True)
def get_upgrade_threshold_from_env():
    """Get upgrade threshold from environment variables"""
    upgrade_threshold_str = os.getenv('UPGRADE_THRESHOLD', '0').strip() # Default to '0'
    if upgrade_threshold_str and upgrade_threshold_str.isdigit():
        upgrade_value = int(upgrade_threshold_str)
        if upgrade_value == 0:
            return None # 0 means disabled
        if 45 <= upgrade_value <= 49:
            return upgrade_value
    return None # Default to disabled (no upgrade)
def check_bn_files_exist(raw_dir, semester_key):
    """Check if BN files actually exist for the given semester"""
    if not os.path.exists(raw_dir):
        logger.error(f"❌ Raw directory doesn't exist: {raw_dir}")
        return False
  
    raw_files = [f for f in os.listdir(raw_dir) if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
  
    if not raw_files:
        logger.error(f"❌ No Excel files found in: {raw_dir}")
        return False
  
    # Check if any files match the semester
    semester_files = []
    for rf in raw_files:
        detected_sem = detect_bn_semester_from_filename(rf)
        if detected_sem == semester_key:
            semester_files.append(rf)
  
    if not semester_files:
        logger.error(f"❌ No files found for semester {semester_key}")
        logger.info(f" Available files: {raw_files}")
        return False
  
    logger.info(f"✅ Found {len(semester_files)} files for {semester_key}: {semester_files}")
    return True
def process_in_non_interactive_mode(params, base_dir_norm):
    """Process exams in non-interactive mode for web interface."""
    logger.info("🔧 Running in NON-INTERACTIVE mode (web interface)")
  
    # Use parameters from environment variables
    selected_set = params['selected_set']
    processing_mode = params['processing_mode']
    selected_semesters = params['selected_semesters']
  
    # Get upgrade threshold from environment variable if provided
    upgrade_min_threshold = get_upgrade_threshold_from_env()
  
    # Get available sets
    available_sets = get_available_bn_sets(base_dir_norm)
  
    if not available_sets:
        logger.error("❌ No BN sets found")
        return False
  
    # Remove BN-COURSES from available sets if present
    available_sets = [s for s in available_sets if s != 'BN-COURSES']
  
    if not available_sets:
        logger.error("❌ No valid BN sets found (only BN-COURSES present)")
        return False
  
    # Determine which sets to process - FIXED: Only process the selected set
    if selected_set == "all":
        sets_to_process = available_sets
        logger.info(f"🎯 Processing ALL sets: {sets_to_process}")
    else:
        # CRITICAL FIX: Only process the specifically selected set
        if selected_set in available_sets:
            sets_to_process = [selected_set]
            logger.info(f"🎯 Processing ONLY selected set: {selected_set}")
        else:
            logger.error(f"❌ Selected set '{selected_set}' not found in available sets: {available_sets}")
            return False # Don't fall back to processing all sets
  
    # Determine which semesters to process
    if processing_mode == "auto" or not selected_semesters or 'all' in selected_semesters:
        semesters_to_process = BN_SEMESTER_ORDER.copy()
        logger.info(f"🎯 Processing ALL semesters: {semesters_to_process}")
    else:
        semesters_to_process = selected_semesters
        logger.info(f"🎯 Processing selected semesters: {semesters_to_process}")
  
    # Load course data once
    try:
        semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_bn_course_data()
    except Exception as e:
        logger.error(f"❌ Could not load course data: {e}")
        return False
  
    ts = datetime.now().strftime(TIMESTAMP_FMT)
  
    # Process each set and semester
    total_processed = 0
    for bn_set in sets_to_process:
        logger.info(f"\n{'='*60}")
        logger.info(f"PROCESSING SET: {bn_set}")
        logger.info(f"{'='*60}")
      
        # UPDATED: Raw and clean directories now under BN folder
        raw_dir = normalize_path(os.path.join(base_dir_norm, "BN", bn_set, "RAW_RESULTS"))
        clean_dir = normalize_path(os.path.join(base_dir_norm, "BN", bn_set, "CLEAN_RESULTS"))
      
        # Create directories if they don't exist
        os.makedirs(raw_dir, exist_ok=True)
        os.makedirs(clean_dir, exist_ok=True)
      
        if not os.path.exists(raw_dir):
            logger.warning(f"⚠️ RAW_RESULTS directory not found: {raw_dir}")
            continue
      
        raw_files = [f for f in os.listdir(raw_dir)
                     if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
        if not raw_files:
            logger.warning(f"⚠️ No raw files in {raw_dir}; skipping {bn_set}")
            continue
      
        logger.info(f"📁 Found {len(raw_files)} raw files in {bn_set}: {raw_files}")
      
        # Create timestamped folder for this set
        set_output_dir = os.path.join(clean_dir, "{}_RESULT-{}".format(bn_set, ts))
        os.makedirs(set_output_dir, exist_ok=True)
        logger.info(f"📁 Created BN set output directory: {set_output_dir}")
      
        # Process selected semesters
        semester_processed = 0
        for semester_key in semesters_to_process:
            if semester_key not in BN_SEMESTER_ORDER:
                logger.warning(f"⚠️ Skipping unknown semester: {semester_key}")
                continue
          
            # Check if there are files for this semester - USING BM'S APPROACH
            semester_files_exist = False
            for rf in raw_files:
                try:
                    detected_sem = detect_bn_semester_from_filename(rf)
                    if detected_sem == semester_key:
                        semester_files_exist = True
                        break
                except ValueError as e:
                    logger.warning(f"⚠️ Could not detect semester for {rf}: {e}")
                    continue
          
            if semester_files_exist:
                logger.info(f"\n🎯 Processing BN {semester_key} in {bn_set}...")
                try:
                    # Add file existence check
                    if not check_bn_files_exist(raw_dir, semester_key):
                        logger.error(f"❌ Skipping {semester_key} - no valid files found")
                        continue
                      
                    # Process the semester with the upgrade threshold
                    result = process_bn_semester_files(
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
                        bn_set,
                        previous_gpas=None,
                        upgrade_min_threshold=upgrade_min_threshold
                    )
                  
                    if result and result.get('success', False):
                        logger.info(f"✅ Successfully processed {semester_key}")
                        total_processed += result.get('files_processed', 0)
                        semester_processed += result.get('files_processed', 0)
                    else:
                        logger.error(f"❌ Failed to process {semester_key}")
                      
                except Exception as e:
                    logger.error(f"❌ Error processing {semester_key}: {e}")
                    traceback.print_exc()
            else:
                logger.warning(f"⚠️ No files found for {semester_key} in {bn_set}, skipping...")
      
        # Create ZIP of BN results ONLY if files were processed
        if semester_processed > 0:
            try:
                zip_success = create_bn_zip_for_set(clean_dir, bn_set, ts, set_output_dir)
                if zip_success:
                    logger.info(f"✅ Successfully created ZIP for {bn_set}")
                else:
                    logger.warning(f"⚠️ ZIP creation failed for {bn_set}, files remain in: {set_output_dir}")
            except Exception as e:
                logger.warning(f"⚠️ Failed to create BN ZIP for {bn_set}: {e}")
                traceback.print_exc()
        else:
            logger.warning(f"⚠️ No files processed for {bn_set}, skipping ZIP creation")
            if os.path.exists(set_output_dir):
                shutil.rmtree(set_output_dir)
  
    # Only return True if files were actually processed
    if total_processed == 0:
        logger.error("❌ No files were actually processed. Check if:")
        logger.error(" - Raw files exist in the correct directory")
        logger.error(" - File naming matches semester patterns")
        logger.error(" - Course data files are available")
        return False
  
    # Print BN-specific summaries
    logger.info("\n📊 BN STUDENT TRACKING SUMMARY:")
    logger.info(f"Total unique BN students tracked: {len(STUDENT_TRACKER)}")
    logger.info(f"Total BN withdrawn students: {len(WITHDRAWN_STUDENTS)}")
    if CARRYOVER_STUDENTS:
        logger.info("\n📋 BN CARRYOVER STUDENT SUMMARY:")
        logger.info(f"Total BN carryover students: {len(CARRYOVER_STUDENTS)}")
  
    return True
def get_form_parameters():
    """Get parameters from environment variables set by the web form."""
    logger.info("🎯 DEBUG - FORM PARAMETERS:")
    logger.info(f" SELECTED_SET: {os.getenv('SELECTED_SET')}")
    logger.info(f" PROCESSING_MODE: {os.getenv('PROCESSING_MODE')}")
    logger.info(f" SELECTED_SEMESTERS: {os.getenv('SELECTED_SEMESTERS')}")
    logger.info(f" UPGRADE_THRESHOLD: {os.getenv('UPGRADE_THRESHOLD')}")
  
    selected_set = os.getenv('SELECTED_SET', 'all')
    processing_mode = os.getenv('PROCESSING_MODE', 'auto')
    selected_semesters_str = os.getenv('SELECTED_SEMESTERS', '')
    pass_threshold = float(os.getenv('PASS_THRESHOLD', '50.0'))
    generate_pdf = os.getenv('GENERATE_PDF', 'True').lower() == 'true'
    track_withdrawn = os.getenv('TRACK_WITHDRAWN', 'True').lower() == 'true'
  
    # Convert semester string to list
    selected_semesters = []
    if selected_semesters_str:
        selected_semesters = selected_semesters_str.split(',')
  
    logger.info("🎯 FINAL PARAMETERS:")
    logger.info(f" Selected Set: {selected_set}")
    logger.info(f" Processing Mode: {processing_mode}")
    logger.info(f" Selected Semesters: {selected_semesters}")
    logger.info(f" Pass Threshold: {pass_threshold}")
  
    return {
        'selected_set': selected_set,
        'processing_mode': processing_mode,
        'selected_semesters': selected_semesters,
        'pass_threshold': pass_threshold,
        'generate_pdf': generate_pdf,
        'track_withdrawn': track_withdrawn
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
# ----------------------------
# Enhanced Course Data Loading for BN
# ----------------------------
def load_bn_course_data():
    """
    Reads N-course-code-creditUnit.xlsx and returns BN course data
    """
    course_file = os.path.join(BN_COURSES_DIR, "N-course-code-creditUnit.xlsx")
    logger.info(f"Loading BN course data from: {course_file}")
  
    if not os.path.exists(course_file):
        raise FileNotFoundError(f"BN course file not found: {course_file}")
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
            logger.warning(f"Warning: sheet '{sheet}' missing expected columns {expected} — skipped")
            continue
          
        # Enhanced data cleaning for BN
        dfx = df.dropna(subset=['COURSE CODE', 'COURSE TITLE'])
        dfx = dfx[~dfx['COURSE CODE'].astype(str).str.contains('TOTAL', case=False, na=False)]
      
        valid_mask = dfx['CU'].astype(str).str.replace('.', '', regex=False).str.isdigit()
        dfx = dfx[valid_mask]
      
        if dfx.empty:
            logger.warning(f"Warning: sheet '{sheet}' has no valid rows after cleaning — skipped")
            continue
          
        codes = dfx['COURSE CODE'].astype(str).str.strip().tolist()
        titles = dfx['COURSE TITLE'].astype(str).str.strip().tolist()
        cus = dfx['CU'].astype(float).astype(int).tolist()
        # Create enhanced course map for BN
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
        # Create BN-specific lookup variations
        norm = normalize_for_matching(sheet)
        semester_lookup[norm] = sheet
        # Add BN-specific variations
        norm_no_bn = norm.replace('bn-', '').replace('bn ', '')
        semester_lookup[norm_no_bn] = sheet
    return semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles
# ----------------------------
# Data Transformation Functions
# ----------------------------
def transform_transposed_data(df, sheet_type):
    """
    Transform transposed data format to wide format for BN.
    Input: Each student appears multiple times with different courses
    Output: Each student appears once with all courses as columns
    """
    logger.info(f"🔄 Transforming {sheet_type} sheet from transposed to wide format...")
  
    # Find the registration and name columns for BN
    reg_col = find_column_by_names(df, ["REG. No", "Reg No", "Registration Number", "EXAM NUMBER"])
    name_col = find_column_by_names(df, ["NAME", "Full Name", "Candidate Name"])
  
    if not reg_col:
        logger.error("❌ Could not find registration column for transformation")
        return df
  
    # Get all course columns (columns that contain course codes)
    course_columns = [col for col in df.columns
                     if col not in [reg_col, name_col] and col not in ['', None]]
  
    logger.info(f"📊 Found {len(course_columns)} course columns: {course_columns}")
  
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
                column_name = "{}_{}".format(course_col, sheet_type)
                student_dict[exam_no][column_name] = score
  
    # Convert dictionary to list
    transformed_data = list(student_dict.values())
  
    # Create new DataFrame
    if transformed_data:
        transformed_df = pd.DataFrame(transformed_data)
        logger.info(f"✅ Transformed data: {len(transformed_df)} students, {len(transformed_df.columns)} columns")
        return transformed_df
    else:
        logger.error("❌ No data after transformation")
        return df
def detect_data_format(df, sheet_type):
    """
    Detect if data is in transposed format (students appear multiple times)
    Returns True if transposed format is detected - IMPROVED for BN
    """
    reg_col = find_column_by_names(df, ["REG. No", "Reg No", "EXAM NUMBER"])
  
    if not reg_col or df.empty:
        return False
  
    # Clean data first
    valid_exam_nos = df[reg_col].dropna().astype(str).str.strip()
    valid_exam_nos = valid_exam_nos[valid_exam_nos != '']
  
    if len(valid_exam_nos) == 0:
        return False
  
    # Count occurrences
    student_counts = valid_exam_nos.value_counts()
    duplicates = (student_counts > 1).sum()
    duplicate_percentage = (duplicates / len(student_counts)) * 100
  
    # More than 30% students appearing multiple times = transposed
    is_transposed = duplicate_percentage > 30
  
    if is_transposed:
        logger.info(f"📊 Transposed data detected in {sheet_type}:")
        logger.info(f" {duplicates} students with duplicates ({duplicate_percentage:.1f}%)")
  
    return is_transposed
# ----------------------------
# Enhanced Course Matching for BN
# ----------------------------
def normalize_course_name(name):
    """Enhanced normalization for BN course title matching"""
    if not isinstance(name, str):
        return ""
  
    # Convert to lowercase and remove extra spaces
    normalized = name.lower().strip()
  
    # Replace multiple spaces with single space
    normalized = re.sub(r'\s+', ' ', normalized)
  
    # Remove special characters and extra words
    normalized = re.sub(r'[^\w\s]', '', normalized)
  
    # BN-specific substitutions for variations
    bn_substitutions = {
        'coomunication': 'communication',
        'nsg': 'nursing',
        'foundation': 'foundations',
        'of of': 'of',
        'emergency care': 'emergency',
        'nursing/ emergency': 'nursing emergency',
        'care i': 'care',
        'foundations of nursing': 'foundations nursing',
        'foundation of nsg': 'foundations nursing',
        'foundation of nursing': 'foundations nursing',
        # BN-specific courses
        'maternal': 'maternal health',
        'child health': 'child health nursing',
        'community health': 'community health nursing',
        'psychiatric': 'psychiatric nursing'
    }
  
    for old, new in bn_substitutions.items():
        normalized = normalized.replace(old, new)
      
    return normalized.strip()
def find_best_course_match(column_name, course_map):
    """Find the best matching BN course using enhanced matching algorithm."""
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
            # Bonus for matching BN key words
            bn_key_words = ['nursing', 'health', 'care', 'maternal', 'child', 'community', 'psychiatric']
            for word in bn_key_words:
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
        if ratio > best_ratio and ratio > 0.6:
            best_ratio = ratio
            best_match = course_info
  
    return best_match
# ----------------------------
# Carryover Management for BN - FIXED VERSION
# ----------------------------
def initialize_carryover_tracker():
    """Initialize the global carryover tracker for BN."""
    global CARRYOVER_STUDENTS
    CARRYOVER_STUDENTS = {}
    # Load previous carryover records from all JSON files
    carryover_jsons = glob.glob(os.path.join(BN_BASE_DIR, "**/co_student*.json"), recursive=True)
    for jf in sorted(carryover_jsons, key=os.path.getmtime): # Load in chronological order, later files override
        try:
            with open(jf, 'r') as f:
                data = json.load(f)
                for student in data:
                    student_key = f"{student['exam_number']}_{student['semester']}"
                    CARRYOVER_STUDENTS[student_key] = student
        except Exception as e:
            logger.warning(f"⚠️ Failed to load carryover from {jf}: {e}")
    logger.info(f"📂 Loaded {len(CARRYOVER_STUDENTS)} previous carryover records")
  
def identify_carryover_students(mastersheet_df, semester_key, set_name, pass_threshold=50.0):
    """
    Identify BN students with carryover courses from current semester processing.
    USING ND SCRIPT LOGIC
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
      
        # Check if student is withdrawn - don't create carryover for withdrawn students
        remarks = str(student.get('REMARKS', ''))
        if 'Advised to Withdraw' in remarks or 'Withdrawn' in remarks:
            continue
      
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
                        'status': 'Failed'
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
                'status': 'Active'
            }
            carryover_students.append(carryover_data)
          
            # Update global tracker
            student_key = f"{exam_no}_{semester_key}"
            CARRYOVER_STUDENTS[student_key] = carryover_data
           
            logger.info(f"📝 Identified carryover: {exam_no} - {len(failed_courses)} failed courses")
  
    logger.info(f"📊 Total carryover students identified: {len(carryover_students)}")
    return carryover_students
def save_carryover_records(carryover_students, output_dir, set_name, semester_key):
    """
    Save BN carryover student records to the clean results folder.
    USING ND SCRIPT SIMPLIFIED APPROACH
    """
    if not carryover_students:
        logger.info("ℹ️ No carryover students to save")
        return None
   
    # Use the output_dir directly (simplified from ND script)
    carryover_dir = os.path.join(output_dir, "CARRYOVER_RECORDS")
    os.makedirs(carryover_dir, exist_ok=True)
   
    logger.info(f"📁 Saving carryover records to: {carryover_dir}")
  
    # Generate filename with set and semester tags
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"co_student_{set_name}_{semester_key}_{timestamp}"
  
    # Save as Excel - SIMPLE STRUCTURE LIKE ND SCRIPT
    excel_file = os.path.join(carryover_dir, f"{filename}.xlsx")
  
    # Prepare data for Excel - simple structure without enhanced formatting
    records_data = []
    for student in carryover_students:
        for course in student['failed_courses']:
            record = {
                'EXAMS NUMBER': student['exam_number'],
                'NAME': student['name'],
                'COURSE CODE': course['course_code'],
                'ORIGINAL SCORE': course['original_score'],
                'SEMESTER': student['semester'],
                'SET': student['set'],
                'RESIT ATTEMPTS': course['resit_attempts'],
                'BEST SCORE': course['best_score'],
                'STATUS': course['status'],
                'IDENTIFIED DATE': student['identified_date']
            }
            records_data.append(record)
  
    if records_data:
        df = pd.DataFrame(records_data)
        df.to_excel(excel_file, index=False)
        logger.info(f"✅ BN Carryover records saved: {excel_file}")
      
        # Add basic formatting only (like ND script)
        try:
            wb = load_workbook(excel_file)
            ws = wb.active
           
            # Simple header formatting only
            header_row = ws[1]
            for cell in header_row:
                cell.font = Font(bold=True)
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
                adjusted_width = min(max_length + 2, 30)
                ws.column_dimensions[column_letter].width = adjusted_width
           
            wb.save(excel_file)
            logger.info("✅ Added basic formatting to BN carryover Excel file")
          
        except Exception as e:
            logger.warning(f"⚠️ Could not add basic formatting to BN carryover file: {e}")
  
    # Save as JSON for easy processing
    json_file = os.path.join(carryover_dir, f"{filename}.json")
    with open(json_file, 'w') as f:
        json.dump(carryover_students, f, indent=2)
  
    logger.info(f"📁 BN Carryover records saved in: {carryover_dir}")
   
    # Print summary
    total_failed_courses = sum(len(s['failed_courses']) for s in carryover_students)
    logger.info(f"📊 Carryover Summary: {len(carryover_students)} students, {total_failed_courses} failed courses")
   
    return carryover_dir
# ----------------------------
# CGPA Tracking for BN - FIXED VERSION
# ----------------------------
def create_bn_cgpa_summary_sheet(mastersheet_path, ts):
    """Create a CGPA summary sheet that aggregates GPA across all BN semesters."""
    try:
        logger.info("📊 Creating BN CGPA Summary Sheet...")
      
        # Load the mastersheet workbook
        wb = load_workbook(mastersheet_path)
      
        # Collect GPA data from all BN semesters
        cgpa_data = {}
      
        for sheet_name in wb.sheetnames:
            if sheet_name in BN_SEMESTER_ORDER:
                df = pd.read_excel(mastersheet_path, sheet_name=sheet_name, header=5)
              
                # Filter valid rows
                exam_col = find_exam_number_column(df)
                if exam_col:
                    df = df[df[exam_col].notna() & (df[exam_col] != '')]
              
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
                'EXAMS NUMBER': exam_no,
                'NAME': data['name']
            }
          
            # Add GPA for each semester
            total_gpa = 0
            semester_count = 0
          
            for semester in BN_SEMESTER_ORDER:
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
        headers = ['EXAMS NUMBER', 'NAME'] + BN_SEMESTER_ORDER + ['CGPA']
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
        logger.info("✅ BN CGPA Summary sheet created successfully")
      
        return summary_df
      
    except Exception as e:
        logger.error(f"❌ Error creating BN CGPA summary sheet: {e}")
        return None
# ----------------------------
# Analysis Sheet for BN - FIXED COMPREHENSIVE VERSION
# ----------------------------
def create_bn_analysis_sheet(mastersheet_path, ts, semester_credit_units):
    """Create an analysis sheet with comprehensive statistics for BN - FIXED VERSION."""
    try:
        logger.info("📈 Creating BN Analysis Sheet...")
      
        wb = load_workbook(mastersheet_path)
      
        # Collect data from all semesters
        analysis_data = {
            'semester': [],
            'total_students': [],
            'passed_all': [],
            'carryover_students': [],
            'withdrawn_students': [],
            'average_gpa': [],
            'pass_rate': [],
            'total_courses': [],
            'total_credit_units': []
        }
      
        for sheet_name in wb.sheetnames:
            if sheet_name in BN_SEMESTER_ORDER:
                df = pd.read_excel(mastersheet_path, sheet_name=sheet_name, header=5)
              
                # Filter valid rows: drop rows where 'EXAMS NUMBER' is NaN or empty
                exam_col = find_exam_number_column(df)
                if exam_col:
                    df = df[df[exam_col].notna() & (df[exam_col] != '')]
              
                # Basic statistics
                total_students = len(df)
              
                # Passed all: REMARKS == 'Passed' (strip whitespace)
                passed_all = 0
                if 'REMARKS' in df.columns:
                    df['REMARKS'] = df['REMARKS'].astype(str).str.strip()
                    passed_all = len(df[df['REMARKS'] == 'Passed'])
              
                # FIXED: Carryover calculation - students with failed courses but not withdrawn
                carryover_count = 0
                if 'REMARKS' in df.columns:
                    carryover_count = len(df[
                        (df['REMARKS'].str.contains('Failed', case=False, na=False)) &
                        (~df['REMARKS'].str.contains('Withdrawn', case=False, na=False))
                    ])
              
                # FIXED: Withdrawn calculation - look for "Advised to Withdraw" or "Withdrawn"
                withdrawn_count = 0
                if 'REMARKS' in df.columns:
                    withdrawn_count = len(df[
                        (df['REMARKS'].str.contains('Advised to Withdraw', case=False, na=False)) |
                        (df['REMARKS'].str.contains('Withdrawn', case=False, na=False))
                    ])
              
                # Average GPA: mean of non-NaN GPA values
                avg_gpa = 0
                if 'GPA' in df.columns:
                    gpa_series = pd.to_numeric(df['GPA'], errors='coerce')
                    avg_gpa = gpa_series.dropna().mean() if not gpa_series.dropna().empty else 0
              
                # FIXED: Pass rate calculation - use passed_all instead of total students
                pass_rate = (passed_all / total_students * 100) if total_students > 0 else 0
              
                # Course and credit unit information
                total_courses = 0
                total_credit_units = 0
                if sheet_name in semester_credit_units:
                    total_courses = len(semester_credit_units[sheet_name])
                    total_credit_units = sum(semester_credit_units[sheet_name].values())
              
                analysis_data['semester'].append(sheet_name)
                analysis_data['total_students'].append(total_students)
                analysis_data['passed_all'].append(passed_all)
                analysis_data['carryover_students'].append(carryover_count)
                analysis_data['withdrawn_students'].append(withdrawn_count)
                analysis_data['average_gpa'].append(round(avg_gpa, 2))
                analysis_data['pass_rate'].append(round(pass_rate, 2))
                analysis_data['total_courses'].append(total_courses)
                analysis_data['total_credit_units'].append(total_credit_units)
      
        # Create analysis dataframe
        analysis_df = pd.DataFrame(analysis_data)
      
        # Add overall statistics
        if not analysis_df.empty:
            overall_stats = {
                'semester': 'OVERALL',
                'total_students': analysis_df['total_students'].sum(),
                'passed_all': analysis_df['passed_all'].sum(),
                'carryover_students': analysis_df['carryover_students'].sum(),
                'withdrawn_students': analysis_df['withdrawn_students'].sum(),
                'average_gpa': round(analysis_df['average_gpa'].mean(), 2),
                'pass_rate': round(analysis_df['pass_rate'].mean(), 2),
                'total_courses': analysis_df['total_courses'].sum(),
                'total_credit_units': analysis_df['total_credit_units'].sum()
            }
            analysis_df = pd.concat([analysis_df, pd.DataFrame([overall_stats])], ignore_index=True)
      
        # Add the analysis sheet to the workbook
        if 'ANALYSIS' in wb.sheetnames:
            del wb['ANALYSIS']
      
        ws = wb.create_sheet('ANALYSIS')
      
        # Write header
        headers = ['SEMESTER', 'TOTAL STUDENTS', 'PASSED ALL', 'CARRYOVER STUDENTS',
                  'WITHDRAWN STUDENTS', 'AVERAGE GPA', 'PASS RATE (%)', 'TOTAL COURSES', 'TOTAL CREDIT UNITS']
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
            ws.cell(row=row_idx+2, column=8, value=row_data['total_courses'])
            ws.cell(row=row_idx+2, column=9, value=row_data['total_credit_units'])
      
        # Style the header
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="27ae60", end_color="27ae60", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
      
        # Style data rows
        for row in range(2, len(analysis_df) + 2):
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=row, column=col)
                if row == len(analysis_df) + 1: # Overall row
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
                cell.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                   top=Side(style="thin"), bottom=Side(style="thin"))
      
        wb.save(mastersheet_path)
        logger.info("✅ BN Analysis sheet created successfully")
      
        return analysis_df
      
    except Exception as e:
        logger.error(f"❌ Error creating BN analysis sheet: {e}")
        traceback.print_exc()
        return None
# ----------------------------
# ZIP File Creation for BN - FIXED VERSION
# ----------------------------
def create_bn_zip_for_set(clean_dir, set_name, ts, set_output_dir):
    """Create ZIP file for BN set with verification."""
    try:
        # Wait for file operations
        time.sleep(2)
      
        # Verify source directory
        if not os.path.exists(set_output_dir):
            logger.error(f"❌ Output directory doesn't exist: {set_output_dir}")
            return False
      
        # Count files
        file_count = sum(len(files) for _, _, files in os.walk(set_output_dir))
        if file_count == 0:
            logger.error(f"❌ No files to zip in: {set_output_dir}")
            return False
      
        # Create ZIP
        zip_filename = f"{set_name}_RESULT-{ts}.zip"
        zip_path = os.path.join(clean_dir, zip_filename)
      
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(set_output_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, set_output_dir)
                    zipf.write(file_path, arcname)
      
        # Verify ZIP
        if os.path.exists(zip_path) and os.path.getsize(zip_path) > 1000:
            with zipfile.ZipFile(zip_path, 'r') as test_zip:
                if len(test_zip.namelist()) == file_count:
                    # Safe to remove original
                    shutil.rmtree(set_output_dir)
                    logger.info(f"✅ ZIP created and verified: {zip_path}")
                    return True
      
        return False
      
    except Exception as e:
        logger.error(f"❌ ZIP creation failed: {e}")
        return False
# ----------------------------
# Enhanced Semester Detection for BN - FIXED VERSION
# ----------------------------
def detect_bn_semester_from_filename(filename):
    """
    Detect semester from filename for BN program - HANDLES ALL FORMATS INCLUDING YOURS.
    Returns semester_key
    """
    filename_upper = filename.upper()
  
    logger.info(f"🔍 Detecting semester from: '{filename}'")
    # ✅ COMPREHENSIVE patterns matching YOUR file format: THIRD-YEAR-SECOND-SEMESTER.xlsx
    semester_patterns = {
        # Pattern 1: Direct word matching (handles hyphens, underscores, spaces)
        # Matches: THIRD-YEAR-SECOND-SEMESTER, THIRD_YEAR_SECOND_SEMESTER, THIRD YEAR SECOND SEMESTER
        r'(?:N[-_\s])?(?:THIRD|3RD)[-_\s]+YEAR[-_\s]+(?:SECOND|2ND)[-_\s]+SEMESTER': "N-THIRD-YEAR-SECOND-SEMESTER",
        r'(?:N[-_\s])?(?:THIRD|3RD)[-_\s]+YEAR[-_\s]+(?:FIRST|1ST)[-_\s]+SEMESTER': "N-THIRD-YEAR-FIRST-SEMESTER",
        r'(?:N[-_\s])?(?:SECOND|2ND)[-_\s]+YEAR[-_\s]+(?:SECOND|2ND)[-_\s]+SEMESTER': "N-SECOND-YEAR-SECOND-SEMESTER",
        r'(?:N[-_\s])?(?:SECOND|2ND)[-_\s]+YEAR[-_\s]+(?:FIRST|1ST)[-_\s]+SEMESTER': "N-SECOND-YEAR-FIRST-SEMESTER",
        r'(?:N[-_\s])?(?:FIRST|1ST)[-_\s]+YEAR[-_\s]+(?:SECOND|2ND)[-_\s]+SEMESTER': "N-FIRST-YEAR-SECOND-SEMESTER",
        r'(?:N[-_\s])?(?:FIRST|1ST)[-_\s]+YEAR[-_\s]+(?:FIRST|1ST)[-_\s]+SEMESTER': "N-FIRST-YEAR-FIRST-SEMESTER",
      
        # Pattern 2: Flexible matching with wildcards
        # Matches variations with different word orders or separators
        r'(?:THIRD|3RD).*YEAR.*(?:SECOND|2ND).*SEMESTER': "N-THIRD-YEAR-SECOND-SEMESTER",
        r'(?:THIRD|3RD).*YEAR.*(?:FIRST|1ST).*SEMESTER': "N-THIRD-YEAR-FIRST-SEMESTER",
        r'(?:SECOND|2ND).*YEAR.*(?:SECOND|2ND).*SEMESTER': "N-SECOND-YEAR-SECOND-SEMESTER",
        r'(?:SECOND|2ND).*YEAR.*(?:FIRST|1ST).*SEMESTER': "N-SECOND-YEAR-FIRST-SEMESTER",
        r'(?:FIRST|1ST).*YEAR.*(?:SECOND|2ND).*SEMESTER': "N-FIRST-YEAR-SECOND-SEMESTER",
        r'(?:FIRST|1ST).*YEAR.*(?:FIRST|1ST).*SEMESTER': "N-FIRST-YEAR-FIRST-SEMESTER",
      
        # Pattern 3: Year X Semester Y format
        # Matches: YEAR 3 SEMESTER 2, YR3SEM2, etc.
        r'(?:YEAR|YR)[-_\s]*3[-_\s]*(?:SEMESTER|SEM)[-_\s]*2': "N-THIRD-YEAR-SECOND-SEMESTER",
        r'(?:YEAR|YR)[-_\s]*3[-_\s]*(?:SEMESTER|SEM)[-_\s]*1': "N-THIRD-YEAR-FIRST-SEMESTER",
        r'(?:YEAR|YR)[-_\s]*2[-_\s]*(?:SEMESTER|SEM)[-_\s]*2': "N-SECOND-YEAR-SECOND-SEMESTER",
        r'(?:YEAR|YR)[-_\s]*2[-_\s]*(?:SEMESTER|SEM)[-_\s]*1': "N-SECOND-YEAR-FIRST-SEMESTER",
        r'(?:YEAR|YR)[-_\s]*1[-_\s]*(?:SEMESTER|SEM)[-_\s]*2': "N-FIRST-YEAR-SECOND-SEMESTER",
        r'(?:YEAR|YR)[-_\s]*1[-_\s]*(?:SEMESTER|SEM)[-_\s]*1': "N-FIRST-YEAR-FIRST-SEMESTER",
    }
  
    for pattern, semester_key in semester_patterns.items():
        match = re.search(pattern, filename_upper)
        if match:
            logger.info(f"✅ MATCHED: '{filename}' → '{semester_key}'")
            logger.info(f" Pattern: {pattern}")
            logger.info(f" Matched text: '{match.group()}'")
            return semester_key
  
    # If no match, log and raise error
    logger.error(f"❌ NO MATCH for filename: '{filename}'")
    logger.error(f" Uppercase: '{filename_upper}'")
    raise ValueError(f"Could not detect semester from filename: {filename}")
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
    """Find exam number column in dataframe"""
    return find_column_by_names(df, ["REG. No", "Reg No", "Registration Number", "EXAM NUMBER", "EXAMS NUMBER"])
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
    """
    global STUDENT_TRACKER, WITHDRAWN_STUDENTS
    logger.info(f"📊 Updating student tracker for {semester_key}")
    logger.info(f"📝 Current students in this semester: {len(exam_numbers)}")
    # Track withdrawn students
    if withdrawn_students:
        for exam_no in withdrawn_students:
            if exam_no not in WITHDRAWN_STUDENTS:
                WITHDRAWN_STUDENTS[exam_no] = {
                    'withdrawn_semester': semester_key,
                    'withdrawn_date': datetime.now().strftime(TIMESTAMP_FMT),
                    'reappeared_semesters': []
                }
                logger.info(f"🚫 Marked as withdrawn: {exam_no} in {semester_key}")
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
                logger.warning(f"⚠️ PREVIOUSLY WITHDRAWN STUDENT REAPPEARED: {exam_no}")
                if exam_no in WITHDRAWN_STUDENTS:
                    if semester_key not in WITHDRAWN_STUDENTS[exam_no]['reappeared_semesters']:
                        WITHDRAWN_STUDENTS[exam_no]['reappeared_semesters'].append(
                            semester_key)
    logger.info(f"📈 Total unique students tracked: {len(STUDENT_TRACKER)}")
    logger.info(f"🚫 Total withdrawn students: {len(WITHDRAWN_STUDENTS)}")
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
        exam_no = str(row['EXAMS NUMBER']).strip()
        if is_student_withdrawn(exam_no):
            withdrawal_history = get_withdrawal_history(exam_no)
            # Only remove if student was withdrawn in a PREVIOUS semester
            if withdrawal_history and withdrawal_history['withdrawn_semester'] != semester_key:
                removed_students.append(exam_no)
                filtered_mastersheet = filtered_mastersheet[filtered_mastersheet["EXAMS NUMBER"] != exam_no]
    if removed_students:
        logger.info(
            f"🚫 Removed {len(removed_students)} previously withdrawn students from {semester_key}:")
        for exam_no in removed_students:
            withdrawal_history = get_withdrawal_history(exam_no)
            logger.info(
                f" - {exam_no} (withdrawn in {withdrawal_history['withdrawn_semester']})")
    return filtered_mastersheet, removed_students
# ----------------------------
# Set Selection Functions
# ----------------------------
def get_available_bn_sets(base_dir):
    """Get all available BN sets (SET47, SET48, etc.)"""
    # UPDATED: Sets are now under BN folder
    bn_dir = os.path.join(base_dir, "BN")
    if not os.path.exists(bn_dir):
        logger.error(f"❌ BN directory not found: {bn_dir}")
        return []
      
    sets = []
    for item in os.listdir(bn_dir):
        item_path = os.path.join(bn_dir, item)
        if os.path.isdir(item_path) and item.upper().startswith("SET"):
            sets.append(item)
    return sorted(sets)
def get_user_set_choice(available_sets):
    """
    Prompt user to choose which set to process.
    Returns the selected set directory name.
    """
    print("\n🎯 AVAILABLE SETS:")
    for i, set_name in enumerate(available_sets, 1):
        print("{}. {}".format(i, set_name))
    print("{}. Process ALL sets".format(len(available_sets) + 1))
    while True:
        try:
            choice = input(
                "\nEnter your choice (1-{}): ".format(len(available_sets) + 1)).strip()
            if not choice:
                print("❌ Please enter a choice.")
                continue
            if choice.isdigit():
                choice_num = int(choice)
                if 1 <= choice_num <= len(available_sets):
                    selected_set = available_sets[choice_num - 1]
                    print("✅ Selected set: {}".format(selected_set))
                    return [selected_set]
                elif choice_num == len(available_sets) + 1:
                    print("✅ Selected: ALL sets")
                    return available_sets
                else:
                    print(
                        "❌ Invalid choice. Please enter a number between 1-{}.".format(len(available_sets) + 1))
            else:
                print("❌ Please enter a valid number.")
        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)
        except Exception as e:
            print("❌ Error: {}. Please try again.".format(e))
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
            return 5.0 # A
        elif score >= 60:
            return 4.0 # B
        elif score >= 50:
            return 3.0 # C
        elif score >= 45:
            return 2.0 # D
        elif score >= 40:
            return 1.0 # E
        else:
            return 0.0 # F
    except BaseException:
        return 0.0
# ----------------------------
# Upgrade Rule Functions
# ----------------------------
def get_upgrade_threshold_from_user(semester_key, set_name):
    """
    Prompt user to choose upgrade threshold for BN results.
    Returns: (min_threshold, upgraded_count) or (None, 0) if skipped
    """
    print("\n🎯 MANAGEMENT THRESHOLD UPGRADE RULE DETECTED")
    print("📚 Semester: {}".format(semester_key))
    print("📁 Set: {}".format(set_name))
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
                print("✅ Upgrade rule selected: {}–49 → 50".format(min_threshold))
                return min_threshold, 0
            else:
                print("❌ Invalid choice. Please enter 0, 45, 46, 47, 48, or 49.")
              
        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)
        except Exception as e:
            print("❌ Error: {}. Please try again.".format(e))
def apply_upgrade_rule(mastersheet, ordered_codes, min_threshold):
    """
    Apply upgrade rule to mastersheet scores.
    Returns: (updated_mastersheet, upgraded_count)
    """
    if min_threshold is None:
        return mastersheet, 0
  
    upgraded_count = 0
    upgraded_students = set()
  
    logger.info(f"🔄 Applying upgrade rule: {min_threshold}–49 → 50")
  
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
                    logger.info(f"🔼 {exam_no} - {code}: {original_score} → 50")
  
    if upgraded_count > 0:
        logger.info(f"✅ Upgraded {upgraded_count} scores from {min_threshold}–49 to 50")
        logger.info(f"📊 Affected {len(upgraded_students)} students")
    else:
        logger.info(f"ℹ️ No scores found in range {min_threshold}–49 to upgrade")
  
    return mastersheet, upgraded_count
# ----------------------------
# Semester Display Info
# ----------------------------
def get_semester_display_info(semester_key):
    """Get display information for a given semester key.
    Returns: (year, semester_num, level_display, semester_display, set_code)"""
    semester_lower = semester_key.lower()
    if 'first-year-first-semester' in semester_lower:
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "BN1"
    elif 'first-year-second-semester' in semester_lower:
        return 1, 2, "YEAR ONE", "SECOND SEMESTER", "BN1"
    elif 'second-year-first-semester' in semester_lower:
        return 2, 1, "YEAR TWO", "FIRST SEMESTER", "BN2"
    elif 'second-year-second-semester' in semester_lower:
        return 2, 2, "YEAR TWO", "SECOND SEMESTER", "BN2"
    elif 'third-year-first-semester' in semester_lower:
        return 3, 1, "YEAR THREE", "FIRST SEMESTER", "BN3"
    elif 'third-year-second-semester' in semester_lower:
        return 3, 2, "YEAR THREE", "SECOND SEMESTER", "BN3"
    elif 'first' in semester_lower and 'second' not in semester_lower:
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "BN1"
    elif 'second' in semester_lower and 'third' not in semester_lower:
        return 1, 2, "YEAR ONE", "SECOND SEMESTER", "BN1"
    elif 'third' in semester_lower:
        return 3, 1, "YEAR THREE", "FIRST SEMESTER", "BN3"
    else:
        # Default to first semester, first year
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "BN1"
# ----------------------------
# GPA Loading Functions - FIXED VERSION
# ----------------------------
def load_previous_gpas_from_processed_files(output_dir, current_semester_key, timestamp):
    """
    Load previous GPA data from previously processed mastersheets - FIXED VERSION.
    Returns dict: {exam_number: previous_gpa}
    """
    previous_gpas = {}
    logger.info(f"\n🔍 LOADING PREVIOUS GPA for: {current_semester_key}")
    # Determine previous semester based on current
    current_year, current_semester_num, _, _, _ = get_semester_display_info(current_semester_key)
    # Map current semester to previous semester
    semester_sequence = {
        (1, 1): None, # First semester of first year - no previous GPA
        (1, 2): "N-FIRST-YEAR-FIRST-SEMESTER",
        (2, 1): "N-FIRST-YEAR-SECOND-SEMESTER",
        (2, 2): "N-SECOND-YEAR-FIRST-SEMESTER",
        (3, 1): "N-SECOND-YEAR-SECOND-SEMESTER",
        (3, 2): "N-THIRD-YEAR-FIRST-SEMESTER"
    }
    prev_semester = semester_sequence.get((current_year, current_semester_num))
    if not prev_semester:
        logger.info("📊 First semester of first year - no previous GPA available")
        return previous_gpas
    logger.info(f"🔍 Looking for previous GPA data from: {prev_semester}")
    # CRITICAL FIX: Look in the output directory directly for mastersheet
    mastersheet_path = os.path.join(output_dir, f"mastersheet_{timestamp}.xlsx")
    if not os.path.exists(mastersheet_path):
        logger.error(f"❌ Mastersheet not found: {mastersheet_path}")
        return previous_gpas
    try:
        # Read the Excel file properly, skipping the header rows
        df = pd.read_excel(mastersheet_path, sheet_name=prev_semester, header=5)
        logger.info(f"📋 Columns in {prev_semester}: {df.columns.tolist()}")
        # Find the actual column names
        exam_col = None
        gpa_col = None
        for col in df.columns:
            col_str = str(col).upper().strip()
            if 'EXAM' in col_str or 'REG' in col_str or 'NUMBER' in col_str:
                exam_col = col
            elif 'GPA' in col_str:
                gpa_col = col
                break
        if exam_col and gpa_col:
            logger.info(f"✅ Found exam column: '{exam_col}', GPA column: '{gpa_col}'")
            gpas_loaded = 0
            for idx, row in df.iterrows():
                exam_no = str(row[exam_col]).strip()
                gpa = row[gpa_col]
                if pd.notna(gpa) and pd.notna(exam_no) and exam_no != 'nan' and exam_no != '':
                    try:
                        previous_gpas[exam_no] = float(gpa)
                        gpas_loaded += 1
                        if gpas_loaded <= 5:
                            logger.info(f"📝 Loaded GPA: {exam_no} → {gpa}")
                    except (ValueError, TypeError):
                        continue
            logger.info(f"✅ Loaded previous GPAs for {gpas_loaded} students from {prev_semester}")
        else:
            logger.error(f"❌ Could not find required columns in {prev_semester}")
    except Exception as e:
        logger.warning(f"⚠️ Could not read mastersheet: {str(e)}")
        traceback.print_exc()
    logger.info(f"📊 FINAL: Loaded {len(previous_gpas)} previous GPAs")
    return previous_gpas
def load_all_previous_gpas_for_cgpa(output_dir, current_semester_key, timestamp):
    """
    Load ALL previous GPAs from all completed semesters for CGPA calculation.
    Returns dict: {exam_number: {'gpas': [gpa1, gpa2, ...], 'credits': [credits1, credits2, ...]}}
    """
    logger.info(f"\n🔍 LOADING ALL PREVIOUS GPAs for CGPA calculation: {current_semester_key}")
    current_year, current_semester_num, _, _, _ = get_semester_display_info(current_semester_key)
    # Determine which semesters to load based on current semester
    semesters_to_load = []
    if current_semester_num == 1 and current_year == 1:
        # First semester - no previous data
        return {}
    elif current_semester_num == 2 and current_year == 1:
        # Second semester of first year - load first semester
        semesters_to_load = ["N-FIRST-YEAR-FIRST-SEMESTER"]
    elif current_semester_num == 1 and current_year == 2:
        # First semester of second year - load both first year semesters
        semesters_to_load = [
            "N-FIRST-YEAR-FIRST-SEMESTER",
            "N-FIRST-YEAR-SECOND-SEMESTER"]
    elif current_semester_num == 2 and current_year == 2:
        # Second semester of second year - load all previous semesters
        semesters_to_load = [
            "N-FIRST-YEAR-FIRST-SEMESTER",
            "N-FIRST-YEAR-SECOND-SEMESTER",
            "N-SECOND-YEAR-FIRST-SEMESTER"
        ]
    elif current_semester_num == 1 and current_year == 3:
        # First semester of third year - load all previous semesters
        semesters_to_load = [
            "N-FIRST-YEAR-FIRST-SEMESTER",
            "N-FIRST-YEAR-SECOND-SEMESTER",
            "N-SECOND-YEAR-FIRST-SEMESTER",
            "N-SECOND-YEAR-SECOND-SEMESTER"
        ]
    elif current_semester_num == 2 and current_year == 3:
        # Second semester of third year - load all previous semesters
        semesters_to_load = [
            "N-FIRST-YEAR-FIRST-SEMESTER",
            "N-FIRST-YEAR-SECOND-SEMESTER",
            "N-SECOND-YEAR-FIRST-SEMESTER",
            "N-SECOND-YEAR-SECOND-SEMESTER",
            "N-THIRD-YEAR-FIRST-SEMESTER"
        ]
    logger.info(f"📚 Semesters to load for CGPA: {semesters_to_load}")
    all_student_data = {}
    mastersheet_path = os.path.join(output_dir, "mastersheet_{}.xlsx".format(timestamp))
    if not os.path.exists(mastersheet_path):
        logger.error(f"❌ Mastersheet not found: {mastersheet_path}")
        return {}
    for semester in semesters_to_load:
        logger.info(f"📖 Loading data from: {semester}")
        try:
            # Read the Excel file properly, skipping the header rows
            df = pd.read_excel(mastersheet_path, sheet_name=semester, header=5)
            # Filter valid rows
            exam_col = find_exam_number_column(df)
            if exam_col:
                df = df[df[exam_col].notna() & (df[exam_col] != '')]
            # Find columns
            exam_col = find_exam_number_column(df)
            gpa_col = None
            credit_col = None
            for col in df.columns:
                col_str = str(col).upper().strip()
                if 'GPA' in col_str:
                    gpa_col = col
                elif 'CU PASSED' in col_str or 'CREDIT' in col_str:
                    credit_col = col
            if exam_col and gpa_col:
                for idx, row in df.iterrows():
                    exam_no = str(row[exam_col]).strip()
                    gpa = row[gpa_col]
                    if pd.notna(gpa) and pd.notna(exam_no) and exam_no != 'nan' and exam_no != '':
                        try:
                            # Get credits completed (use CU Passed if available, otherwise estimate)
                            credits_completed = 0
                            if credit_col and pd.notna(row[credit_col]):
                                credits_completed = int(row[credit_col])
                            else:
                                # Estimate credits based on typical semester load
                                if 'FIRST-YEAR-FIRST-SEMESTER' in semester:
                                    credits_completed = 30 # Typical first semester credits
                                elif 'FIRST-YEAR-SECOND-SEMESTER' in semester:
                                    credits_completed = 30 # Typical second semester credits
                                elif 'SECOND-YEAR-FIRST-SEMESTER' in semester:
                                    credits_completed = 30 # Typical third semester credits
                                elif 'SECOND-YEAR-SECOND-SEMESTER' in semester:
                                    credits_completed = 30 # Typical fourth semester credits
                                elif 'THIRD-YEAR-FIRST-SEMESTER' in semester:
                                    credits_completed = 30 # Typical fifth semester credits
                            if exam_no not in all_student_data:
                                all_student_data[exam_no] = {'gpas': [], 'credits': []}
                            all_student_data[exam_no]['gpas'].append(float(gpa))
                            all_student_data[exam_no]['credits'].append(credits_completed)
                        except (ValueError, TypeError):
                            continue
        except Exception as e:
            logger.warning(f"⚠️ Could not load data from {semester}: {str(e)}")
    logger.info(f"📊 Loaded cumulative data for {len(all_student_data)} students")
    return all_student_data
def calculate_cgpa(student_data, current_gpa, current_credits):
    """Calculate Cumulative GPA (CGPA) based on all previous semesters and current semester."""
    if not student_data:
        return current_gpa
    total_grade_points = 0.0
    total_credits = 0
    # Add previous semesters
    for prev_gpa, prev_credits in zip(student_data['gpas'], student_data['credits']):
        total_grade_points += prev_gpa * prev_credits
        total_credits += prev_credits
    # Add current semester
    total_grade_points += current_gpa * current_credits
    total_credits += current_credits
    if total_credits > 0:
        return round(total_grade_points / total_credits, 2)
    else:
        return current_gpa
def determine_student_status(row, total_cu, pass_threshold):
    """Determine student status based on performance metrics.
    Returns: 'Pass', 'Carry Over', 'Probation', or 'Withdrawn'"""
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
    """Format failed courses remark with line breaks for long lists.
    Returns list of formatted lines."""
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
        elif len(current_line) + len(course) + 2 <= max_line_length: # +2 for ", "
            current_line += ", " + course
        else:
            lines.append(current_line)
            current_line = course
    if current_line:
        lines.append(current_line)
    return lines
def get_user_semester_choice():
    """Prompt user to choose which semesters to process.
    Returns list of semester keys to process."""
    print("\n🎯 SEMESTER PROCESSING OPTIONS:")
    print("1. Process ALL semesters in order")
    for i, semester in enumerate(BN_SEMESTER_ORDER, 2):
        year, sem_num, level, sem_display, set_code = get_semester_display_info(semester)
        print("{}. Process {} - {} only".format(i, level, sem_display))
    print("{}. Custom selection".format(len(BN_SEMESTER_ORDER) + 2))
    while True:
        try:
            choice = input("\nEnter your choice (1-{}): ".format(len(BN_SEMESTER_ORDER) + 2)).strip()
            if choice == "1":
                return BN_SEMESTER_ORDER.copy()
            elif choice.isdigit():
                choice_num = int(choice)
                if 2 <= choice_num <= len(BN_SEMESTER_ORDER) + 1:
                    return [BN_SEMESTER_ORDER[choice_num - 2]]
                elif choice_num == len(BN_SEMESTER_ORDER) + 2:
                    return get_custom_semester_selection()
                else:
                    print("❌ Invalid choice. Please enter a number between 1-{}.".format(len(BN_SEMESTER_ORDER) + 2))
            else:
                print("❌ Please enter a valid number.")
        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)
        except Exception as e:
            print("❌ Error: {}. Please try again.".format(e))
def get_custom_semester_selection():
    """Allow user to select multiple semesters for processing."""
    print("\n📚 AVAILABLE SEMESTERS:")
    for i, semester in enumerate(BN_SEMESTER_ORDER, 1):
        year, sem_num, level, sem_display, set_code = get_semester_display_info(semester)
        print("{}. {} - {}".format(i, level, sem_display))
    print("{}. Select all".format(len(BN_SEMESTER_ORDER) + 1))
    selected = []
    while True:
        try:
            choices = input(
                "\nEnter semester numbers separated by commas (1-{}): ".format(len(BN_SEMESTER_ORDER) + 1)).strip()
            if not choices:
                print("❌ Please enter at least one semester number.")
                continue
            choice_list = [c.strip() for c in choices.split(',')]
            # Check for "select all" option
            if str(len(BN_SEMESTER_ORDER) + 1) in choice_list:
                return BN_SEMESTER_ORDER.copy()
            # Validate and convert choices
            valid_choices = []
            for choice in choice_list:
                if not choice.isdigit():
                    print("❌ '{}' is not a valid number.".format(choice))
                    continue
                choice_num = int(choice)
                if 1 <= choice_num <= len(BN_SEMESTER_ORDER):
                    valid_choices.append(choice_num)
                else:
                    print("❌ '{}' is not a valid semester number.".format(choice))
            if valid_choices:
                selected_semesters = [BN_SEMESTER_ORDER[i - 1] for i in valid_choices]
                print("✅ Selected semesters: {}".format([get_semester_display_info(sem)[3] for sem in selected_semesters]))
                return selected_semesters
            else:
                print("❌ No valid semesters selected. Please try again.")
        except KeyboardInterrupt:
            print("\n👋 Operation cancelled by user.")
            sys.exit(0)
        except Exception as e:
            print("❌ Error: {}. Please try again.".format(e))
# ----------------------------
# PDF Generation - Individual Student Report - FIXED VERSION
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
    upgrade_min_threshold=None):
    """Create a PDF with one page per student matching the sample format exactly."""
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
                logo_img = Image(logo_path, width=0.8 * inch, height=0.8 * inch)
            except Exception as e:
                logger.warning(f"Could not load logo: {e}")
        # Header table with logo and title
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
        # Passport photo table (separate box)
        passport_data = [
            [Paragraph("Affix Recent<br/>Passport<br/>Photograph", styles['Normal'])]
        ]
        passport_table = Table(passport_data, colWidths=[1.5 * inch], rowHeights=[1.2 * inch])
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
        combined_table = Table(combined_data, colWidths=[5.8 * inch, 1.5 * inch])
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
                # Apply upgrade rule for PDF display consistency
                if upgrade_min_threshold is not None and upgrade_min_threshold <= score_val <= 49:
                    score_val = 50.0
                    score_display = "50"
                else:
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
        # Calculate current semester GPA
        current_gpa = round(total_grade_points / total_units, 2) if total_units > 0 else 0.0
        # Get previous GPA if available - FIXED: Ensure we're using the correct exam number
        exam_no_clean = str(r.get("EXAMS NUMBER", "")).strip()
        previous_gpa = previous_gpas.get(exam_no_clean, None) if previous_gpas else None
        # Calculate CGPA if available
        cgpa = None
        if cgpa_data and exam_no_clean in cgpa_data:
            cgpa = calculate_cgpa(cgpa_data[exam_no_clean], current_gpa, total_units_passed)
        # Get values from dataframe
        tcpe = round(total_grade_points, 1)
        tcup = total_units_passed
        tcuf = total_units_failed
        # Determine student status based on performance
        student_status = determine_student_status(r, total_cu, pass_threshold)
        # Check if student was previously withdrawn
        withdrawal_history = get_withdrawal_history(exam_no_clean)
        previously_withdrawn = withdrawal_history is not None
        # Format failed courses with line breaks if needed
        failed_courses_formatted = format_failed_courses_remark(failed_courses_list)
        # Combine course-specific remarks with overall status
        final_remarks_lines = []
        if previously_withdrawn and withdrawal_history['withdrawn_semester'] == semester_key:
            if failed_courses_formatted:
                final_remarks_lines.append("Failed: {}".format(failed_courses_formatted[0]))
                if len(failed_courses_formatted) > 1:
                    final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("Advised to Withdraw")
            else:
                final_remarks_lines.append("Advised to Withdraw")
        elif previously_withdrawn:
            withdrawn_semester = withdrawal_history['withdrawn_semester']
            year, sem_num, level, sem_display, set_code = get_semester_display_info(withdrawn_semester)
            final_remarks_lines.append("STUDENT WAS WITHDRAWN FROM {} - {}".format(level, sem_display))
            final_remarks_lines.append("This result should not be processed as student was previously withdrawn")
        elif student_status == "Pass":
            final_remarks_lines.append("Passed")
        elif student_status == "Carry Over":
            if failed_courses_formatted:
                final_remarks_lines.append("Failed: {}".format(failed_courses_formatted[0]))
                if len(failed_courses_formatted) > 1:
                    final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("To Carry Over Courses")
            else:
                final_remarks_lines.append("To Carry Over Courses")
        elif student_status == "Probation":
            if failed_courses_formatted:
                final_remarks_lines.append("Failed: {}".format(failed_courses_formatted[0]))
                if len(failed_courses_formatted) > 1:
                    final_remarks_lines.extend(failed_courses_formatted[1:])
                final_remarks_lines.append("Placed on Probation")
            else:
                final_remarks_lines.append("Placed on Probation")
        elif student_status == "Withdrawn":
            if failed_courses_formatted:
                final_remarks_lines.append("Failed: {}".format(failed_courses_formatted[0]))
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
        # Summary section - MODIFIED: Always include previous GPA and overall GPA, with 'N/A' if not available
        summary_data = [
            [Paragraph("<b>SUMMARY</b>", styles['Normal']), "", "", ""],
            [Paragraph("<b>TCPE:</b>", styles['Normal']), str(tcpe),
             Paragraph("<b>CURRENT GPA:</b>", styles['Normal']), str(display_gpa)],
            [Paragraph("<b>TCUP:</b>", styles['Normal']), str(tcup),
             Paragraph("<b>PREVIOUS GPA:</b>", styles['Normal']), str(previous_gpa if previous_gpa is not None else 'N/A')],
            [Paragraph("<b>TCUF:</b>", styles['Normal']), str(tcuf),
             Paragraph("<b>OVERALL GPA:</b>", styles['Normal']), str(display_cgpa)],
        ]
        # Add remarks with multiple lines if needed
        remarks_paragraph = Paragraph(final_remarks, remarks_style)
        summary_data.append([
            Paragraph("<b>REMARKS:</b>", styles['Normal']),
            remarks_paragraph, "", ""
        ])
        # Calculate row heights based on content
        row_heights = [0.3 * inch] * len(summary_data)
        # Adjust height for remarks row based on number of lines
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
        # Signature section
        sig_data = [ ["", ""],
                    ["____________________", "____________________"],
                    [Paragraph("<b>EXAMS SECRETARY</b>", ParagraphStyle('SigStyle', parent=styles['Normal'], fontSize=10, alignment=TA_CENTER)),
                     Paragraph("<b>V.P. ACADEMICS</b>", ParagraphStyle('SigStyle', parent=styles['Normal'], fontSize=10, alignment=TA_CENTER))] ]
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
    logger.info(f"✅ Individual student PDF written: {out_pdf_path}")
# ----------------------------
# Main BN Processing Functions - FIXED VERSION
# ----------------------------
def process_bn_semester_files(
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
    """Process all files for a specific BN semester."""
    logger.info(f"\n{'='*60}")
    logger.info(f"PROCESSING BN SEMESTER: {semester_key}")
    logger.info(f"{'='*60}")
    # Filter files for this semester
    semester_files = []
    for rf in raw_files:
        try:
            detected_sem = detect_bn_semester_from_filename(rf)
            if detected_sem == semester_key:
                semester_files.append(rf)
        except ValueError as e:
            logger.warning(f"⚠️ Could not detect semester for {rf}: {e}")
            continue
    if not semester_files:
        logger.warning(f"⚠️ No files found for semester {semester_key}")
        return {
            'success': False,
            'files_processed': 0,
            'error': 'No files found for semester'
        }
    logger.info(f"📁 Found {len(semester_files)} files for {semester_key}: {semester_files}")
    files_processed = 0
    # Process each file for this semester
    for rf in semester_files:
        raw_path = os.path.join(raw_dir, rf)
        logger.info(f"\n📄 Processing: {rf}")
        try:
            # Enhanced file validation
            is_valid, validation_msg = validate_raw_file(raw_path)
            if not is_valid:
                logger.error(f"❌ File validation failed for {rf}: {validation_msg}")
                continue
            # Load previous GPAs for this specific semester - FIXED: Pass correct parameters
            current_previous_gpas = load_previous_gpas_from_processed_files(output_dir, semester_key, ts)
            # Load CGPA data (all previous semesters)
            cgpa_data = load_all_previous_gpas_for_cgpa(output_dir, semester_key, ts)
            # Process the file - FIXED: Pass the loaded GPAs correctly
            result = process_bn_single_file(
                raw_path,
                raw_dir,
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
                previous_gpas=current_previous_gpas, # FIXED: Pass the loaded previous GPAs
                cgpa_data=cgpa_data,
                upgrade_min_threshold=upgrade_min_threshold
            )
            if result is not None:
                logger.info(f"✅ Successfully processed {rf}")
                files_processed += 1
            else:
                logger.error(f"❌ Failed to process {rf}")
        except Exception as e:
            logger.error(f"❌ Error processing {rf}: {e}")
            logger.info("🔍 Detailed error traceback:")
            traceback.print_exc()
            continue # Continue with next file instead of stopping entirely
  
    # Return result with proper structure
    if files_processed > 0:
        # Create CGPA summary after processing all files
        mastersheet_path = os.path.join(output_dir, "mastersheet_{}.xlsx".format(ts))
        if os.path.exists(mastersheet_path):
            try:
                create_bn_cgpa_summary_sheet(mastersheet_path, ts)
                create_bn_analysis_sheet(mastersheet_path, ts, semester_credit_units)
            except Exception as e:
                logger.warning(f"⚠️ Could not create summary sheets: {e}")
      
        return {
            'success': True,
            'files_processed': files_processed,
            'semester': semester_key
        }
    else:
        return {
            'success': False,
            'files_processed': 0,
            'error': 'No files processed'
        }
def process_bn_single_file(
        path,
        raw_dir, # Added
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
        previous_gpas=None,
        cgpa_data=None,
        upgrade_min_threshold=None,
        is_resit=False):
    """Process a single BN raw file with all enhanced features."""
    fname = os.path.basename(path)
  
    # Enhanced file validation
    is_valid, validation_msg = validate_raw_file(path)
    if not is_valid:
        logger.error(f"❌ File validation failed: {validation_msg}")
        return None
      
    logger.info(f"🔍 Processing BN file: {fname} for semester: {semester_key}")
    try:
        xl = pd.ExcelFile(path)
        logger.info(f"✅ Successfully opened BN Excel file: {fname}")
        logger.info(f"📋 Sheets found: {xl.sheet_names}")
      
        # Check if file has any sheets
        if not xl.sheet_names:
            logger.error(f"❌ Excel file has no sheets: {path}")
            return None
          
    except Exception as e:
        logger.error(f"❌ Error opening BN excel {path}: {e}")
        return None
    expected_sheets = ['CA', 'OBJ', 'EXAM']
    dfs = {}
  
    for s in expected_sheets:
        if s in xl.sheet_names:
            try:
                dfs[s] = pd.read_excel(path, sheet_name=s, dtype=str, header=0)
                logger.info(f"✅ Loaded BN sheet {s} with shape: {dfs[s].shape}")
              
                # Check if data is in transposed format and transform if needed
                if detect_data_format(dfs[s], s):
                    logger.info(f"🔄 BN Data in {s} sheet is in transposed format, transforming...")
                    dfs[s] = transform_transposed_data(dfs[s], s)
                    logger.info(f"✅ Transformed BN {s} sheet to wide format")
                  
            except Exception as e:
                logger.error(f"❌ Error reading BN sheet {s}: {e}")
                dfs[s] = pd.DataFrame()
        else:
            logger.warning(f"⚠️ BN Sheet {s} not found in {fname}")
            dfs[s] = pd.DataFrame()
          
    if not dfs:
        logger.error("No CA/OBJ/EXAM sheets detected — skipping file.")
        return None
    # Use the provided semester key
    sem = semester_key
    year, semester_num, level_display, semester_display, set_code = get_semester_display_info(sem)
    logger.info(f"📁 Processing: {level_display} - {semester_display} - Set: {set_code}")
    logger.info(f"📊 Using course sheet: {sem}")
    logger.info(f"📊 Previous GPAs provided: {len(previous_gpas) if previous_gpas else 0} students")
    logger.info(f"📊 CGPA data available for: {len(cgpa_data) if cgpa_data else 0} students")
    # Check if semester exists in course maps
    if sem not in semester_course_maps:
        logger.error(f"❌ Semester '{sem}' not found in course data. Available semesters: {list(semester_course_maps.keys())}")
        return None
    course_map = semester_course_maps[sem]
    credit_units = semester_credit_units[sem]
    course_titles = semester_course_titles[sem]
    # Extract ordered codes from enhanced course map
    ordered_titles = list(course_map.keys())
    ordered_codes = [course_map[t]['code'] for t in ordered_titles if course_map.get(t)]
    ordered_codes = [c for c in ordered_codes if credit_units.get(c, 0) > 0]
    filtered_credit_units = {c: credit_units[c] for c in ordered_codes}
    total_cu = sum(filtered_credit_units.values())
    reg_no_cols = {s: find_column_by_names(df, ["REG. No", "Reg No", "Registration Number", "EXAM NUMBER"]) for s, df in dfs.items()}
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
            logger.warning(f"Skipping sheet {s}: no reg column found")
            continue
        df["REG. No"] = df[regcol].astype(str).str.strip()
        if namecol:
            df["NAME"] = df[namecol].astype(str).str.strip()
        else:
            df["NAME"] = pd.NA
        to_drop = [c for c in [regcol, namecol] if c and c not in ["REG. No", "NAME"]]
        df.drop(columns=to_drop, errors="ignore", inplace=True)
        # Enhanced course matching using the new algorithm
        for col in [c for c in df.columns if c not in ["REG. No", "NAME"]]:
            best_match = find_best_course_match(col, course_map)
            if best_match:
                matched_code = best_match['code']
                newcol = "{}_{}".format(matched_code, s.upper())
                df.rename(columns={col: newcol}, inplace=True)
                logger.info(f"✅ Matched '{col}' -> '{matched_code}'")
        cur_cols = ["REG. No", "NAME"] + [c for c in df.columns if c.endswith("_{}".format(s.upper()))]
        cur = df[cur_cols].copy()
        if merged is None:
            merged = cur
        else:
            merged = merged.merge(cur, on="REG. No", how="outer", suffixes=('', '_dup'))
            if "NAME_dup" in merged.columns:
                merged["NAME"] = merged["NAME"].combine_first(merged["NAME_dup"])
                merged.drop(columns=["NAME_dup"], inplace=True)
    if merged is None or merged.empty:
        logger.error("No data merged from sheets — skipping file.")
        return None
    mastersheet = merged[["REG. No", "NAME"]].copy()
    mastersheet.rename(columns={"REG. No": "EXAMS NUMBER"}, inplace=True)
    for code in ordered_codes:
        ca_col = "{}_CA".format(code)
        obj_col = "{}_OBJ".format(code)
        exam_col = "{}_EXAM".format(code)
        ca_series = pd.to_numeric(merged[ca_col], errors="coerce") if ca_col in merged.columns else pd.Series([0] * len(merged), index=merged.index)
        obj_series = pd.to_numeric(merged[obj_col], errors="coerce") if obj_col in merged.columns else pd.Series([0] * len(merged), index=merged.index)
        exam_series = pd.to_numeric(merged[exam_col], errors="coerce") if exam_col in merged.columns else pd.Series([0] * len(merged), index=merged.index)
        ca_norm = (ca_series / 20) * 100
        obj_norm = (obj_series / 20) * 100
        exam_norm = (exam_series / 80) * 100
        ca_norm = ca_norm.fillna(0).clip(upper=100)
        obj_norm = obj_norm.fillna(0).clip(upper=100)
        exam_norm = exam_norm.fillna(0).clip(upper=100)
        total = (ca_norm * 0.2) + (((obj_norm + exam_norm) / 2) * 0.8)
        mastersheet[code] = total.round(0).clip(upper=100).values
    # APPLY FLEXIBLE UPGRADE RULE
    if should_use_interactive_mode():
        upgrade_min_threshold, upgraded_scores_count = get_upgrade_threshold_from_user(semester_key, set_name)
    else:
        # In non-interactive mode, use the provided threshold or None
        upgraded_scores_count = 0
        if upgrade_min_threshold is not None:
            logger.info(f"🔄 Applying upgrade rule from parameters: {upgrade_min_threshold}–49 → 50")
  
    if upgrade_min_threshold is not None:
        mastersheet, upgraded_scores_count = apply_upgrade_rule(mastersheet, ordered_codes, upgrade_min_threshold)
    for c in ordered_codes:
        if c not in mastersheet.columns:
            mastersheet[c] = 0
    # (RE)CALCULATE REMARKS AND METRICS AFTER POSSIBLE UPDATES
    def compute_remarks(row):
        """Compute remarks with expanded failed courses list."""
        fails = [c for c in ordered_codes if float(row.get(c, 0) or 0) < pass_threshold]
        if not fails:
            return "Passed"
        failed_courses_str = ", ".join(sorted(fails))
        return "Failed: {}".format(failed_courses_str)
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
    # Calculate GPA
    def calculate_gpa(row):
        row_tcpe = row["TCPE"]
        return round((row_tcpe / total_cu), 2) if total_cu > 0 else 0.0
    mastersheet["GPA"] = mastersheet.apply(calculate_gpa, axis=1)
    mastersheet["AVERAGE"] = mastersheet[[c for c in ordered_codes]].mean(axis=1).round(0)
    # FILTER OUT PREVIOUSLY WITHDRAWN STUDENTS
    mastersheet, removed_students = filter_out_withdrawn_students(mastersheet, semester_key)
    # Identify withdrawn students in this semester (after filtering)
    withdrawn_students = []
    for idx, row in mastersheet.iterrows():
        student_status = determine_student_status(row, total_cu, pass_threshold)
        if student_status == "Withdrawn":
            exam_no = str(row["EXAMS NUMBER"]).strip()
            withdrawn_students.append(exam_no)
            mark_student_withdrawn(exam_no, semester_key)
            logger.info(f"🚫 Student {exam_no} marked as withdrawn in {semester_key}")
    # Update student tracker with current semester's students (after filtering)
    exam_numbers = mastersheet["EXAMS NUMBER"].astype(str).str.strip().tolist()
    update_student_tracker(semester_key, exam_numbers, withdrawn_students)
    # IDENTIFY CARRYOVER STUDENTS - ENSURE THIS IS CALLED
    carryover_students = identify_carryover_students(mastersheet, semester_key, set_name, pass_threshold)
    if carryover_students:
        carryover_dir = save_carryover_records(carryover_students, output_dir, set_name, semester_key)
        logger.info(f"✅ Saved {len(carryover_students)} BN carryover records to: {carryover_dir}")
    else:
        logger.info("✅ No carryover students identified for this semester")
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
    if is_resit:
        # For resit processing, skip saving and PDF generation
        logger.info("ℹ️ Skipping save and PDF for resit data")
        return mastersheet
    # Create proper output directory structure
    output_subdir = output_dir # Use the provided output_dir directly
    os.makedirs(output_subdir, exist_ok=True)
    out_xlsx = os.path.join(output_subdir, "mastersheet_{}.xlsx".format(ts))
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
            img.anchor = 'A1'
            ws.add_image(img, "A1")
        except Exception as e:
            logger.warning(f"⚠ Could not place logo: {e}")
    # FIXED MERGE RANGE
    last_letter = get_column_letter(len(out_cols))
    ws.merge_cells(f"C1:{last_letter}1")
    title_cell = ws["C1"]
    title_cell.value = "FCT COLLEGE OF NURSING SCIENCES, GWAGWALADA-ABUJA"
    title_cell.font = Font(bold=True, size=16, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    border = Border(left=Side(style="medium"), right=Side(style="medium"), top=Side(style="medium"), bottom=Side(style="medium"))
    title_cell.border = border
    # Use expanded semester name in the subtitle
    expanded_semester_name = "{} {}".format(level_display, semester_display)
    ws.merge_cells(f"C2:{last_letter}2")
    subtitle_cell = ws["C2"]
    subtitle_cell.value = "{}/{} SESSION BASIC NURSING {} EXAMINATIONS RESULT — {}".format(datetime.now().year, datetime.now().year + 1, expanded_semester_name, datetime.now().strftime('%B %d, %Y'))
    subtitle_cell.font = Font(bold=True, size=12, color="000000")
    subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")
    start_row = 3
    display_course_titles = []
    for t, c in zip(ordered_titles, [course_map[t]['code'] for t in ordered_titles]):
        if c in ordered_codes:
            display_course_titles.append(course_map[t]['original_name'])
    ws.append([""] * 3 + display_course_titles + [""] * 5)
    for i, cell in enumerate(ws[start_row][3:3 + len(display_course_titles)], start=3):
        cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=45)
        cell.font = Font(bold=True, size=9)
    ws.row_dimensions[start_row].height = 18
    cu_list = [filtered_credit_units.get(c, "") for c in ordered_codes]
    ws.append([""] * 3 + cu_list + [""] * 5)
    for cell in ws[start_row + 1][3:3 + len(cu_list)]:
        cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=135)
        cell.font = Font(bold=True, size=9)
        cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    headers = out_cols
    ws.append(headers)
    for cell in ws[start_row + 2]:
        cell.font = Font(bold=True, size=10, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for _, r in mastersheet.iterrows():
        rowvals = [r[col] for col in headers]
        ws.append(rowvals)
    # Freeze the column headers
    ws.freeze_panes = ws.cell(row=start_row + 3, column=1)
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for row in ws.iter_rows(min_row=start_row + 3, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
    # Colorize course columns - SPECIAL COLOR FOR UPGRADED SCORES
    upgraded_fill = PatternFill(start_color="E6FFCC", end_color="E6FFCC", fill_type="solid")
    passed_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    failed_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    for idx, code in enumerate(ordered_codes, start=4):
        col_letter = get_column_letter(idx)
        for r_idx in range(start_row + 3, ws.max_row + 1):
            cell = ws.cell(row=r_idx, column=idx)
            try:
                val = float(cell.value) if cell.value not in (None, "") else 0
                if upgrade_min_threshold is not None and upgrade_min_threshold <= val <= 49:
                    # This score was upgraded - use special color
                    cell.fill = upgraded_fill
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
            for row_idx in range(start_row + 3, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(horizontal="left", vertical="center")
        # Center align S/N column
        elif col_name == "S/N":
            col_letter = get_column_letter(col_idx)
            for row_idx in range(start_row + 3, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(horizontal="center", vertical="center")
    # Calculate optimal column widths
    longest_name_len = max([len(str(x)) for x in mastersheet["NAME"].fillna("")]) if "NAME" in mastersheet.columns else 10
    name_col_width = min(max(longest_name_len + 2, 10), NAME_WIDTH_CAP)
    # Enhanced REMARKS column width calculation
    longest_remark_len = 0
    for remark in mastersheet["REMARKS"].fillna(""):
        remark_str = str(remark)
        if remark_str.startswith("Failed:"):
            failed_courses = remark_str.replace("Failed: ", "")
            failed_length = len(failed_courses)
            total_length = failed_length + 15
        else:
            total_length = len(remark_str)
        if total_length > longest_remark_len:
            longest_remark_len = total_length
    remarks_col_width = min(max(longest_remark_len + 4, 40), 80)
    # Apply text wrapping and left alignment to REMARKS column
    remarks_col_idx = headers.index("REMARKS") + 1
    for row_idx in range(start_row + 3, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=remarks_col_idx)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for col_idx in range(1, ws.max_column + 1):
        column_letter = get_column_letter(col_idx)
        if column_letter == "A": # S/N
            ws.column_dimensions[column_letter].width = 6
        elif column_letter == "B" or headers[col_idx - 1] in ["EXAMS NUMBER", "EXAM NO"]:
            ws.column_dimensions[column_letter].width = 18
        elif headers[col_idx - 1] == "NAME":
            ws.column_dimensions[column_letter].width = name_col_width
        elif 4 <= col_idx < 4 + len(ordered_codes): # course columns
            ws.column_dimensions[column_letter].width = 8
        elif headers[col_idx - 1] in ["REMARKS"]:
            ws.column_dimensions[column_letter].width = remarks_col_width
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
    ws.append([f"A total of {total_students} students registered and sat for the Examination"])
    ws.append([f"A total of {passed_all} students passed in all courses registered and are to proceed to Second Semester, BN 1"])
    ws.append([f"A total of {gpa_above_2_failed} students with Grade Point Average (GPA) of 2.00 and above failed various courses, but passed at least 45% of the total registered credit units, and are to carry these courses over to the next session."])
    ws.append([f"A total of {gpa_below_2_failed} students with Grade Point Average (GPA) below 2.00 failed various courses, but passed at least 45% of the total registered credit units, and are placed on Probation, to carry these courses over to the next session."])
    ws.append([f"A total of {failed_over_45_percent} students failed in more than 45% of their registered credit units in various courses and have been advised to withdraw"])
    # Add upgrade notice in summary section
    if upgrade_min_threshold is not None:
        ws.append([f"✅ Upgraded all scores between {upgrade_min_threshold}–49 to 50 as per management decision ({upgraded_scores_count} scores upgraded)"])
    # Add removed withdrawn students info
    if removed_students:
        ws.append([f"NOTE: {len(removed_students)} previously withdrawn students were removed from this semester's results as they should not be processed."])
    ws.append(["The above decisions are in line with the provisions of the provisions of the General Information Section of the NMCN/NBTE Examinations Regulations (Pg 4) adopted by the College."])
    ws.append([])
    ws.append(["________________________", "", "", "________________________", "", "", "", "", "", "", "", "", ""])
    ws.append(["Mrs. Abini Hauwa", "", "", "Mrs. Olukemi Ogunleye", "", "", "", "", "", "", "", "", ""])
    ws.append(["Head of Exams", "", "", "Chairman, ND/HND Program C'tee", "", "", "", "", "", "", "", "", ""])
    wb.save(out_xlsx)
    logger.info(f"✅ Mastersheet saved: {out_xlsx}")
    # Generate individual student PDF with previous GPAs and CGPA - FIXED: Pass previous_gpas correctly
    safe_sem = re.sub(r'[^\w\-]', '_', sem)
    student_pdf_path = os.path.join(output_subdir, "mastersheet_students_{}_{}.pdf".format(ts, safe_sem))
    logger.info("📊 FINAL CHECK before PDF generation:")
    logger.info(f" Previous GPAs loaded: {len(previous_gpas) if previous_gpas else 0}")
    logger.info(f" CGPA data available for: {len(cgpa_data) if cgpa_data else 0} students")
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
            previous_gpas=previous_gpas, # FIXED: Pass the actual previous_gpas parameter
            cgpa_data=cgpa_data,
            total_cu=total_cu,
            pass_threshold=pass_threshold,
            upgrade_min_threshold=upgrade_min_threshold
        )
        logger.info(f"✅ PDF generated successfully for {sem}")
    except Exception as e:
        logger.error(f"❌ Failed to generate student PDF for {sem}: {e}")
        traceback.print_exc()
    return mastersheet
# ----------------------------
# Main runner
# ----------------------------
def main():
    """Main entry point with comprehensive error handling."""
    try:
        # Initialize
        initialize_student_tracker()
        initialize_carryover_tracker()
      
        logger.info("Starting BN Examination Results Processing with Enhanced Features...")
        ts = datetime.now().strftime(TIMESTAMP_FMT)
        # Check if running in web mode
        if is_web_mode():
            uploaded_file_path = get_uploaded_file_path()
            if uploaded_file_path and os.path.exists(uploaded_file_path):
                logger.info("🔧 Running in WEB MODE with uploaded file")
                success = process_uploaded_file(uploaded_file_path, normalize_path(BASE_DIR))
                if success:
                    logger.info("✅ Uploaded file processing completed successfully")
                else:
                    logger.error("❌ Uploaded file processing failed")
                return
            else:
                logger.error("❌ No uploaded file found in web mode")
                return
        # Get parameters from form
        params = get_form_parameters()
      
        # Use the parameters
        global DEFAULT_PASS_THRESHOLD
        DEFAULT_PASS_THRESHOLD = params['pass_threshold']
      
        base_dir_norm = normalize_path(BASE_DIR)
        logger.info(f"Using base directory: {base_dir_norm}")
        # Check if we should use interactive or non-interactive mode
        if should_use_interactive_mode():
            logger.info("🔧 Running in INTERACTIVE mode (CLI)")
          
            try:
                semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_bn_course_data()
            except Exception as e:
                logger.error(f"❌ Could not load BN course data: {e}")
                return
            # Get available sets and let user choose
            available_sets = get_available_bn_sets(base_dir_norm)
          
            if not available_sets:
                logger.error(f"No BN SET* directories found in {base_dir_norm}. Nothing to process.")
                return
            logger.info(f"📚 Found {len(available_sets)} available BN sets: {available_sets}")
            # Let user choose which set(s) to process
            sets_to_process = get_user_set_choice(available_sets)
            logger.info(f"\n🎯 PROCESSING SELECTED SETS: {sets_to_process}")
            for bn_set in sets_to_process:
                logger.info(f"\n{'='*60}")
                logger.info(f"PROCESSING BN SET: {bn_set}")
                logger.info(f"{'='*60}")
              
                raw_dir = normalize_path(os.path.join(base_dir_norm, "BN", bn_set, "RAW_RESULTS"))
                clean_dir = normalize_path(os.path.join(base_dir_norm, "BN", bn_set, "CLEAN_RESULTS"))
                # Create directories if they don't exist
                os.makedirs(raw_dir, exist_ok=True)
                os.makedirs(clean_dir, exist_ok=True)
                if not os.path.exists(raw_dir):
                    logger.warning(f"⚠️ BN RAW_RESULTS directory not found: {raw_dir}")
                    continue
                raw_files = [f for f in os.listdir(raw_dir)
                             if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
                if not raw_files:
                    logger.warning(f"⚠️ No raw files in {raw_dir}; skipping {bn_set}")
                    continue
                logger.info(f"📁 Found {len(raw_files)} raw files in {bn_set}: {raw_files}")
                # Create timestamped folder for this set
                set_output_dir = os.path.join(clean_dir, "{}_RESULT-{}".format(bn_set, ts))
                os.makedirs(set_output_dir, exist_ok=True)
                logger.info(f"📁 Created BN set output directory: {set_output_dir}")
                # Get user choice for which semesters to process
                semesters_to_process = get_user_semester_choice()
                logger.info(f"\n🎯 PROCESSING SELECTED SEMESTERS for {bn_set}: {[get_semester_display_info(sem)[3] for sem in semesters_to_process]}")
                # Process selected semesters in the correct order
                semester_processed = 0
                for semester_key in semesters_to_process:
                    if semester_key not in BN_SEMESTER_ORDER:
                        logger.warning(f"⚠️ Skipping unknown semester: {semester_key}")
                        continue
                    # Check if there are files for this semester - USING BM'S APPROACH
                    semester_files_exist = False
                    for rf in raw_files:
                        try:
                            detected_sem = detect_bn_semester_from_filename(rf)
                            if detected_sem == semester_key:
                                semester_files_exist = True
                                break
                        except ValueError as e:
                            logger.warning(f"⚠️ Could not detect semester for {rf}: {e}")
                            continue
                    if semester_files_exist:
                        logger.info(f"\n🎯 Processing BN {semester_key} in {bn_set}...")
                        result = process_bn_semester_files(
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
                            bn_set)
                        if result and result.get('success', False):
                            semester_processed += result.get('files_processed', 0)
                    else:
                        logger.warning(f"⚠️ No files found for BN {semester_key} in {bn_set}, skipping...")
              
                # Create ZIP of BN results ONLY if files were processed
                if semester_processed > 0:
                    try:
                        zip_success = create_bn_zip_for_set(clean_dir, bn_set, ts, set_output_dir)
                        if zip_success:
                            logger.info(f"✅ Successfully created ZIP for {bn_set}")
                        else:
                            logger.warning(f"⚠️ ZIP creation failed for {bn_set}, files remain in: {set_output_dir}")
                    except Exception as e:
                        logger.warning(f"⚠️ Failed to create BN ZIP for {bn_set}: {e}")
                        traceback.print_exc()
                else:
                    logger.warning(f"⚠️ No files processed for {bn_set}, skipping ZIP creation")
                    if os.path.exists(set_output_dir):
                        shutil.rmtree(set_output_dir)
            # Print BN-specific summaries
            logger.info("\n📊 BN STUDENT TRACKING SUMMARY:")
            logger.info(f"Total unique BN students tracked: {len(STUDENT_TRACKER)}")
            logger.info(f"Total BN withdrawn students: {len(WITHDRAWN_STUDENTS)}")
            if CARRYOVER_STUDENTS:
                logger.info("\n📋 BN CARRYOVER STUDENT SUMMARY:")
                logger.info(f"Total BN carryover students: {len(CARRYOVER_STUDENTS)}")
            # Analyze student progression
            sem_counts = {}
            for student_data in STUDENT_TRACKER.values():
                sem_count = len(student_data['semesters_present'])
                if sem_count not in sem_counts:
                    sem_counts[sem_count] = 0
                sem_counts[sem_count] += 1
            for sem_count, student_count in sorted(sem_counts.items()):
                logger.info(f"Students present in {sem_count} semester(s): {student_count}")
            logger.info("\n✅ BN Examination Results Processing completed successfully.")
        else:
            logger.info("🔧 Running in NON-INTERACTIVE mode (Web)")
            success = process_in_non_interactive_mode(params, base_dir_norm)
            if success:
                logger.info("✅ BN Examination Results Processing completed successfully")
            else:
                logger.error("❌ BN Examination Results Processing failed")
            return
          
    except KeyboardInterrupt:
        logger.info("\n👋 Process interrupted by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"❌ Fatal error: {e}")
        traceback.print_exc()
        sys.exit(1)
    finally:
        # Cleanup
        logger.info("\n📊 Final Summary:")
        logger.info(f" Students tracked: {len(STUDENT_TRACKER)}")
        logger.info(f" Carryover students: {len(CARRYOVER_STUDENTS)}")
        logger.info(f" Withdrawn students: {len(WITHDRAWN_STUDENTS)}")
if __name__ == "__main__":
    main()