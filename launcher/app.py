# app.py (Fully Fixed Version)
import os
import subprocess
import re
import sys
import zipfile
import shutil
import time
import socket
import logging
import json
import glob
from pathlib import Path
from datetime import datetime
from flask import (
    Flask,
    request,
    redirect,
    url_for,
    render_template,
    flash,
    session,
    send_file,
    send_from_directory,
    jsonify,
)
from functools import wraps
from dotenv import load_dotenv
from jinja2 import TemplateNotFound
from werkzeug.utils import secure_filename

# Configure logging
logging.basicConfig(
    level=logging.DEBUG, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# Define directory structure relative to project root
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SCRIPT_DIR = os.path.join(PROJECT_ROOT, "scripts")
BASE_DIR = os.getenv("BASE_DIR", "/home/ernest/student_result_cleaner/EXAMS_INTERNAL")

# Launcher-specific directories
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "templates")
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")

# Log paths for verification
logger.info(f"PROJECT_ROOT: {PROJECT_ROOT}")
logger.info(f"SCRIPT_DIR: {SCRIPT_DIR}")
logger.info(f"BASE_DIR (EXAMS_INTERNAL): {BASE_DIR}")
logger.info(f"TEMPLATE_DIR: {TEMPLATE_DIR}")
logger.info(f"STATIC_DIR: {STATIC_DIR}")
logger.info(f"Template dir exists: {os.path.exists(TEMPLATE_DIR)}")
logger.info(f"Static dir exists: {os.path.exists(STATIC_DIR)}")

# Verify templates
if os.path.exists(TEMPLATE_DIR):
    templates = os.listdir(TEMPLATE_DIR)
    logger.info(f"Templates found: {templates}")
    if "login.html" in templates:
        logger.info(f"login.html found in {TEMPLATE_DIR}")
    else:
        logger.warning(f"login.html NOT found in {TEMPLATE_DIR}")
else:
    logger.error(f"Template directory not found: {TEMPLATE_DIR}")

# Initialize Flask with explicit paths
app = Flask(__name__, template_folder=TEMPLATE_DIR, static_folder=STATIC_DIR)
app.logger.setLevel(logging.DEBUG)
app.secret_key = os.getenv("FLASK_SECRET", "default_secret_key_1234567890")

# Configuration
PASSWORD = os.getenv("STUDENT_CLEANER_PASSWORD", "admin")
COLLEGE = os.getenv("COLLEGE_NAME", "FCT College of Nursing Sciences, Gwagwalada")
DEPARTMENT = os.getenv("DEPARTMENT", "Examinations Office")

# Define sets for templates
ND_SETS = ["ND-2024", "ND-2025"]
BN_SETS = ["SET47", "SET48"]
BM_SETS = ["SET2023", "SET2024", "SET2025"]
PROGRAMS = ["ND", "BN", "BM"]

def get_program_from_set(set_name):
    """Determine program from set name"""
    if set_name in ND_SETS:
        return "ND"
    elif set_name in BN_SETS:
        return "BN"
    elif set_name in BM_SETS:
        return "BM"
    return None

# ============================================================================
# NEW: Semester Key Standardization Function
# ============================================================================
def standardize_semester_key(semester_key):
    """Standardize semester key to canonical format."""
    if not semester_key:
        return None
    
    key_upper = semester_key.upper()
    
    # Define canonical mappings
    canonical_mappings = {
        # First Year First Semester variants
        ("FIRST", "YEAR", "FIRST", "SEMESTER"): "ND-First-YEAR-First-SEMESTER",
        ("1ST", "YEAR", "1ST", "SEMESTER"): "ND-First-YEAR-First-SEMESTER",
        ("YEAR", "1", "SEMESTER", "1"): "ND-First-YEAR-First-SEMESTER",
        
        # First Year Second Semester variants
        ("FIRST", "YEAR", "SECOND", "SEMESTER"): "ND-First-YEAR-SECOND-SEMESTER",
        ("1ST", "YEAR", "2ND", "SEMESTER"): "ND-First-YEAR-SECOND-SEMESTER",
        ("YEAR", "1", "SEMESTER", "2"): "ND-First-YEAR-SECOND-SEMESTER",
        
        # Second Year First Semester variants
        ("SECOND", "YEAR", "FIRST", "SEMESTER"): "ND-SECOND-YEAR-First-SEMESTER",
        ("2ND", "YEAR", "1ST", "SEMESTER"): "ND-SECOND-YEAR-First-SEMESTER",
        ("YEAR", "2", "SEMESTER", "1"): "ND-SECOND-YEAR-First-SEMESTER",
        
        # Second Year Second Semester variants
        ("SECOND", "YEAR", "SECOND", "SEMESTER"): "ND-SECOND-YEAR-SECOND-SEMESTER",
        ("2ND", "YEAR", "2ND", "SEMESTER"): "ND-SECOND-YEAR-SECOND-SEMESTER",
        ("YEAR", "2", "SEMESTER", "2"): "ND-SECOND-YEAR-SECOND-SEMESTER",
    }
    
    # Extract key components using regex
    patterns = [
        r'(FIRST|1ST|YEAR.?1).*?(FIRST|1ST|SEMESTER.?1)',
        r'(FIRST|1ST|YEAR.?1).*?(SECOND|2ND|SEMESTER.?2)',
        r'(SECOND|2ND|YEAR.?2).*?(FIRST|1ST|SEMESTER.?1)',
        r'(SECOND|2ND|YEAR.?2).*?(SECOND|2ND|SEMESTER.?2)',
    ]
    
    for pattern_idx, pattern in enumerate(patterns):
        if re.search(pattern, key_upper):
            if pattern_idx == 0:
                return "ND-First-YEAR-First-SEMESTER"
            elif pattern_idx == 1:
                return "ND-First-YEAR-SECOND-SEMESTER"
            elif pattern_idx == 2:
                return "ND-SECOND-YEAR-First-SEMESTER"
            elif pattern_idx == 3:
                return "ND-SECOND-YEAR-SECOND-SEMESTER"
    
    # If no match, return original
    logger.warning(f"Could not standardize semester key: {semester_key}")
    return semester_key

# Jinja2 filters
def datetimeformat(timestamp):
    try:
        dt = datetime.fromtimestamp(timestamp)
        return dt.strftime("%b %d, %Y %I:%M %p")
    except Exception:
        return "Unknown"
app.jinja_env.filters["datetimeformat"] = datetimeformat

def filesizeformat(size):
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"
app.jinja_env.filters["filesizeformat"] = filesizeformat

def is_local_environment():
    try:
        hostname = socket.gethostname()
        ip = socket.gethostbyname(hostname)
        if "railway" in hostname.lower() or ip.startswith("10.") or ip.startswith("172."):
            return False
        return os.path.exists(BASE_DIR)
    except Exception as e:
        logger.error(f"Error checking environment: {e}")
        return True

# Check if BASE_DIR exists
if not os.path.exists(BASE_DIR):
    logger.info(f"Creating BASE_DIR: {BASE_DIR}")
    os.makedirs(BASE_DIR, exist_ok=True)
else:
    logger.info(f"BASE_DIR already exists: {BASE_DIR}")

# Define required subdirectories structure - ONLY create if they don't exist
required_dirs = [
    # ND Structure
    os.path.join(BASE_DIR, "ND", "ND-COURSES"),
    os.path.join(BASE_DIR, "ND", "ND-2024", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "ND", "ND-2024", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "ND", "ND-2025", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "ND", "ND-2025", "CLEAN_RESULTS"),
    
    # BN Structure
    os.path.join(BASE_DIR, "BN", "BN-COURSES"),
    os.path.join(BASE_DIR, "BN", "SET47", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BN", "SET47", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "BN", "SET48", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BN", "SET48", "CLEAN_RESULTS"),
    
    # BM Structure
    os.path.join(BASE_DIR, "BM", "BM-COURSES"),
    os.path.join(BASE_DIR, "BM", "SET2023", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2023", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2024", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2024", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2025", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2025", "CLEAN_RESULTS"),
    
    # Other Results Structure
    os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_PUTME_RESULT"),
    os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT"),
    os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_CANDIDATE_BATCHES"),
    os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_UTME_CANDIDATES"),
    os.path.join(BASE_DIR, "CAOSCE_RESULT", "RAW_CAOSCE_RESULT"),
    os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT"),
    os.path.join(BASE_DIR, "OBJ_RESULT", "RAW_OBJ"),
    os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ"),
    os.path.join(BASE_DIR, "JAMB_DB", "RAW_JAMB_DB"),
    os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB"),
]

# Only create directories that don't exist
for dir_path in required_dirs:
    if not os.path.exists(dir_path):
        try:
            os.makedirs(dir_path, exist_ok=True)
            logger.info(f"Created subdirectory: {dir_path}")
        except Exception as e:
            logger.error(f"Could not create {dir_path}: {e}")
    else:
        logger.info(f"Directory already exists: {dir_path}")

# Script mapping - FIXED: Correct mapping for integrated_carryover_processor
SCRIPT_MAP = {
    "utme": "utme_result.py",
    "caosce": "caosce_result.py",
    "clean": "obj_results.py",
    "split": "split_names.py",
    "exam_processor_nd": "exam_result_processor.py",
    "exam_processor_bn": "exam_processor_bn.py",
    "exam_processor_bm": "exam_processor_bm.py",
    "integrated_carryover_processor": "integrated_carryover_processor.py",  # FIXED: Correct script
}

# Success indicators - Update for integrated_carryover_processor
SUCCESS_INDICATORS = {
    "utme": [
        r"Processing: (PUTME 2025-Batch\d+[A-Z] Post-UTME Quiz-grades\.xlsx)",
        r"Saved processed file: (UTME_RESULT_.*?\.csv)",
        r"Saved processed file: (UTME_RESULT_.*?\.xlsx)",
        r"Saved processed file: (PUTME_COMBINE_RESULT_.*?\.xlsx)",
    ],
    "caosce": [
        r"Processed (CAOSCE SET2023A.*?|VIVA \([0-9]+\)\.xlsx) \(\d+ rows read\)",
        r"Saved processed file: (CAOSCE_RESULT_.*?\.csv)",
    ],
    "clean": [
        r"Processing: (Set2025-.*?\.xlsx|ND\d{4}-SET\d+.*?\.xlsx|.*ND.*SET.*\.xlsx)",
        r"Cleaned CSV saved in.*?cleaned_(Set2025-.*?\.csv|ND\d{4}-SET\d+.*?\.csv)",
        r"Master CSV saved in.*?master_cleaned_results\.csv",
        r"All processing completed successfully!",
    ],
    "split": [r"Saved processed file: (clean_jamb_DB_.*?\.csv)"],
    "exam_processor_nd": [
        r"PROCESSING SEMESTER: (ND-[A-Za-z0-9\s\-]+)",
        r"Successfully processed .*",
        r"Mastersheet saved:.*",
        r"Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"Processing complete",
        r"ND Examination Results Processing completed successfully",
        r"Applying upgrade rule:.*→ 50",
        r"Upgraded \d+ scores from.*to 50",
        r"Identified \d+ carryover students",
        r"Carryover records saved:.*",
        r"Processing resit results for.*",
        r"Updated \d+ scores for \d+ students",
    ],
    "exam_processor_bn": [
        r"PROCESSING SET: (SET47|SET48)",
        r"Successfully processed .*",
        r"Mastersheet saved:.*",
        r"Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"Processing complete",
        r"Basic Nursing Examination Results Processing completed successfully",
        r"Applying upgrade rule:.*→ 50",
        r"Upgraded \d+ scores from.*to 50",
        r"Identified \d+ carryover students",
        r"Carryover records saved:.*",
        r"Processing resit results for.*",
        r"Updated \d+ scores for \d+ students",
    ],
    "exam_processor_bm": [
        r"PROCESSING SET: (SET2023|SET2024|SET2025)",
        r"Successfully processed .*",
        r"Mastersheet saved:.*",
        r"Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"Processing complete",
        r"Basic Midwifery Examination Results Processing completed successfully",
        r"Applying upgrade rule:.*→ 50",
        r"Upgraded \d+ scores from.*to 50",
        r"Identified \d+ carryover students",
        r"Carryover records saved:.*",
        r"Processing resit results for.*",
        r"Updated \d+ scores for \d+ students",
    ],
    "integrated_carryover_processor": [
        r"Updated \d+ scores for \d+ students",
        r"Updated mastersheet saved:.*",
        r"Updated carryover Excel saved:.*",
        r"Updated individual student PDF written:.*",
        r"Carryover processing completed successfully",
        r"Processing resit results for.*",
    ],
}

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv", "zip", "pdf"}

# Helper Functions
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def get_raw_directory(script_name, program=None, set_name=None):
    """Get the RAW_RESULTS directory for a specific script/program/set"""
    logger.info(f"Getting raw directory for: script={script_name}, program={program}, set={set_name}")
    
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] or script_name == "integrated_carryover_processor":
        if program and set_name:
            raw_dir = os.path.join(BASE_DIR, program, set_name, "RAW_RESULTS")
            logger.info(f"Exam processor raw directory: {raw_dir}")
            return raw_dir
        return BASE_DIR
    
    raw_paths = {
        "utme": os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_PUTME_RESULT"),
        "caosce": os.path.join(BASE_DIR, "CAOSCE_RESULT", "RAW_CAOSCE_RESULT"),
        "clean": os.path.join(BASE_DIR, "OBJ_RESULT", "RAW_OBJ"),
        "split": os.path.join(BASE_DIR, "JAMB_DB", "RAW_JAMB_DB"),
    }
    raw_dir = raw_paths.get(script_name, BASE_DIR)
    logger.info(f"Other script raw directory: {raw_dir}")
    return raw_dir

def get_clean_directory(script_name, program=None, set_name=None):
    """Get the CLEAN_RESULTS directory for a specific script/program/set"""
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] or script_name == "integrated_carryover_processor":
        if program and set_name:
            return os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        return BASE_DIR
    
    clean_paths = {
        "utme": os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT"),
        "caosce": os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT"),
        "clean": os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ"),
        "split": os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB"),
    }
    return clean_paths.get(script_name, BASE_DIR)

# ============================================================================
# FIX 1: get_input_directory function - UPDATED VERSION
# ============================================================================
def get_input_directory(script_name, program=None, set_name=None):
    """Returns the correct input directory for raw results - FIXED VERSION"""
    logger.info(f"Getting input directory for: {script_name}, program={program}, set={set_name}")
    
    # CRITICAL FIX: For integrated_carryover_processor, determine program from set
    if script_name == "integrated_carryover_processor":
        if not program and set_name:
            program = get_program_from_set(set_name)
            logger.info(f"Determined program={program} from set={set_name}")
        
        # CRITICAL: Return the CARRYOVER subdirectory for carryover processor
        if program and set_name and set_name != "all":
            carryover_dir = os.path.join(BASE_DIR, program, set_name, "RAW_RESULTS", "CARRYOVER")
            logger.info(f"Carryover input directory: {carryover_dir}")
            return carryover_dir
        elif program:
            # Fallback to program directory if no specific set
            return os.path.join(BASE_DIR, program)
        else:
            logger.error(f"Could not determine program for set {set_name}")
            return BASE_DIR
    
    # For exam processors (non-carryover)
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
        if program and set_name and set_name != "all":
            # Regular exam processing - RAW_RESULTS directory (no CARRYOVER subdirectory)
            input_dir = os.path.join(BASE_DIR, program, set_name, "RAW_RESULTS")
            logger.info(f"Exam processor specific input directory: {input_dir}")
            return input_dir
        input_dir = os.path.join(BASE_DIR, program) if program else BASE_DIR
        logger.info(f"Exam processor general input directory: {input_dir}")
        return input_dir
    
    # For other scripts
    input_dir = get_raw_directory(script_name)
    logger.info(f"Other script input directory: {input_dir}")
    return input_dir

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "logged_in" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

def check_exam_processor_files(input_dir, program, selected_set=None):
    """Check if exam processor files exist, optionally filtering by selected set"""
    logger.info(f"Checking exam processor files in: {input_dir}")
    logger.info(f"Program: {program}, Selected Set: {selected_set}")
    
    if not os.path.isdir(input_dir):
        logger.error(f"Input directory doesn't exist: {input_dir}")
        return False

    course_dir = os.path.join(BASE_DIR, program, f"{program}-COURSES")
    course_file = None
    course_file_found = False
    
    if os.path.exists(course_dir):
        course_patterns = [
            "N-course-code-creditUnit.xlsx",
            "M-course-code-creditUnit.xlsx", 
            "course-code-creditUnit.xlsx",
            "*course*.xlsx",
            "*credit*.xlsx"
        ]
        
        for pattern in course_patterns:
            for file in os.listdir(course_dir):
                if file.lower().endswith('.xlsx') and any(keyword in file.lower() for keyword in ['course', 'credit']):
                    course_file = os.path.join(course_dir, file)
                    course_file_found = True
                    logger.info(f"Found course file: {course_file}")
                    break
            if course_file:
                break
    
    if not course_file_found:
        logger.warning(f"Course file not found in: {course_dir}")

    program_dir = os.path.join(BASE_DIR, program)
    if not os.path.exists(program_dir):
        logger.error(f"Program directory not found: {program_dir}")
        return False

    valid_sets = BN_SETS if program == "BN" else (BM_SETS if program == "BM" else ND_SETS)
    
    if selected_set and selected_set != "all":
        if selected_set in valid_sets:
            program_sets = [selected_set]
            logger.info(f"Processing specific set: {selected_set}")
        else:
            logger.error(f"Invalid set selected: {selected_set}")
            return False
    else:
        program_sets = []
        for item in os.listdir(program_dir):
            item_path = os.path.join(program_dir, item)
            if os.path.isdir(item_path) and item in valid_sets:
                program_sets.append(item)
        logger.info(f"Processing all sets: {program_sets}")

    if not program_sets:
        logger.error(f"No {program} sets found in {program_dir} (valid sets: {valid_sets})")
        return False

    total_files_found = 0
    files_found = []
    
    for program_set in program_sets:
        raw_results_path = os.path.join(program_dir, program_set, "RAW_RESULTS")
        logger.info(f"Checking raw results path: {raw_results_path}")
        
        if not os.path.exists(raw_results_path):
            logger.warning(f"RAW_RESULTS not found in {raw_results_path}")
            continue
            
        files = []
        for f in os.listdir(raw_results_path):
            file_path = os.path.join(raw_results_path, f)
            if (os.path.isfile(file_path) and 
                f.lower().endswith((".xlsx", ".xls")) and
                not f.startswith("~$") and
                not f.startswith(".")):
                files.append(f)
                files_found.append(f"{program_set}/{f}")
        
        total_files_found += len(files)
        
        if files:
            logger.info(f"Found {len(files)} files in {raw_results_path}: {files}")
        else:
            logger.warning(f"No Excel files found in {raw_results_path}")

    logger.info(f"Total Excel files found for {program}: {total_files_found}")
    logger.info(f"Files found: {files_found}")
    
    return total_files_found > 0

def check_putme_files(input_dir):
    if not os.path.isdir(input_dir):
        return False
    excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".xlsx", ".xls")) and "PUTME" in f.upper()]
    candidate_batches_dir = os.path.join(os.path.dirname(input_dir), "RAW_CANDIDATE_BATCHES")
    batch_files = [f for f in os.listdir(candidate_batches_dir) if f.lower().endswith(".csv") and "BATCH" in f.upper()] if os.path.isdir(candidate_batches_dir) else []
    return len(excel_files) > 0 and len(batch_files) > 0

def check_internal_exam_files(input_dir):
    """Check for internal exam files - UPDATED to recognize both Set pattern and new ND-SET pattern"""
    if not os.path.isdir(input_dir):
        logger.error(f"Directory doesn't exist: {input_dir}")
        return False
    
    # Get all valid files (same as processing script)
    csv_files = glob.glob(os.path.join(input_dir, "*.csv"))
    xls_files = glob.glob(os.path.join(input_dir, "*.xls")) + glob.glob(os.path.join(input_dir, "*.xlsx"))
    all_files = [f for f in (csv_files + xls_files) if not os.path.basename(f).startswith("~$")]
    
    # Also check for files with patterns that indicate internal exam results
    pattern_files = [
        f for f in os.listdir(input_dir) 
        if f.lower().endswith((".xlsx", ".xls", ".csv")) 
        and not f.startswith("~")
        and (
            f.startswith("Set") or  # Original pattern
            "ND" in f.upper() and "SET" in f.upper() or  # New pattern like ND2024-SET1
            "OBJ" in f.upper() or  # Objective results
            "RESULT" in f.upper()  # General result files
        )
    ]
    
    # Use whichever method gives us more files
    final_files = list(set(all_files + pattern_files))
    
    logger.info(f"Internal exam files found in {input_dir}: {len(final_files)} files")
    
    return len(final_files) > 0

def check_caosce_files(input_dir):
    if not os.path.isdir(input_dir):
        return False
    excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".xlsx", ".xls")) and "CAOSCE" in f.upper()]
    return len(excel_files) > 0

def check_split_files(input_dir):
    if not os.path.isdir(input_dir):
        return False
    valid_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".csv", ".xlsx", ".xls")) and not f.startswith("~")]
    return len(valid_files) > 0

def check_input_files(input_dir, script_name, selected_set=None):
    """Check input files with optional set filtering"""
    logger.info(f"Checking input files for {script_name} in {input_dir} (selected_set: {selected_set})")
    
    if not os.path.isdir(input_dir):
        logger.error(f"Input directory doesn't exist: {input_dir}")
        return False
        
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm", "integrated_carryover_processor"]:
        program = script_name.split("_")[-1].upper()
        if script_name == "integrated_carryover_processor":
            program = get_program_from_set(selected_set)
            if program:
                logger.info(f"Set program to {program} for carryover based on set {selected_set}")
            else:
                logger.error(f"Could not determine program for set {selected_set}")
                return False
        return check_exam_processor_files(input_dir, program, selected_set)
    elif script_name == "utme":
        return check_putme_files(input_dir)
    elif script_name == "clean":
        return check_internal_exam_files(input_dir)
    elif script_name == "caosce":
        return check_caosce_files(input_dir)
    elif script_name == "split":
        return check_split_files(input_dir)
    try:
        dir_contents = os.listdir(input_dir)
        valid_extensions = (".csv", ".xlsx", ".xls", ".pdf")
        input_files = [f for f in dir_contents if f.lower().endswith(valid_extensions) and not f.startswith("~")]
        return len(input_files) > 0
    except Exception:
        return False

def get_exam_processor_status(program, selected_set=None):
    """Get detailed status for exam processor including file counts"""
    logger.info(f"Getting status for {program}, set: {selected_set}")
    
    if program == "PROCESSOR" and selected_set:
        program = get_program_from_set(selected_set)
        if program:
            logger.info(f"Overrode program to {program} based on set {selected_set}")
        else:
            logger.warning(f"Could not determine program from set {selected_set}")
    
    status_info = {
        'ready': False,
        'course_file': False,
        'raw_files_count': 0,
        'raw_files_list': [],
        'sets_ready': {}
    }
    
    course_dir = os.path.join(BASE_DIR, program, f"{program}-COURSES")
    if os.path.exists(course_dir):
        for file in os.listdir(course_dir):
            if file.lower().endswith('.xlsx') and any(keyword in file.lower() for keyword in ['course', 'credit']):
                status_info['course_file'] = True
                break
    
    program_dir = os.path.join(BASE_DIR, program)
    if not os.path.exists(program_dir):
        logger.error(f"Program directory not found: {program_dir}")
        return status_info
    
    valid_sets = BN_SETS if program == "BN" else (BM_SETS if program == "BM" else ND_SETS)
    
    sets_to_check = []
    if selected_set and selected_set != "all":
        if selected_set in valid_sets:
            sets_to_check = [selected_set]
    else:
        sets_to_check = valid_sets
    
    total_files = 0
    for set_name in sets_to_check:
        raw_results_path = os.path.join(program_dir, set_name, "RAW_RESULTS")
        set_files = []
        
        if os.path.exists(raw_results_path):
            for f in os.listdir(raw_results_path):
                file_path = os.path.join(raw_results_path, f)
                if (os.path.isfile(file_path) and 
                    f.lower().endswith((".xlsx", ".xls")) and
                    not f.startswith("~$") and
                    not f.startswith(".")):
                    set_files.append(f)
        
        status_info['sets_ready'][set_name] = len(set_files) > 0
        status_info['raw_files_count'] += len(set_files)
        status_info['raw_files_list'].extend([f"{set_name}/{f}" for f in set_files])
        total_files += len(set_files)
    
    status_info['ready'] = status_info['course_file'] and total_files > 0
    
    logger.info(f"Status for {program}: course_file={status_info['course_file']}, raw_files={total_files}, ready={status_info['ready']}")
    return status_info
    
def count_processed_files(output_lines, script_name, selected_set=None):
    success_indicators = SUCCESS_INDICATORS.get(script_name, [])
    processed_files_set = set()
    logger.info(f"Raw output lines for {script_name}:")
    for line in output_lines:
        if line.strip():
            logger.info(f"  OUTPUT: {line}")
    for line in output_lines:
        for indicator in success_indicators:
            match = re.search(indicator, line, re.IGNORECASE)
            if match:
                if script_name in ["exam_processor_bn", "exam_processor_bm", "exam_processor_nd", "integrated_carryover_processor"]:
                    if "PROCESSING SEMESTER:" in line.upper() or "PROCESSING SET:" in line.upper():
                        set_name = match.group(1)
                        if set_name:
                            processed_files_set.add(f"Set: {set_name}")
                            logger.info(f"DETECTED SET: {set_name}")
                    elif "Mastersheet saved:" in line:
                        file_name = match.group(0).split("Mastersheet saved:")[-1].strip()
                        processed_files_set.add(f"Saved: {file_name}")
                    elif "Identified" in line and "carryover students" in line:
                        processed_files_set.add("Carryover students identified")
                    elif "Carryover records saved:" in line:
                        processed_files_set.add("Carryover records saved")
                    elif "Processing resit results for" in line:
                        processed_files_set.add("Resit processing started")
                    elif "Updated" in line and "scores for" in line and "students" in line:
                        processed_files_set.add("Resit scores updated")
                elif script_name == "utme":
                    if "Processing:" in line:
                        file_name = match.group(1)
                        processed_files_set.add(f"Processed: {file_name}")
                    elif "Saved processed file:" in line:
                        file_name = match.group(1)
                        processed_files_set.add(f"Saved: {file_name}")
                elif script_name == "clean":
                    if "Processing:" in line:
                        file_name = match.group(1)
                        processed_files_set.add(f"Processed: {file_name}")
                    elif "Cleaned CSV saved" in line:
                        file_name = match.group(1) if match.groups() else "cleaned_file"
                        processed_files_set.add(f"Cleaned: {file_name}")
                    elif "Master CSV saved" in line:
                        processed_files_set.add("Master file created")
                    elif "All processing completed successfully!" in line:
                        processed_files_set.add("Processing completed")
                else:
                    file_name = match.group(1) if match.groups() else line
                    processed_files_set.add(file_name)
    logger.info(f"Processed items for {script_name}: {processed_files_set}")
    return len(processed_files_set)

def get_success_message(script_name, processed_files, output_lines, selected_set=None):
    if processed_files == 0:
        return None
    if script_name == "clean":
        if any("All processing completed successfully!" in line for line in output_lines):
            return f"Successfully processed internal examination results! Generated master file and individual cleaned files."
        return f"Processed {processed_files} internal examination file(s)."
    elif script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
        program = script_name.split("_")[-1].upper()
        program_name = {"ND": "ND", "BN": "Basic Nursing", "BM": "Basic Midwifery"}.get(program, program)
        upgrade_info = ""
        upgrade_count = ""
        carryover_info = ""
        resit_info = ""
        resit_updates = ""
        
        for line in output_lines:
            if "Applying upgrade rule:" in line:
                upgrade_match = re.search(r"Applying upgrade rule: (\d+)–49 → 50", line)
                if upgrade_match:
                    upgrade_info = f" Upgrade rule applied: {upgrade_match.group(1)}-49 → 50"
                    break
            elif "Upgraded" in line:
                upgrade_count_match = re.search(r"Upgraded (\d+) scores", line)
                if upgrade_count_match:
                    upgrade_count = f" Upgraded {upgrade_count_match.group(1)} scores"
                    break
            elif "Identified" in line and "carryover students" in line:
                carryover_match = re.search(r"Identified (\d+) carryover students", line)
                if carryover_match:
                    carryover_info = f" Identified {carryover_match.group(1)} carryover students"
                    break
            elif "Updated" in line and "scores for" in line and "students" in line:
                resit_match = re.search(r"Updated (\d+) scores for (\d+) students", line)
                if resit_match:
                    resit_updates = f" Resit: updated {resit_match.group(1)} scores for {resit_match.group(2)} students"
                    break
            elif "Processing resit results for" in line:
                resit_info = " Resit processing completed"
                break
        
        if resit_updates:
            return f"{program_name} Examination processing completed!{resit_updates}{upgrade_info}{upgrade_count}{carryover_info}"
        elif any(f"{program_name} Examination Results Processing completed successfully" in line for line in output_lines):
            return f"{program_name} Examination processing completed successfully! Processed {processed_files} set(s).{upgrade_info}{upgrade_count}{carryover_info}{resit_info}"
        elif any("Processing complete" in line for line in output_lines):
            return f"{program_name} Examination processing completed! Processed {processed_files} set(s).{upgrade_info}{upgrade_count}{carryover_info}{resit_info}"
        return f"Processed {processed_files} {program_name} examination set(s).{upgrade_info}{upgrade_count}{carryover_info}{resit_info}"
    elif script_name == "integrated_carryover_processor":
        resit_updates = ""
        for line in output_lines:
            if "Updated" in line and "scores for" in line and "students" in line:
                resit_match = re.search(r"Updated (\d+) scores for (\d+) students", line)
                if resit_match:
                    resit_updates = f"Updated {resit_match.group(1)} scores for {resit_match.group(2)} students"
                    break
        if resit_updates:
            return f"Carryover processing completed! {resit_updates}"
        elif any("Carryover processing completed successfully" in line for line in output_lines):
            return "Carryover processing completed successfully"
        elif any("Processing resit results for" in line for line in output_lines):
            return "Carryover processing completed but no scores updated"
        return "Carryover processing completed"
    elif script_name == "utme":
        if any("Processing completed successfully" in line for line in output_lines):
            return f"PUTME processing completed successfully! Processed {processed_files} batch file(s)."
        return f"Processed {processed_files} PUTME batch file(s)."
    elif script_name == "caosce":
        if any("Processed" in line for line in output_lines):
            return f"CAOSCE processing completed! Processed {processed_files} file(s)."
        return f"Processed {processed_files} CAOSCE file(s)."
    elif script_name == "split":
        if any("Saved processed file:" in line for line in output_lines):
            return f"JAMB name splitting completed! Processed {processed_files} file(s)."
        return f"Processed {processed_files} JAMB file(s)."
    return f"Successfully processed {processed_files} file(s)."

def _get_script_path(script_name):
    script_path = os.path.join(SCRIPT_DIR, SCRIPT_MAP.get(script_name, ""))
    logger.info(f"Script path for {script_name}: {script_path}")
    if not os.path.exists(script_path):
        logger.error(f"Script not found: {script_path}")
        raise FileNotFoundError(f"Script {script_name} not found at {script_path}")
    return script_path

def get_files_by_category():
    """Get files organized by category and semester/set from existing ZIP files only"""
    from dataclasses import dataclass

    @dataclass
    class FileInfo:
        name: str
        relative_path: str
        folder: str
        size: int
        modified: int
        semester: str = ""
        set_name: str = ""

    files_by_category = {
        "nd_results": {},
        "bn_results": {},
        "bm_results": {},
        "putme_results": [],
        "caosce_results": [],
        "internal_results": [],
        "jamb_results": [],
    }

    # Only show ZIP files in download center
    for program in ["ND", "BN", "BM"]:
        program_dir = os.path.join(BASE_DIR, program)
        if not os.path.exists(program_dir):
            continue
        
        sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
        for set_name in sets:
            clean_dir = os.path.join(program_dir, set_name, "CLEAN_RESULTS")
            if not os.path.exists(clean_dir):
                continue
            
            category = f"{program.lower()}_results"
            if set_name not in files_by_category[category]:
                files_by_category[category][set_name] = []
            
            # Only add ZIP files
            for file in os.listdir(clean_dir):
                if file.lower().endswith(".zip"):
                    file_path = os.path.join(clean_dir, file)
                    try:
                        relative_path = os.path.relpath(file_path, BASE_DIR)
                        folder = os.path.basename(os.path.dirname(file_path))
                        
                        file_info = FileInfo(
                            name=file,
                            relative_path=relative_path,
                            folder=folder,
                            size=os.path.getsize(file_path),
                            modified=os.path.getmtime(file_path),
                            set_name=set_name
                        )
                        files_by_category[category][set_name].append(file_info)
                        logger.info(f"Categorized as {category}/{set_name}: {file}")
                    except Exception as e:
                        logger.error(f"Error processing file {file}: {e}")

    # Other result types - only show ZIP files
    result_mappings = {
        "PUTME_RESULT/CLEAN_PUTME_RESULT": "putme_results",
        "CAOSCE_RESULT/CLEAN_CAOSCE_RESULT": "caosce_results",
        "OBJ_RESULT/CLEAN_OBJ": "internal_results",
        "JAMB_DB/CLEAN_JAMB_DB": "jamb_results",
    }
    
    for path, category in result_mappings.items():
        clean_dir = os.path.join(BASE_DIR, path)
        if not os.path.exists(clean_dir):
            continue
        
        # Only add ZIP files
        for file in os.listdir(clean_dir):
            if file.lower().endswith(".zip"):
                file_path = os.path.join(clean_dir, file)
                try:
                    relative_path = os.path.relpath(file_path, BASE_DIR)
                    folder = os.path.basename(os.path.dirname(file_path))
                    file_info = FileInfo(
                        name=file,
                        relative_path=relative_path,
                        folder=folder,
                        size=os.path.getsize(file_path),
                        modified=os.path.getmtime(file_path)
                    )
                    files_by_category[category].append(file_info)
                    logger.info(f"Categorized as {category}: {file}")
                except Exception as e:
                    logger.error(f"Error processing file {file}: {e}")

    for category, files in files_by_category.items():
        if isinstance(files, dict):
            total_files = sum(len(files_in_set) for files_in_set in files.values())
            logger.info(f"{category}: {total_files} ZIP files across {len(files)} sets")
        else:
            logger.info(f"{category}: {len(files)} ZIP files")
    
    return files_by_category

def get_sets_and_folders():
    """Get sets and their result folders from CLEAN_RESULTS directories - ONLY ZIP FILES"""
    from dataclasses import dataclass

    @dataclass
    class FileInfo:
        name: str
        relative_path: str
        size: int
        modified: int

    @dataclass
    class FolderInfo:
        name: str
        files: list

    sets = {}
    logger.info(f"Scanning BASE_DIR for ZIP files only: {BASE_DIR}")

    for program in ["ND", "BN", "BM"]:
        program_dir = os.path.join(BASE_DIR, program)
        if not os.path.exists(program_dir):
            logger.warning(f"Program directory not found: {program_dir}")
            continue
        logger.info(f"Scanning program: {program_dir}")
        valid_sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
        for set_name in os.listdir(program_dir):
            set_path = os.path.join(program_dir, set_name)
            if not os.path.isdir(set_path) or set_name not in valid_sets:
                logger.warning(f"Skipping invalid set {set_name} for program {program}")
                continue
            logger.info(f"Scanning set: {set_path}")
            folders = []
            clean_results_path = os.path.join(set_path, "CLEAN_RESULTS")
            if os.path.exists(clean_results_path):
                logger.info(f"Found CLEAN_RESULTS: {clean_results_path}")
                try:
                    # Only show ZIP files in file browser
                    zip_files = []
                    for file in os.listdir(clean_results_path):
                        if file.lower().endswith(".zip"):
                            file_path = os.path.join(clean_results_path, file)
                            try:
                                relative_path = os.path.relpath(file_path, BASE_DIR)
                                zip_files.append(FileInfo(
                                    name=file,
                                    relative_path=relative_path,
                                    size=os.path.getsize(file_path),
                                    modified=os.path.getmtime(file_path)
                                ))
                                logger.info(f"Found ZIP file: {file} in {clean_results_path}")
                            except Exception as e:
                                logger.error(f"Error processing file {file}: {e}")
                    
                    if zip_files:
                        folders.append(FolderInfo(name="ZIP Results", files=zip_files))
                        logger.info(f"Added ZIP files folder with {len(zip_files)} files")
                except Exception as e:
                    logger.error(f"Error scanning {clean_results_path}: {e}")
            if folders:
                set_key = f"{program}_{set_name}"
                sets[set_key] = folders
                logger.info(f"Added set {set_key} with {len(folders)} folders")

    # Other result types - only show ZIP files
    result_mappings = {
        "PUTME_RESULT": {
            "clean_dir": "CLEAN_PUTME_RESULT",
            "base_dir": "PUTME_RESULT"
        },
        "CAOSCE_RESULT": {
            "clean_dir": "CLEAN_CAOSCE_RESULT", 
            "base_dir": "CAOSCE_RESULT"
        },
        "INTERNAL_RESULT": {
            "clean_dir": "CLEAN_OBJ",
            "base_dir": "OBJ_RESULT"
        },
        "JAMB_DB": {
            "clean_dir": "CLEAN_JAMB_DB",
            "base_dir": "JAMB_DB"
        }
    }
    
    for result_type, mapping in result_mappings.items():
        clean_dir_name = mapping["clean_dir"]
        base_dir_name = mapping["base_dir"]
        
        result_path = os.path.join(BASE_DIR, base_dir_name, clean_dir_name)
        logger.info(f"Scanning {result_type} at: {result_path}")
        
        if not os.path.exists(result_path):
            logger.warning(f"Directory not found: {result_path}")
            continue
            
        folders = []
        try:
            # Only show ZIP files
            zip_files = []
            for file in os.listdir(result_path):
                if file.lower().endswith(".zip"):
                    file_path = os.path.join(result_path, file)
                    try:
                        relative_path = os.path.relpath(file_path, BASE_DIR)
                        zip_files.append(FileInfo(
                            name=file,
                            relative_path=relative_path,
                            size=os.path.getsize(file_path),
                            modified=os.path.getmtime(file_path)
                        ))
                        logger.info(f"Found ZIP file: {file} in {result_path}")
                    except Exception as e:
                        logger.error(f"Error processing file {file}: {e}")
            
            if zip_files:
                folders.append(FolderInfo(name="ZIP Results", files=zip_files))
                logger.info(f"Added ZIP files folder with {len(zip_files)} files")
                    
        except Exception as e:
            logger.error(f"Error scanning {result_path}: {e}")
            
        if folders:
            sets[result_type] = folders
            logger.info(f"Added result type {result_type} with {len(folders)} folders")
        else:
            logger.info(f"No ZIP folders found for {result_type}")

    logger.info(f"Total sets found with ZIP files: {len(sets)}")
    for set_name, folders in sets.items():
        total_files = sum(len(folder.files) for folder in folders)
        logger.info(f"{set_name}: {total_files} ZIP files in {len(folders)} folders")
        
    return sets

# ============================================================================
# FIXED: get_carryover_records function with ZIP file support
# ============================================================================
def get_carryover_records(program, set_name, semester_key=None):
    """Get carryover records for a specific program, set, and semester - FIXED VERSION."""
    try:
        # Standardize semester key first
        if semester_key:
            semester_key = standardize_semester_key(semester_key)
            logger.info(f"🔑 Using standardized semester key: {semester_key}")
        
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            logger.info(f"❌ Clean directory not found: {clean_dir}")
            return []
        
        # Look for both folders and ZIP files (REGULAR results, not carryover)
        timestamp_items = []
        
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
            
            # CRITICAL FIX: Only include regular result files (exclude CARRYOVER files)
            # Check if it starts with set name and contains "RESULT" but NOT "CARRYOVER"
            if (item.startswith(f"{set_name}_RESULT-") and 
                "RESULT" in item.upper() and 
                not "CARRYOVER" in item.upper()):
                
                if os.path.isdir(item_path) or item.endswith('.zip'):
                    timestamp_items.append(item)
                    logger.info(f"✅ Found regular result: {item}")
        
        if not timestamp_items:
            logger.info(f"❌ No regular result files found in: {clean_dir}")
            logger.info(f"📁 Available files: {[f for f in os.listdir(clean_dir) if not f.startswith('.')]}")
            return []
        
        # Sort to get the latest result
        latest_item = sorted(timestamp_items)[-1]
        latest_path = os.path.join(clean_dir, latest_item)
        logger.info(f"✅ Using latest result: {latest_item}")
        
        # Extract from ZIP or use folder
        if latest_item.endswith('.zip'):
            return get_carryover_records_from_zip(latest_path, set_name, semester_key)
        else:
            carryover_dir = os.path.join(latest_path, "CARRYOVER_RECORDS")
            if not os.path.exists(carryover_dir):
                logger.info(f"❌ No CARRYOVER_RECORDS folder in: {latest_path}")
                return []
            return load_carryover_json_files(carryover_dir, semester_key)
            
    except Exception as e:
        logger.error(f"Error getting carryover records: {e}")
        return []

def get_carryover_records_from_zip(zip_path, set_name, semester_key=None):
    """Extract carryover records from ZIP file"""
    try:
        logger.info(f"📦 Extracting carryover records from ZIP: {zip_path}")
        carryover_files = []
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # List all files in ZIP for debugging
            all_files = zip_ref.namelist()
            logger.info(f"📁 Files in ZIP: {all_files}")
            
            # Look for carryover JSON files
            json_files = [f for f in all_files if f.startswith("CARRYOVER_RECORDS/") and f.endswith('.json')]
            
            if not json_files:
                logger.info(f"❌ No carryover JSON files found in ZIP")
                return []
                
            for json_file in json_files:
                file_semester = extract_semester_from_filename(json_file)
                file_semester_standardized = standardize_semester_key(file_semester)
                
                if semester_key and file_semester_standardized != semester_key:
                    logger.info(f"   ⏭️ Skipping (doesn't match target)")
                    continue
                    
                try:
                    with zip_ref.open(json_file) as f:
                        data = json.load(f)
                        carryover_files.append({
                            'filename': os.path.basename(json_file),
                            'semester': file_semester_standardized,
                            'data': data,
                            'count': len(data),
                            'file_path': f"{zip_path}/{json_file}"
                        })
                        logger.info(f"✅ Loaded carryover record: {json_file}")
                except Exception as e:
                    logger.error(f"Error loading carryover file {json_file}: {e}")
        
        logger.info(f"✅ Loaded {len(carryover_files)} carryover records from ZIP")
        return carryover_files
        
    except Exception as e:
        logger.error(f"Error extracting carryover records from ZIP: {e}")
        return []

# ============================================================================
# FIXED: load_carryover_json_files function - UPDATED VERSION
# ============================================================================
def load_carryover_json_files(carryover_dir, semester_key=None):
    """Load carryover JSON files from directory - FIXED."""
    carryover_files = []
    
    # Standardize the target semester key
    if semester_key:
        semester_key = standardize_semester_key(semester_key)
    
    for file in os.listdir(carryover_dir):
        if file.startswith("co_student_") and file.endswith(".json"):
            # Extract semester from filename and standardize it
            file_semester = extract_semester_from_filename(file)
            file_semester_standardized = standardize_semester_key(file_semester)
            
            logger.info(f"📄 Found carryover file: {file}")
            logger.info(f"   Original semester: {file_semester}")
            logger.info(f"   Standardized: {file_semester_standardized}")
            logger.info(f"   Target semester: {semester_key}")
            
            # If semester_key is specified, only load matching files
            if semester_key and file_semester_standardized != semester_key:
                logger.info(f"   ⏭️ Skipping (doesn't match target)")
                continue
            
            file_path = os.path.join(carryover_dir, file)
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f)
                    carryover_files.append({
                        'filename': file,
                        'semester': file_semester_standardized,  # Use standardized key
                        'data': data,
                        'count': len(data),
                        'file_path': file_path
                    })
                    logger.info(f"   ✅ Loaded: {len(data)} records")
            except Exception as e:
                logger.error(f"Error loading {file}: {e}")
    
    logger.info(f"📊 Total carryover files loaded: {len(carryover_files)}")
    return carryover_files

# ============================================================================
# FIXED: extract_semester_from_filename function - ENHANCED VERSION
# ============================================================================
def extract_semester_from_filename(filename):
    """Extract semester from filename using comprehensive pattern matching - FIXED."""
    filename_upper = filename.upper()
    
    # First, try to extract any semester-like pattern from filename
    semester_pattern = r'(ND[-_]?(?:FIRST|SECOND|1ST|2ND)[-_]?YEAR[-_]?(?:FIRST|SECOND|1ST|2ND)[-_]?SEMESTER)'
    match = re.search(semester_pattern, filename_upper)
    
    if match:
        extracted = match.group(1)
        # Standardize the extracted pattern
        standardized = standardize_semester_key(extracted)
        logger.info(f"✅ Extracted and standardized: '{filename}' → '{standardized}'")
        return standardized
    
    # Fallback to comprehensive pattern mapping
    semester_patterns = {
        "ND-First-YEAR-First-SEMESTER": [
            "FIRST.YEAR.FIRST.SEMESTER", "FIRST-YEAR-FIRST-SEMESTER", "FIRST_YEAR_FIRST_SEMESTER",
            "1ST.YEAR.1ST.SEMESTER", "1ST-YEAR-1ST-SEMESTER", "1ST_YEAR_1ST_SEMESTER",
            "YEAR1.SEMESTER1", "YEAR-1-SEMESTER-1", "YEAR_1_SEMESTER_1",
            "FIRST.SEMESTER.FIRST.YEAR", "1ST.SEMESTER.1ST.YEAR",
            "ND-FIRST-YEAR-FIRST-SEMESTER", "ND-1ST-YEAR-1ST-SEMESTER"
        ],
        "ND-First-YEAR-SECOND-SEMESTER": [
            "FIRST.YEAR.SECOND.SEMESTER", "FIRST-YEAR-SECOND-SEMESTER", "FIRST_YEAR_SECOND_SEMESTER",
            "1ST.YEAR.2ND.SEMESTER", "1ST-YEAR-2ND-SEMESTER", "1ST_YEAR_2ND_SEMESTER",
            "YEAR1.SEMESTER2", "YEAR-1-SEMESTER-2", "YEAR_1_SEMESTER_2",
            "SECOND.SEMESTER.FIRST.YEAR", "2ND.SEMESTER.1ST.YEAR",
            "ND-FIRST-YEAR-SECOND-SEMESTER", "ND-1ST-YEAR-2ND-SEMESTER"
        ],
        "ND-SECOND-YEAR-First-SEMESTER": [
            "SECOND.YEAR.FIRST.SEMESTER", "SECOND-YEAR-FIRST-SEMESTER", "SECOND_YEAR_FIRST_SEMESTER",
            "2ND.YEAR.1ST.SEMESTER", "2ND-YEAR-1ST-SEMESTER", "2ND_YEAR_1ST_SEMESTER",
            "YEAR2.SEMESTER1", "YEAR-2-SEMESTER-1", "YEAR_2_SEMESTER_1",
            "FIRST.SEMESTER.SECOND.YEAR", "1ST.SEMESTER.2ND.YEAR",
            "ND-SECOND-YEAR-FIRST-SEMESTER", "ND-2ND-YEAR-1ST-SEMESTER"
        ],
        "ND-SECOND-YEAR-SECOND-SEMESTER": [
            "SECOND.YEAR.SECOND.SEMESTER", "SECOND-YEAR-SECOND-SEMESTER", "SECOND_YEAR_SECOND_SEMESTER",
            "2ND.YEAR.2ND.SEMESTER", "2ND-YEAR-2ND-SEMESTER", "2ND_YEAR_2ND_SEMESTER",
            "YEAR2.SEMESTER2", "YEAR-2-SEMESTER-2", "YEAR_2_SEMESTER_2",
            "SECOND.SEMESTER.SECOND.YEAR", "2ND.SEMESTER.2ND.YEAR",
            "ND-SECOND-YEAR-SECOND-SEMESTER", "ND-2ND-YEAR-2ND-SEMESTER"
        ]
    }
    
    for semester_key, patterns in semester_patterns.items():
        for pattern in patterns:
            flexible_pattern = pattern.replace('.', '[._\\- ]?')
            if re.search(flexible_pattern, filename_upper, re.IGNORECASE):
                logger.info(f"✅ Matched semester '{semester_key}' for filename: {filename}")
                return semester_key
    
    logger.warning(f"❌ Could not determine semester for filename: {filename}")
    return "UNKNOWN_SEMESTER"

# ============================================================================
# FIXED: get_carryover_summary function - UPDATED VERSION
# ============================================================================
def get_carryover_summary(program, set_name):
    """Get summary of carryover students - FIXED VERSION with ZIP support"""
    # CRITICAL FIX: Check if CLEAN_RESULTS exists before trying to scan
    clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
    if not os.path.exists(clean_dir):
        logger.info(f"⚠️ CLEAN_RESULTS not found for {program}/{set_name}, skipping")
        return {
            'total_students': 0,
            'total_courses': 0,
            'by_semester': {},
            'recent_semester': None,
            'recent_count': 0
        }
    
    carryover_records = get_carryover_records(program, set_name)
    summary = {
        'total_students': 0,
        'total_courses': 0,
        'by_semester': {},
        'recent_semester': None,
        'recent_count': 0
    }
    
    logger.info(f"📊 Carryover summary for {program}/{set_name}: {len(carryover_records)} records found")
    
    for record in carryover_records:
        semester = record['semester']
        if semester not in summary['by_semester']:
            summary['by_semester'][semester] = 0
        
        summary['by_semester'][semester] += record['count']
        summary['total_students'] += record['count']
        summary['total_courses'] += sum(len(student['failed_courses']) for student in record['data'])
    
    if summary['by_semester']:
        summary['recent_semester'] = max(summary['by_semester'].keys(), 
                                       key=lambda x: summary['by_semester'][x])
        summary['recent_count'] = summary['by_semester'][summary['recent_semester']]
    
    logger.info(f"📊 Final summary: {summary}")
    return summary

def rename_carryover_files(carryover_output_dir, semester_key, resit_timestamp):
    """Rename files in carryover output directory to include CARRYOVER prefix"""
    renamed_files = []
    
    for root, dirs, files in os.walk(carryover_output_dir):
        for file in files:
            if file.lower().endswith(('.xlsx', '.csv', '.pdf')):
                old_path = os.path.join(root, file)
                
                if file.startswith('CARRYOVER_'):
                    renamed_files.append(old_path)
                    continue
                
                file_extension = os.path.splitext(file)[1]
                file_base = os.path.splitext(file)[0]
                
                new_filename = f"CARRYOVER_{semester_key}_{resit_timestamp}_{file_base}{file_extension}"
                new_path = os.path.join(root, new_filename)
                
                try:
                    os.rename(old_path, new_path)
                    renamed_files.append(new_path)
                    logger.info(f"Renamed carryover file: {file} → {new_filename}")
                except Exception as e:
                    logger.error(f"Failed to rename {file}: {e}")
                    renamed_files.append(old_path)
    
    return renamed_files

def debug_resit_processing_details(program_code, set_name, semester_key, resit_file_path, clean_dir):
    """Debug function to provide detailed information about resit processing"""
    debug_info = {
        'program': program_code,
        'set': set_name,
        'semester': semester_key,
        'resit_file': resit_file_path,
        'resit_file_exists': os.path.exists(resit_file_path),
        'clean_dir': clean_dir,
        'clean_dir_exists': os.path.exists(clean_dir),
        'timestamp_folders': [],
        'carryover_records': []
    }
    
    if os.path.exists(clean_dir):
        debug_info['timestamp_folders'] = [f for f in os.listdir(clean_dir) 
                                         if f.startswith(f"{set_name}_RESULT-") and os.path.isdir(os.path.join(clean_dir, f))]
    
    latest_folder = sorted(debug_info['timestamp_folders'])[-1] if debug_info['timestamp_folders'] else None
    if latest_folder:
        output_dir = os.path.join(clean_dir, latest_folder)
        carryover_dir = os.path.join(output_dir, "CARRYOVER_RECORDS")
        if os.path.exists(carryover_dir):
            debug_info['carryover_records'] = [f for f in os.listdir(carryover_dir) 
                                             if f.startswith("co_student_") and semester_key in f and f.endswith('.json')]
    
    logger.info("RESIT PROCESSING DEBUG INFO:")
    for key, value in debug_info.items():
        logger.info(f"  {key}: {value}")
    
    return debug_info

def cleanup_scattered_files(clean_dir, zip_filename):
    """Remove all scattered files and folders after successful zipping"""
    try:
        # Remove any result directories
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
            if os.path.isdir(item_path) and (item.startswith(f"{item.split('_')[0]}_RESULT-") or "RESULT" in item):
                shutil.rmtree(item_path)
                logger.info(f"Removed scattered folder: {item_path}")
            
            # Remove individual files (keep only ZIP files)
            elif os.path.isfile(item_path) and not item.lower().endswith('.zip'):
                os.remove(item_path)
                logger.info(f"Removed scattered file: {item_path}")
                
        logger.info(f"Cleanup completed for {clean_dir}. Only ZIP files remain.")
        return True
    except Exception as e:
        logger.error(f"Error during cleanup: {e}")
        return False

# ============================================================================
# ENHANCED: verify_semester_in_results function with comprehensive matching
# ============================================================================
def verify_semester_in_results(result_path, semester_key):
    """Enhanced function to verify that the semester exists in the result files with comprehensive matching"""
    try:
        logger.info(f"🔍 Verifying semester '{semester_key}' in: {result_path}")
        
        # Normalize semester key for comparison
        normalized_key = semester_key.upper().replace("-", " ").replace("_", " ").replace(".", " ").strip()
        normalized_key = ' '.join(normalized_key.split())  # Remove extra spaces
        
        # Remove program prefix if present for more flexible matching
        search_key = normalized_key
        if normalized_key.startswith(('ND ', 'BN ', 'BM ')):
            search_key = ' '.join(normalized_key.split()[1:])
        
        logger.info(f"📝 Searching for semester pattern: '{search_key}'")
        
        def check_item_for_semester(item_name):
            """Check if a single item matches the semester pattern"""
            item_normalized = item_name.upper().replace("-", " ").replace("_", " ").replace(".", " ").strip()
            item_normalized = ' '.join(item_normalized.split())
            
            # Comprehensive matching strategies
            matching_strategies = [
                # Exact match (after normalization)
                search_key in item_normalized,
                # Match with common variations
                any(term in item_normalized for term in search_key.split()),
                # Match semester numbers (1st, 2nd, etc.)
                any(f"{num}{suffix}" in item_normalized 
                   for num, suffix in [('1', 'ST'), ('2', 'ND'), ('3', 'RD')] 
                   for term in search_key.split() if term in ['FIRST', 'SECOND', 'THIRD']),
                # Match abbreviated forms
                any(abbr in item_normalized 
                   for abbr in ['SEM1', 'SEM2', 'SEM3', 'YR1', 'YR2', 'YR3'] 
                   if any(term in search_key for term in ['FIRST', 'SECOND', 'THIRD', '1', '2', '3']))
            ]
            
            if any(matching_strategies):
                logger.info(f"✅ Semester match found: '{item_name}' → '{semester_key}'")
                return True
            return False
        
        # Check directory contents
        if os.path.isdir(result_path):
            for item in os.listdir(result_path):
                if check_item_for_semester(item):
                    return True
        
        # Check ZIP contents
        elif result_path.endswith('.zip'):
            try:
                with zipfile.ZipFile(result_path, 'r') as zip_ref:
                    for item in zip_ref.namelist():
                        if check_item_for_semester(item):
                            return True
            except Exception as e:
                logger.error(f"Error reading ZIP file: {e}")
        
        logger.warning(f"❌ Semester '{semester_key}' not found in results")
        logger.info(f"💡 Search pattern used: '{search_key}'")
        return False
        
    except Exception as e:
        logger.error(f"Error verifying semester: {e}")
        return True  # Continue processing anyway to avoid blocking

def get_available_semesters(result_path):
    """Get list of available semesters in the result files with comprehensive detection"""
    semesters = set()
    try:
        def extract_semester_from_item(item_name):
            """Extract semester from item name using comprehensive patterns"""
            item_upper = item_name.upper()
            
            # Comprehensive semester detection patterns
            semester_patterns = {
                "ND-First-YEAR-First-SEMESTER": [
                    "FIRST.YEAR.FIRST.SEMESTER", "FIRST-YEAR-FIRST-SEMESTER", "FIRST_YEAR_FIRST_SEMESTER",
                    "1ST.YEAR.1ST.SEMESTER", "1ST-YEAR-1ST-SEMESTER", "1ST_YEAR_1ST_SEMESTER",
                    "YEAR1.SEMESTER1", "YEAR-1-SEMESTER-1", "YEAR_1_SEMESTER_1"
                ],
                "ND-First-YEAR-SECOND-SEMESTER": [
                    "FIRST.YEAR.SECOND.SEMESTER", "FIRST-YEAR-SECOND-SEMESTER", "FIRST_YEAR_SECOND_SEMESTER",
                    "1ST.YEAR.2ND.SEMESTER", "1ST-YEAR-2ND-SEMESTER", "1ST_YEAR_2ND_SEMESTER",
                    "YEAR1.SEMESTER2", "YEAR-1-SEMESTER-2", "YEAR_1_SEMESTER_2"
                ],
                "ND-SECOND-YEAR-First-SEMESTER": [
                    "SECOND.YEAR.FIRST.SEMESTER", "SECOND-YEAR-FIRST-SEMESTER", "SECOND_YEAR_FIRST_SEMESTER",
                    "2ND.YEAR.1ST.SEMESTER", "2ND-YEAR-1ST-SEMESTER", "2ND_YEAR_1ST_SEMESTER",
                    "YEAR2.SEMESTER1", "YEAR-2-SEMESTER-1", "YEAR_2_SEMESTER_1"
                ],
                "ND-SECOND-YEAR-SECOND-SEMESTER": [
                    "SECOND.YEAR.SECOND.SEMESTER", "SECOND-YEAR-SECOND-SEMESTER", "SECOND_YEAR_SECOND_SEMESTER",
                    "2ND.YEAR.2ND.SEMESTER", "2ND-YEAR-2ND-SEMESTER", "2ND_YEAR_2ND_SEMESTER",
                    "YEAR2.SEMESTER2", "YEAR-2-SEMESTER-2", "YEAR_2_SEMESTER_2"
                ]
            }
            
            for semester, patterns in semester_patterns.items():
                for pattern in patterns:
                    flexible_pattern = pattern.replace('.', '[._\\- ]?')
                    if re.search(flexible_pattern, item_upper, re.IGNORECASE):
                        return semester
            
            return None
        
        if os.path.isdir(result_path):
            for item in os.listdir(result_path):
                semester = extract_semester_from_item(item)
                if semester:
                    semesters.add(semester)
        
        elif result_path.endswith('.zip'):
            with zipfile.ZipFile(result_path, 'r') as zip_ref:
                for item in zip_ref.namelist():
                    semester = extract_semester_from_item(item)
                    if semester:
                        semesters.add(semester)
        
        return list(semesters)
        
    except Exception as e:
        logger.error(f"Error getting available semesters: {e}")
        return []

# Routes
@app.route("/", methods=["GET"])
def index():
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    try:
        app.logger.info(f"Login route accessed - Method: {request.method}")
        logger.info(f"Attempting to load template: {os.path.join(TEMPLATE_DIR, 'login.html')}")
        if request.method == "POST":
            password = request.form.get("password")
            app.logger.info(f"Login attempt - password provided: {bool(password)}")
            if password == PASSWORD:
                session["logged_in"] = True
                flash("Successfully logged in!", "success")
                app.logger.info("Login successful")
                return redirect(url_for("dashboard"))
            else:
                flash("Invalid password. Please try again.", "error")
                app.logger.warning("Invalid password attempt")
        return render_template(
            "login.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None
        )
    except TemplateNotFound as e:
        app.logger.error(f"Template not found: {e} at {TEMPLATE_DIR}")
        logger.error(f"Template error: {e} at {TEMPLATE_DIR}")
        return f"<h1>Template Error</h1><p>Template not found: login.html at {TEMPLATE_DIR}</p><p>Available templates: {os.listdir(TEMPLATE_DIR) if os.path.exists(TEMPLATE_DIR) else 'None'}</p>", 500
    except Exception as e:
        app.logger.error(f"Login error: {e}")
        logger.error(f"Login error: {e}")
        return f"<h1>Server Error</h1><p>{str(e)}</p>", 500

# ============================================================================
# FIXED: Dashboard route - UPDATED VERSION
# ============================================================================
@app.route("/dashboard")
@login_required
def dashboard():
    try:
        carryover_summaries = {}
        for program in ["ND", "BN", "BM"]:
            program_sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
            carryover_summaries[program] = {}
            for set_name in program_sets:
                # CRITICAL FIX: Check if CLEAN_RESULTS exists before trying to get summary
                clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
                if not os.path.exists(clean_dir):
                    logger.info(f"⚠️ Skipping {program}/{set_name} - CLEAN_RESULTS not found")
                    continue
                
                summary = get_carryover_summary(program, set_name)
                if summary['total_students'] > 0:
                    carryover_summaries[program][set_name] = summary
                    logger.info(f"✅ Added carryover summary for {program}/{set_name}: {summary['total_students']} students")
        
        # Check internal exam files status for dashboard
        internal_input_dir = get_input_directory('clean')
        internal_files_exist = check_internal_exam_files(internal_input_dir)
        
        return render_template(
            "dashboard.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            carryover_summaries=carryover_summaries,
            internal_files_exist=internal_files_exist,
            internal_input_dir=internal_input_dir
        )
    except TemplateNotFound as e:
        logger.error(f"Dashboard template not found: {e}")
        flash("Dashboard template not found.", "error")
        return redirect(url_for("login"))
    except Exception as e:
        app.logger.error(f"Dashboard error: {e}")
        flash(f"Error loading dashboard: {str(e)}", "error")
        return redirect(url_for("login"))

@app.route("/debug_paths")
@login_required
def debug_paths():
    try:
        paths = {
            "Project Root": PROJECT_ROOT,
            "Script Directory": SCRIPT_DIR,
            "Base Directory": BASE_DIR,
            "Template Directory": TEMPLATE_DIR,
            "Static Directory": STATIC_DIR,
        }
        path_status = {key: os.path.exists(path) for key, path in paths.items()}
        templates = os.listdir(TEMPLATE_DIR) if os.path.exists(TEMPLATE_DIR) else []
        script_paths = {name: os.path.join(SCRIPT_DIR, path) for name, path in SCRIPT_MAP.items()}
        script_status = {name: os.path.exists(path) for name, path in script_paths.items()}
        env_vars = {
            "FLASK_SECRET": os.getenv("FLASK_SECRET", "Not set"),
            "STUDENT_CLEANER_PASSWORD": os.getenv("STUDENT_CLEANER_PASSWORD", "Not set"),
            "COLLEGE_NAME": COLLEGE,
            "DEPARTMENT": DEPARTMENT,
            "BASE_DIR": BASE_DIR,
        }
        return render_template(
            "debug_paths.html",
            environment="Railway Production" if not is_local_environment() else "Local Development",
            paths=paths,
            path_status=path_status,
            templates=templates,
            script_paths=script_paths,
            script_status=script_status,
            env_vars=env_vars,
            college=COLLEGE,
            department=DEPARTMENT,
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            nd_sets=ND_SETS,
            bn_sets=BN_SETS,
            bm_sets=BM_SETS,
            programs=PROGRAMS
        )
    except TemplateNotFound as e:
        logger.error(f"Debug paths template not found: {e}")
        flash("Debug paths template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"Debug paths error: {e}")
        flash(f"Error loading debug page: {str(e)}", "error")
        return redirect(url_for("dashboard"))

@app.route("/debug_internal_files")
@login_required
def debug_internal_files():
    """Debug route to check internal exam files detection"""
    input_dir = os.path.join(BASE_DIR, "OBJ_RESULT", "RAW_OBJ")
    files_info = {
        'input_dir': input_dir,
        'exists': os.path.exists(input_dir),
        'files_found': [],
        'check_internal_exam_files_result': check_internal_exam_files(input_dir)
    }
    
    if os.path.exists(input_dir):
        # Method 1: glob method (same as processing script)
        csv_files = glob.glob(os.path.join(input_dir, "*.csv"))
        xls_files = glob.glob(os.path.join(input_dir, "*.xls")) + glob.glob(os.path.join(input_dir, "*.xlsx"))
        all_files = [f for f in (csv_files + xls_files) if not os.path.basename(f).startswith("~$")]
        
        # Method 2: Pattern-based method
        pattern_files = [
            f for f in os.listdir(input_dir) 
            if f.lower().endswith((".xlsx", ".xls", ".csv")) 
            and not f.startswith("~")
            and (
                f.startswith("Set") or
                "ND" in f.upper() and "SET" in f.upper() or
                "OBJ" in f.upper() or
                "RESULT" in f.upper()
            )
        ]
        
        files_info['glob_method_files'] = [os.path.basename(f) for f in all_files]
        files_info['pattern_method_files'] = pattern_files
        files_info['all_files_in_dir'] = os.listdir(input_dir)
        
        # Final combined list
        final_files = list(set(all_files + [os.path.join(input_dir, f) for f in pattern_files]))
        files_info['final_files'] = [os.path.basename(f) for f in final_files]
        files_info['files_found'] = files_info['final_files']
    
    return jsonify(files_info)

@app.route("/debug_dir_contents")
@login_required
def debug_dir_contents():
    try:
        dir_contents = {}
        for dir_path in [BASE_DIR]:
            contents = {}
            if os.path.exists(dir_path):
                for root, dirs, files in os.walk(dir_path):
                    relative_path = os.path.relpath(root, dir_path)
                    contents[relative_path or "."] = {
                        "dirs": dirs,
                        "files": [f for f in files if f.lower().endswith((".xlsx", ".csv", ".pdf", ".zip"))]
                    }
            dir_contents[dir_path] = contents
        return jsonify(dir_contents)
    except Exception as e:
        app.logger.error(f"Debug dir contents error: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/upload_center", methods=["GET", "POST"])
@login_required
def upload_center():
    try:
        if request.method == "GET":
            return render_template(
                "upload_center.html",
                college=COLLEGE,
                department=DEPARTMENT,
                environment="Railway Production" if not is_local_environment() else "Local Development",
                logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
                nd_sets=ND_SETS,
                bn_sets=BN_SETS,
                bm_sets=BM_SETS,
                programs=PROGRAMS
            )
        return redirect(url_for("dashboard"))
    except TemplateNotFound as e:
        logger.error(f"Upload center template not found: {e}")
        flash("Upload center template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"Upload center error: {e}")
        flash(f"Error loading upload center: {str(e)}", "error")
        return redirect(url_for("dashboard"))

@app.route("/handle_upload", methods=["POST"])
@login_required
def handle_upload():
    """Handle file uploads from the upload center"""
    try:
        program = request.form.get("program")
        files = request.files.getlist("files")
        
        if not program:
            flash("Please select a program.", "error")
            return redirect(url_for("upload_center"))
            
        if not files or all(file.filename == '' for file in files):
            flash("Please select at least one file to upload.", "error")
            return redirect(url_for("upload_center"))
        
        program_map = {
            "nd": ("exam_processor_nd", "ND"),
            "bn": ("exam_processor_bn", "BN"), 
            "bm": ("exam_processor_bm", "BM")
        }
        
        if program not in program_map:
            flash("Invalid program selected.", "error")
            return redirect(url_for("upload_center"))
            
        script_name, program_code = program_map[program]
        
        set_name = None
        if program == "nd":
            set_name = request.form.get("nd_set")
        elif program == "bn":
            set_name = request.form.get("bn_set")
        elif program == "bm":
            set_name = request.form.get("bm_set")
            
        if not set_name:
            flash(f"Please select a {program.upper()} set.", "error")
            return redirect(url_for("upload_center"))
        
        raw_dir = get_raw_directory(script_name, program_code, set_name)
        logger.info(f"Uploading to directory: {raw_dir}")
        
        os.makedirs(raw_dir, exist_ok=True)
        
        saved_files = []
        skipped_files = []
        
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(raw_dir, filename)
                
                file.save(file_path)
                saved_files.append(filename)
                logger.info(f"Saved file: {file_path}")
                
                if filename.lower().endswith(".zip"):
                    try:
                        with zipfile.ZipFile(file_path, "r") as zip_ref:
                            zip_ref.extractall(raw_dir)
                        logger.info(f"Extracted ZIP file: {filename}")
                        saved_files.append(f"{filename} (extracted)")
                    except zipfile.BadZipFile:
                        logger.error(f"Invalid ZIP file: {filename}")
                        skipped_files.append(f"{filename} (invalid ZIP)")
                    except Exception as e:
                        logger.error(f"Error extracting ZIP {filename}: {e}")
                        skipped_files.append(f"{filename} (extraction error)")
            else:
                skipped_files.append(file.filename if file.filename else "unknown file")
        
        if saved_files:
            success_msg = f"Successfully uploaded {len(saved_files)} file(s) to {program_code}/{set_name}/RAW_RESULTS"
            if skipped_files:
                success_msg += f" (Skipped {len(skipped_files)} invalid files)"
            flash(success_msg, "success")
            logger.info(f"Upload completed: {success_msg}")
        else:
            flash("No valid files were uploaded. Please check file formats.", "error")
        
        if skipped_files:
            flash(f"Skipped files: {', '.join(skipped_files)}", "warning")
            
        return redirect(url_for("upload_center"))
        
    except Exception as e:
        app.logger.error(f"Upload error: {e}")
        flash(f"Upload failed: {str(e)}", "error")
        return redirect(url_for("upload_center"))

# ============================================================================
# FIXED: handle_resit_upload function - UPDATED VERSION
# ============================================================================
@app.route("/handle_resit_upload", methods=["POST"])
@login_required
def handle_resit_upload():
    """Handle dynamic set selection and multiple semesters - FIXED"""
    try:
        logger.info("CARRYOVER UPLOAD: Route called")
        
        # FIX: Get program from form explicitly
        program = request.form.get("program", "").lower()
        set_name = request.form.get("nd_set") or request.form.get("bn_set") or request.form.get("bm_set")
        selected_semesters = request.form.getlist("selected_semesters")
        resit_file = request.files.get("resit_file")
        
        logger.info(f"Received - Program: {program}, Set: {set_name}, Semesters: {selected_semesters}, File: {resit_file.filename if resit_file else 'None'}")
        
        # Validation
        if not resit_file or resit_file.filename == '':
            flash("Please select a file", "error")
            return redirect(url_for("upload_center"))
        
        if not program or program == "":
            flash("Please select a program", "error")
            return redirect(url_for("upload_center"))
        
        if not set_name or set_name == "unknown":
            flash("Please select a set", "error")
            return redirect(url_for("upload_center"))
        
        if not selected_semesters:
            flash("Please select at least one semester", "error")
            return redirect(url_for("upload_center"))
        
        # Convert program to uppercase
        program_map = {"nd": "ND", "bn": "BN", "bm": "BM"}
        program_code = program_map.get(program, "ND")
        
        # CRITICAL FIX: Correct path construction
        raw_dir = os.path.join(BASE_DIR, program_code, set_name, "RAW_RESULTS", "CARRYOVER")
        os.makedirs(raw_dir, exist_ok=True)
        logger.info(f"Target directory: {raw_dir}")
        
        # Save file
        filename = secure_filename(resit_file.filename)
        file_path = os.path.join(raw_dir, filename)
        resit_file.save(file_path)
        
        logger.info(f"File saved: {file_path}")
        
        # Verify file was saved
        if os.path.exists(file_path):
            file_size = os.path.getsize(file_path)
            logger.info(f"✅ Resit file saved successfully: {file_path} ({file_size} bytes)")
            
            semester_display = ", ".join(selected_semesters)
            flash(f"Successfully uploaded carryover file to {program_code}/{set_name}/RAW_RESULTS/CARRYOVER for semesters: {semester_display}", "success")
        else:
            logger.error(f"❌ Resit file was not saved: {file_path}")
            flash("Failed to save resit file", "error")
            return redirect(url_for("upload_center"))
        
        return redirect(url_for("upload_center"))
        
    except Exception as e:
        logger.error(f"ERROR in handle_resit_upload: {str(e)}")
        import traceback
        logger.error(f"Stack trace: {traceback.format_exc()}")
        flash(f"Upload failed: {str(e)}", "error")
        return redirect(url_for("upload_center"))

@app.route("/download_center")
@login_required
def download_center():
    try:
        files_by_category = get_files_by_category()
        for category, files in files_by_category.items():
            if isinstance(files, dict):
                app.logger.info(f"Download center - {category}: {sum(len(f) for f in files.values())} ZIP files across {len(files)} sets")
            else:
                app.logger.info(f"Download center - {category}: {len(files)} ZIP files")
        return render_template(
            "download_center.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            files_by_category=files_by_category,
            nd_sets=ND_SETS,
            bn_sets=BN_SETS,
            bm_sets=BM_SETS,
            programs=PROGRAMS
        )
    except TemplateNotFound as e:
        logger.error(f"Download center template not found: {e}")
        flash("Download center template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"Download center error: {e}")
        flash(f"Error loading download center: {str(e)}", "error")
        return redirect(url_for("dashboard"))

@app.route("/file_browser")
@login_required
def file_browser():
    try:
        sets = get_sets_and_folders()
        
        nd_sets = []
        bn_sets = []
        bm_sets = []
        
        for key in sets.keys():
            if key.startswith('ND_'):
                nd_sets.append(key.replace('ND_', ''))
            elif key.startswith('BN_'):
                bn_sets.append(key.replace('BN_', ''))
            elif key.startswith('BM_'):
                bm_sets.append(key.replace('BM_', ''))
        
        app.logger.info(f"File browser - ND sets: {nd_sets}, BN sets: {bn_sets}, BM sets: {bm_sets}")
        app.logger.info(f"File browser - Total sets: {len(sets)}")
        
        for set_key, folders in sets.items():
            total_files = sum(len(folder.files) for folder in folders)
            app.logger.info(f"Set '{set_key}': {total_files} ZIP files across {len(folders)} folders")
        
        return render_template(
            "file_browser.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            sets=sets,
            nd_sets=nd_sets,
            bn_sets=bn_sets,
            bm_sets=bm_sets,
            programs=PROGRAMS,
            BASE_DIR=BASE_DIR,
            processed_dir=BASE_DIR
        )
    except TemplateNotFound as e:
        logger.error(f"File browser template not found: {e}")
        flash("File browser template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"File browser error: {e}")
        flash(f"Error loading file browser: {str(e)}", "error")
        return redirect(url_for("dashboard"))

# ============================================================================
# FIXED: Carryover route - UPDATED VERSION
# ============================================================================
@app.route("/carryover")
@login_required
def carryover():
    """Carryover student management dashboard"""
    try:
        carryover_data = {}
        
        for program in ["ND", "BN", "BM"]:
            program_sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
            carryover_data[program] = {}
            
            for set_name in program_sets:
                # CRITICAL FIX: Check if CLEAN_RESULTS exists before trying to get records
                clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
                if not os.path.exists(clean_dir):
                    logger.info(f"⚠️ Skipping {program}/{set_name} - CLEAN_RESULTS not found")
                    continue
                
                records = get_carryover_records(program, set_name)
                if records:
                    carryover_data[program][set_name] = {
                        'records': records,
                        'total_students': sum(record['count'] for record in records),
                        'total_semesters': len(records)
                    }
                    logger.info(f"✅ Added carryover data for {program}/{set_name}: {len(records)} records")
        
        return render_template(
            "carryover_management.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            carryover_data=carryover_data,
            nd_sets=ND_SETS,
            bn_sets=BN_SETS,
            bm_sets=BM_SETS,
            programs=PROGRAMS
        )
    except TemplateNotFound as e:
        logger.error(f"Carryover management template not found: {e}")
        flash("Carryover management template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"Carryover management error: {e}")
        flash(f"Error loading carryover management: {str(e)}", "error")
        return redirect(url_for("dashboard"))

# ============================================================================
# NEW: Diagnostic Route for Debugging
# ============================================================================
@app.route("/debug_carryover_files/<program>/<set_name>")
@login_required
def debug_carryover_files_detail(program, set_name):
    """Debug route to check carryover file semester matching - ENHANCED VERSION."""
    try:
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            return jsonify({'error': f"Clean directory not found: {clean_dir}"})
        
        debug_info = {
            'clean_dir': clean_dir,
            'items': [],
            'regular_results': [],
            'carryover_results': [],
            'carryover_records': [],
            'available_files': os.listdir(clean_dir)
        }
        
        # List all items with detailed classification
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
            item_info = {
                'name': item,
                'type': 'dir' if os.path.isdir(item_path) else 'file',
                'is_regular_result': (item.startswith(f"{set_name}_RESULT-") and 
                                    "RESULT" in item.upper() and 
                                    not "CARRYOVER" in item.upper()),
                'is_carryover_result': 'CARRYOVER' in item.upper(),
                'is_zip': item.endswith('.zip')
            }
            debug_info['items'].append(item_info)
            
            if item_info['is_regular_result']:
                debug_info['regular_results'].append(item)
            elif item_info['is_carryover_result']:
                debug_info['carryover_results'].append(item)
        
        # Check carryover records in latest REGULAR result (not carryover result)
        if debug_info['regular_results']:
            latest_regular = sorted(debug_info['regular_results'])[-1]
            debug_info['latest_regular'] = latest_regular
            latest_path = os.path.join(clean_dir, latest_regular)
            
            if latest_regular.endswith('.zip'):
                # Extract from ZIP
                try:
                    with zipfile.ZipFile(latest_path, 'r') as zip_ref:
                        # Look for carryover JSON files
                        json_files = [f for f in zip_ref.namelist() 
                                    if f.startswith("CARRYOVER_RECORDS/") and f.endswith('.json')]
                        
                        for json_file in json_files:
                            extracted = extract_semester_from_filename(json_file)
                            standardized = standardize_semester_key(extracted)
                            debug_info['carryover_records'].append({
                                'filename': json_file,
                                'extracted_semester': extracted,
                                'standardized_semester': standardized,
                                'source': f'ZIP: {latest_regular}'
                            })
                except Exception as e:
                    debug_info['zip_error'] = str(e)
            else:
                # Check directory
                carryover_dir = os.path.join(latest_path, "CARRYOVER_RECORDS")
                if os.path.exists(carryover_dir):
                    for file in os.listdir(carryover_dir):
                        if file.endswith('.json'):
                            extracted = extract_semester_from_filename(file)
                            standardized = standardize_semester_key(extracted)
                            debug_info['carryover_records'].append({
                                'filename': file,
                                'extracted_semester': extracted,
                                'standardized_semester': standardized,
                                'source': f'DIR: {latest_regular}'
                            })
        
        return jsonify(debug_info)
        
    except Exception as e:
        return jsonify({'error': str(e)})

# ============================================================================
# FIXED: process_resit function - COMPREHENSIVE FIXED VERSION
# ============================================================================
@app.route("/process_resit", methods=["POST"])
@login_required
def process_resit():
    """Process resit results for carryover students - FIXED."""
    try:
        logger.info("RESIT PROCESSING: Starting comprehensive processing")
        
        # Get form data
        program = request.form.get("resit_program", "").lower()
        set_name = request.form.get("resit_set", "").strip()
        semester_key = request.form.get("resit_semester", "").strip()
        resit_file = request.files.get("resit_file")
        
        logger.info(f"Processing resit for: {program}/{set_name}/{semester_key}")
        
        # Validation
        if not all([program, set_name, semester_key, resit_file]):
            missing = []
            if not program: missing.append("program")
            if not set_name: missing.append("set")
            if not semester_key: missing.append("semester")
            if not resit_file: missing.append("resit file")
            flash(f"Missing required fields: {', '.join(missing)}", "error")
            return redirect(url_for("carryover"))
        
        # Map program
        program_map = {"nd": "ND", "bn": "BN", "bm": "BM"}
        program_code = program_map.get(program, "ND")
        
        # CRITICAL FIX: Standardize semester key immediately
        semester_key = standardize_semester_key(semester_key)
        logger.info(f"📝 Standardized semester key: {semester_key}")
        
        # Save resit file
        resit_dir = os.path.join(BASE_DIR, program_code, set_name, "RAW_RESULTS", "CARRYOVER")
        os.makedirs(resit_dir, exist_ok=True)
        
        filename = secure_filename(resit_file.filename)
        resit_file_path = os.path.join(resit_dir, filename)
        resit_file.save(resit_file_path)
        
        if not os.path.exists(resit_file_path):
            flash("Failed to save resit file", "error")
            return redirect(url_for("carryover"))
        
        logger.info(f"✅ Resit file saved: {resit_file_path}")
        
        # CRITICAL FIX: Find the correct base results (NOT carryover results)
        clean_dir = os.path.join(BASE_DIR, program_code, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            flash(f"No clean results found for {program_code}/{set_name}. Process regular results first.", "error")
            return redirect(url_for("carryover"))
        
        # Look for REGULAR result files (not carryover)
        regular_results = []
        
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
            
            # Regular results (base processing)
            if item.startswith(f"{set_name}_RESULT-") and not item.startswith("CARRYOVER_"):
                if os.path.isdir(item_path) or item.endswith('.zip'):
                    regular_results.append(item)
                    logger.info(f"Found regular result: {item}")
        
        # CRITICAL: Use regular results, not carryover results
        if not regular_results:
            logger.error(f"No regular results found for {program_code}/{set_name}")
            logger.info(f"Available items: {os.listdir(clean_dir)}")
            flash(f"No regular results found for {program_code}/{set_name}. Please run regular processing first.", "error")
            return redirect(url_for("carryover"))
        
        # Sort and get the latest REGULAR result
        regular_results.sort()
        latest_regular = regular_results[-1]
        latest_regular_path = os.path.join(clean_dir, latest_regular)
        
        logger.info(f"📁 Using base result: {latest_regular}")
        logger.info(f"📁 Semester to process: {semester_key}")
        logger.info(f"📁 Resit file: {filename}")
        
        # Extract ZIP if needed
        temp_extract_dir = None
        if latest_regular.endswith('.zip'):
            logger.info("Extracting base result ZIP...")
            temp_extract_dir = os.path.join(clean_dir, f"TEMP_EXTRACT_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            os.makedirs(temp_extract_dir, exist_ok=True)
            
            try:
                with zipfile.ZipFile(latest_regular_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_extract_dir)
                base_result_path = temp_extract_dir
                logger.info(f"Extracted to: {temp_extract_dir}")
            except Exception as e:
                logger.error(f"ZIP extraction failed: {e}")
                if temp_extract_dir and os.path.exists(temp_extract_dir):
                    shutil.rmtree(temp_extract_dir)
                flash(f"Failed to extract base results: {str(e)}", "error")
                return redirect(url_for("carryover"))
        else:
            base_result_path = latest_regular_path
        
        # Verify the base result contains the semester we're trying to process
        if not verify_semester_in_results(base_result_path, semester_key):
            logger.warning(f"Semester {semester_key} not found in base results. Available: {get_available_semesters(base_result_path)}")
            # Continue anyway - the processor might handle this
        
        # Setup environment
        env = os.environ.copy()
        env["BASE_DIR"] = BASE_DIR
        env["SELECTED_SET"] = set_name
        env["SELECTED_SEMESTERS"] = semester_key
        env["PASS_THRESHOLD"] = "50.0"
        env["RESIT_FILE_PATH"] = resit_file_path
        env["PROCESS_RESIT"] = "true"
        env["BASE_RESULT_PATH"] = base_result_path
        
        logger.info("Environment setup:")
        for key in ["BASE_DIR", "SELECTED_SET", "SELECTED_SEMESTERS", "RESIT_FILE_PATH", "BASE_RESULT_PATH"]:
            logger.info(f"  {key}: {env[key]}")
        
        # Run the carryover processor
        script_name = "integrated_carryover_processor"
        script_path = _get_script_path(script_name)
        
        logger.info(f"Running script: {script_path}")
        
        try:
            result = subprocess.run(
                [sys.executable, script_path],
                env=env,
                text=True,
                capture_output=True,
                timeout=600,
            )
        except subprocess.TimeoutExpired:
            logger.error("Carryover processing timed out")
            flash("Processing timed out after 10 minutes.", "error")
            return redirect(url_for("carryover"))
        
        # Clean up temporary directory
        if temp_extract_dir and os.path.exists(temp_extract_dir):
            try:
                shutil.rmtree(temp_extract_dir)
                logger.info("Cleaned up temporary extraction directory")
            except Exception as e:
                logger.warning(f"Failed to clean temp directory: {e}")
        
        # Parse results
        output_lines = result.stdout.splitlines()
        error_lines = result.stderr.splitlines()
        
        # Log outputs
        logger.info("=== SCRIPT OUTPUT ===")
        for line in output_lines:
            logger.info(f"OUT: {line}")
        
        if error_lines:
            logger.info("=== SCRIPT ERRORS ===")
            for line in error_lines:
                logger.error(f"ERR: {line}")
        
        # Check for specific success patterns
        success_indicators = [
            "Carryover processing completed successfully",
            "CARRYOVER PROCESSING COMPLETED",
            "Updated.*scores for.*students",
            "Resit processing completed",
        ]
        
        score_updates = 0
        students_updated = 0
        processing_successful = False
        
        for line in output_lines:
            # Check for success
            if any(indicator in line for indicator in success_indicators):
                processing_successful = True
            
            # Extract score updates
            if "Updated" in line and "scores" in line and "students" in line:
                match = re.search(r"Updated (\d+) scores for (\d+) students", line)
                if match:
                    score_updates = int(match.group(1))
                    students_updated = int(match.group(2))
                    logger.info(f"Found score updates: {score_updates} scores, {students_updated} students")
        
        # Check for new carryover files
        new_carryover_files = [f for f in os.listdir(clean_dir) 
                              if f.startswith("CARRYOVER_") and f.endswith(".zip")]
        latest_carryover = sorted(new_carryover_files)[-1] if new_carryover_files else None
        
        # Report results
        if result.returncode == 0 and processing_successful:
            if score_updates > 0:
                message = f"✅ Resit processing completed! Updated {score_updates} scores for {students_updated} students in {semester_key}"
                if latest_carryover:
                    message += f" | Results: {latest_carryover}"
                flash(message, "success")
            else:
                message = f"ℹ️ Resit processing completed but no scores updated for {semester_key}"
                if latest_carryover:
                    message += f" | Check: {latest_carryover}"
                flash(message, "info")
        else:
            error_msg = f"❌ Resit processing failed for {semester_key}"
            if error_lines:
                # Get the most relevant error
                for line in reversed(error_lines):
                    if line.strip() and not line.startswith("File") and not line.startswith("Traceback"):
                        error_msg += f": {line.strip()}"
                        break
            flash(error_msg, "error")
        
        return redirect(url_for("carryover"))
        
    except Exception as e:
        logger.error(f"RESIT PROCESSING ERROR: {str(e)}", exc_info=True)
        flash(f"Resit processing failed: {str(e)}", "error")
        return redirect(url_for("carryover"))

@app.route("/debug_resit_files/<program>/<set_name>")
@login_required
def debug_resit_files(program, set_name):
    """Debug route to check generated resit files"""
    try:
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            return jsonify({'error': f"Clean directory not found: {clean_dir}"})
        
        resit_folders = [f for f in os.listdir(clean_dir) 
                        if f.startswith("RESIT_") and os.path.isdir(os.path.join(clean_dir, f))]
        
        debug_info = {
            'clean_dir': clean_dir,
            'resit_folders': [],
            'resit_zips': []
        }
        
        for folder in resit_folders:
            folder_path = os.path.join(clean_dir, folder)
            files = []
            for root, dirs, filenames in os.walk(folder_path):
                for file in filenames:
                    if file.lower().endswith(('.xlsx', '.csv', '.pdf')):
                        files.append({
                            'name': file,
                            'path': os.path.relpath(os.path.join(root, file), clean_dir),
                            'size': os.path.getsize(os.path.join(root, file))
                        })
            debug_info['resit_folders'].append({
                'name': folder,
                'files': files
            })
        
        zip_files = [f for f in os.listdir(clean_dir) 
                    if f.startswith("RESIT_") and f.endswith('.zip')]
        
        for zip_file in zip_files:
            zip_path = os.path.join(clean_dir, zip_file)
            debug_info['resit_zips'].append({
                'name': zip_file,
                'size': os.path.getsize(zip_path),
                'path': zip_path
            })
        
        return jsonify(debug_info)
        
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route("/run_script/<script_name>", methods=["GET", "POST"])
@login_required
def run_script(script_name):
    try:
        processing_type = request.form.get("processing_type", "regular")
        if processing_type == "carryover":
            script_name = "integrated_carryover_processor"

        if script_name not in SCRIPT_MAP:
            flash("Invalid script requested.", "error")
            return redirect(url_for("dashboard"))
            
        program = script_name.split("_")[-1].upper() if script_name.startswith("exam_processor") or script_name == "integrated_carryover_processor" else None
        script_desc = {
            "utme": "PUTME Examination Results",
            "caosce": "CAOSCE Examination Results",
            "clean": "Internal Examination Results",
            "split": "JAMB Candidate Name Split",
            "exam_processor_nd": "ND Examination Results Processing",
            "exam_processor_bn": "Basic Nursing Examination Results Processing",
            "exam_processor_bm": "Basic Midwifery Examination Results Processing",
            "integrated_carryover_processor": "Carryover & Resit Processing",
        }.get(script_name, "Script")
        
        script_path = _get_script_path(script_name)
        
        selected_set = None
        if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm", "integrated_carryover_processor"]:
            if processing_type == "carryover":
                selected_set = request.form.get("carryover_set", "all")
            else:
                selected_set = request.form.get(
                    "selected_set" if script_name == "exam_processor_nd" or script_name == "integrated_carryover_processor" else
                    "nursing_set" if script_name == "exam_processor_bn" else
                    "midwifery_set", "all"
                )
            logger.info(f"Selected set for {script_name}: {selected_set}")
        
        # FIXED: Use the corrected get_input_directory function
        input_dir = get_input_directory(script_name, program, selected_set)
        logger.info(f"Final input directory for {script_name}: {input_dir}")
        
        if request.method == "GET":
            template_map = {
                "utme": "utme_form.html",
                "exam_processor_nd": "exam_processor_form.html",
                "exam_processor_bn": "basic_nursing_form.html",
                "exam_processor_bm": "basic_midwifery_form.html",
                "clean": "internal_exam_form.html",  # ADDED: Template for clean script
            }
            template = template_map.get(script_name)
            
            # For scripts that don't have a form template, redirect to dashboard
            if not template:
                flash(f"{script_desc} can be run directly from the dashboard.", "info")
                return redirect(url_for("dashboard"))
                
            if template:
                status_info = {}
                if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm", "integrated_carryover_processor"]:
                    status_info = get_exam_processor_status(program)
                
                # For clean script, check input files and provide status
                if script_name == "clean":
                    files_exist = check_input_files(input_dir, script_name, selected_set)
                    status_info = {
                        'ready': files_exist,
                        'input_dir': input_dir,
                        'files_found': files_exist
                    }
                
                return render_template(
                    template,
                    college=COLLEGE,
                    department=DEPARTMENT,
                    environment="Railway Production" if not is_local_environment() else "Local Development",
                    logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
                    selected_program=program,
                    programs=PROGRAMS,
                    nd_sets=ND_SETS,
                    bn_sets=BN_SETS,
                    bm_sets=BM_SETS,
                    status_info=status_info,
                    script_name=script_name,
                    script_desc=script_desc
                )
                
        if request.method == "POST":
            # Define env at the beginning for ALL script types
            env = os.environ.copy()
            env["BASE_DIR"] = BASE_DIR  # Add BASE_DIR to environment for all scripts
            
            # Check input files before processing
            if not check_input_files(input_dir, script_name, selected_set):
                flash(f"No valid input files found in {input_dir} for {script_desc}.", "error")
                return redirect(url_for("dashboard"))
            
            is_carryover_processing = processing_type == "carryover"
            
            if script_name == "utme":
                clean_dir = get_clean_directory(script_name)
                before_files = set(os.listdir(clean_dir)) if os.path.exists(clean_dir) else set()
                convert_value = request.form.get("convert_value", "").strip()
                convert_column = request.form.get("convert_column", "n")
                cmd = ["python3", script_path]
                if convert_value:
                    cmd.extend(["--non-interactive", "--converted-score-max", convert_value])
                result = subprocess.run(
                    cmd,
                    input=f"{convert_column}\n",
                    text=True,
                    capture_output=True,
                    check=True,
                    timeout=300,
                    env=env
                )
                output_lines = result.stdout.splitlines()
                processed_files = count_processed_files(output_lines, script_name)
                success_msg = get_success_message(script_name, processed_files, output_lines)
                if success_msg:
                    flash(success_msg, "success")
                else:
                    flash(f"No files processed for {script_desc}. Check input files in {input_dir}.", "error")
                if result.returncode == 0:
                    after_files = set(os.listdir(clean_dir))
                    new_files = after_files - before_files
                    if new_files:
                        zip_filename = f"utme_processed_{datetime.now().strftime('%d%m%Y_%H%M%S')}.zip"
                        zip_path = os.path.join(clean_dir, zip_filename)
                        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_f:
                            for file in new_files:
                                file_path = os.path.join(clean_dir, file)
                                zip_f.write(file_path, file)
                        # Clean up scattered files
                        cleanup_scattered_files(clean_dir, zip_filename)
                        flash(f"Zipped results ready: {zip_filename}", "success")
                return redirect(url_for("dashboard"))
                
            elif script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm", "integrated_carryover_processor"]:
                if is_carryover_processing:
                    selected_set = request.form.get("carryover_set", "all")
                    selected_semesters = request.form.get("carryover_semester", "")
                    env["SELECTED_SEMESTERS"] = selected_semesters
                else:
                    selected_set = request.form.get(
                        "selected_set" if script_name == "exam_processor_nd" or script_name == "integrated_carryover_processor" else
                        "nursing_set" if script_name == "exam_processor_bn" else
                        "midwifery_set", "all"
                    )
                    selected_semesters = request.form.getlist('selected_semesters')
                    if not selected_semesters or 'all' in selected_semesters:
                        selected_semesters = ['all']
                    env["SELECTED_SEMESTERS"] = ','.join(selected_semesters)
                
                pass_threshold = request.form.get("pass_threshold", "50.0")
                upgrade_threshold = request.form.get("upgrade_threshold", "0")
                generate_pdf = "generate_pdf" in request.form
                track_withdrawn = "track_withdrawn" in request.form
                
                selected_semesters = []
                if is_carryover_processing:
                    carryover_semester = request.form.get('carryover_semester')
                    if carryover_semester:
                        selected_semesters = [carryover_semester]
                        logger.info(f"CARRYOVER PROCESSING: Single semester selected: {carryover_semester}")
                    else:
                        flash("Please select a semester for carryover processing.", "error")
                        return redirect(url_for("dashboard"))
                else:
                    selected_semesters = request.form.getlist('selected_semesters')
                    if not selected_semesters or 'all' in selected_semesters:
                        selected_semesters = ['all']
                
                env["SELECTED_SET"] = selected_set
                env["SELECTED_SEMESTERS"] = ','.join(selected_semesters)
                env["PASS_THRESHOLD"] = pass_threshold
                
                if is_carryover_processing:
                    env["PROCESS_RESIT"] = "true"
                    resit_file = request.files.get('carryover_file')
                    if resit_file and resit_file.filename:
                        if "ND-" in selected_set:
                            program = "ND"
                        elif "SET4" in selected_set:
                            program = "BN"
                        else:
                            program = "BM"
                        
                        carryover_dir = os.path.join(BASE_DIR, program, selected_set, "RAW_RESULTS", "CARRYOVER")
                        os.makedirs(carryover_dir, exist_ok=True)
                        filename = secure_filename(resit_file.filename)
                        resit_path = os.path.join(carryover_dir, filename)
                        resit_file.save(resit_path)
                        env["RESIT_FILE_PATH"] = resit_path
                        logger.info(f"Resit file saved: {resit_path}")
                    else:
                        flash("Please upload a resit file for carryover processing.", "error")
                        return redirect(url_for("dashboard"))
                
                if upgrade_threshold and upgrade_threshold.strip() and upgrade_threshold != "0":
                    env["UPGRADE_THRESHOLD"] = upgrade_threshold.strip()
                env["GENERATE_PDF"] = str(generate_pdf)
                env["TRACK_WITHDRAWN"] = str(track_withdrawn)
                
                logger.info(f"Running {script_name} with environment:")
                logger.info(f"  BASE_DIR: {env['BASE_DIR']}")
                logger.info(f"  SELECTED_SET: {env['SELECTED_SET']}")
                logger.info(f"  SELECTED_SEMESTERS: {env['SELECTED_SEMESTERS']}")
                logger.info(f"  PASS_THRESHOLD: {env['PASS_THRESHOLD']}")
                logger.info(f"  PROCESS_RESIT: {env.get('PROCESS_RESIT', 'False')}")
                logger.info(f"  RESIT_FILE_PATH: {env.get('RESIT_FILE_PATH', 'Not set')}")
                logger.info(f"  UPGRADE_THRESHOLD: {env.get('UPGRADE_THRESHOLD', 'Not set')}")
                
                result = subprocess.run(
                    [sys.executable, script_path],
                    env=env,
                    text=True,
                    capture_output=True,
                    timeout=600,
                )
                
                logger.info("=== SCRIPT STDOUT ===")
                for line in result.stdout.splitlines():
                    logger.info(line)
                logger.info("=== SCRIPT STDERR ===")
                for line in result.stderr.splitlines():
                    logger.info(line)
                    
                output_lines = result.stdout.splitlines()
                processed_files = count_processed_files(output_lines, script_name, selected_set)
                success_msg = get_success_message(script_name, processed_files, output_lines, selected_set)
                
                # UPDATED CODE BLOCK WITH AUTOMATIC ZIPPING FOR ALL FILES
                if result.returncode == 0:
                    if success_msg:
                        flash(success_msg, "success")
                    else:
                        flash(f"{script_desc} completed but no files were processed.", "warning")
                    
                    clean_dir = get_clean_directory(script_name, program, selected_set)
                    if os.path.exists(clean_dir):
                        # Check if script already created a ZIP
                        existing_zips = [f for f in os.listdir(clean_dir) 
                                        if f.startswith(f"{selected_set}_RESULT-") and f.endswith('.zip')]
                        
                        if existing_zips:
                            # Script already created ZIP - just notify and clean up scattered files
                            latest_zip = sorted(existing_zips)[-1]
                            flash(f"Results already zipped by processor: {latest_zip}", "info")
                            logger.info(f"Script already created ZIP: {latest_zip}")
                            # Clean up any remaining scattered files
                            cleanup_scattered_files(clean_dir, latest_zip)
                        else:
                            # No ZIP exists - create one (fallback)
                            result_dirs = [d for d in os.listdir(clean_dir) 
                                         if d.startswith(f"{selected_set}_RESULT-") and os.path.isdir(os.path.join(clean_dir, d))]
                            
                            if result_dirs:
                                latest = sorted(result_dirs)[-1]
                                folder_path = os.path.join(clean_dir, latest)
                                
                                all_files = []
                                for root, dirs, files in os.walk(folder_path):
                                    for file in files:
                                        if file.lower().endswith(('.xlsx', '.csv', '.pdf', '.zip')):
                                            all_files.append(os.path.join(root, file))
                                
                                if all_files:
                                    zip_filename = f"{selected_set}_RESULT-{latest.split('-')[-1]}.zip"
                                    zip_path = os.path.join(clean_dir, zip_filename)
                                    
                                    try:
                                        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_f:
                                            for file_path in all_files:
                                                arcname = os.path.relpath(file_path, folder_path)
                                                zip_f.write(file_path, arcname)
                                        
                                        if os.path.exists(zip_path) and os.path.getsize(zip_path) > 100:
                                            flash(f"Zipped results ready: {zip_filename} ({len(all_files)} files)", "success")
                                            logger.info(f"Created fallback ZIP: {zip_path}")
                                            # Clean up ALL scattered files
                                            cleanup_scattered_files(clean_dir, zip_filename)
                                        else:
                                            logger.warning(f"ZIP file created but appears empty: {zip_path}")
                                            flash(f"ZIP created but appears empty for {selected_set}", "warning")
                                            
                                    except Exception as e:
                                        logger.error(f"Failed to create fallback ZIP: {e}")
                else:
                    flash(f"Script failed: {result.stderr or 'Unknown error'}", "error")
                # END OF UPDATED CODE BLOCK
                
                return redirect(url_for("dashboard"))
                
            else:
                # For clean, caosce, split scripts - use base environment
                logger.info(f"Running {script_name} with base environment")
                
                clean_dir = get_clean_directory(script_name)
                before_files = set(os.listdir(clean_dir)) if os.path.exists(clean_dir) else set()
                
                result = subprocess.run(
                    [sys.executable, script_path],
                    env=env,
                    text=True,
                    capture_output=True,
                    timeout=300,
                )
                
                output_lines = result.stdout.splitlines()
                processed_files = count_processed_files(output_lines, script_name)
                success_msg = get_success_message(script_name, processed_files, output_lines)
                
                if result.returncode == 0:
                    if success_msg:
                        flash(success_msg, "success")
                    else:
                        flash(f"No files processed for {script_desc}. Check input files in {input_dir}.", "error")
                    
                    # Add zipping for clean script
                    if script_name == "clean":
                        timestamp_folders = [f for f in os.listdir(clean_dir) 
                                             if f.startswith("obj_result_") and os.path.isdir(os.path.join(clean_dir, f))]
                        if timestamp_folders:
                            latest = sorted(timestamp_folders)[-1]
                            folder_path = os.path.join(clean_dir, latest)
                            zip_filename = f"{latest}.zip"
                            zip_path = os.path.join(clean_dir, zip_filename)
                            
                            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_f:
                                for root, _, files in os.walk(folder_path):
                                    for file in files:
                                        file_path = os.path.join(root, file)
                                        arcname = os.path.relpath(file_path, clean_dir)
                                        zip_f.write(file_path, arcname)
                            
                            if os.path.exists(zip_path):
                                flash(f"Zipped results ready: {zip_filename}", "success")
                                # Clean up ALL scattered files
                                cleanup_scattered_files(clean_dir, zip_filename)
                    else:
                        # For caosce and split
                        after_files = set(os.listdir(clean_dir))
                        new_files = after_files - before_files
                        if new_files:
                            zip_prefix = script_name + "_processed"
                            zip_filename = f"{zip_prefix}_{datetime.now().strftime('%d%m%Y_%H%M%S')}.zip"
                            zip_path = os.path.join(clean_dir, zip_filename)
                            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_f:
                                for file in new_files:
                                    file_path = os.path.join(clean_dir, file)
                                    zip_f.write(file_path, file)
                            # Clean up scattered files
                            cleanup_scattered_files(clean_dir, zip_filename)
                            flash(f"Zipped results ready: {zip_filename}", "success")
                else:
                    flash(f"Script failed: {result.stderr or 'Unknown error'}", "error")
                    
                return redirect(url_for("dashboard"))
        
    except FileNotFoundError as e:
        app.logger.error(f"Script file error: {e}")
        flash(f"Script not found for {script_desc}: {str(e)}", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"Run script error: {e}")
        flash(f"Server error processing {script_desc}: {str(e)}", "error")
        return redirect(url_for("dashboard"))

@app.route("/upload/<script_name>", methods=["POST"])
@login_required
def upload_files(script_name):
    try:
        if script_name not in SCRIPT_MAP:
            flash("Invalid script requested.", "error")
            return redirect(url_for("upload_center"))
            
        program = script_name.split("_")[-1].upper() if script_name.startswith("exam_processor") or script_name == "integrated_carryover_processor" else None
        script_desc = {
            "utme": "PUTME Results",
            "caosce": "CAOSCE Results",
            "clean": "Internal Examinations",
            "split": "JAMB Database",
            "exam_processor_nd": "ND Examinations",
            "exam_processor_bn": "Basic Nursing",
            "exam_processor_bm": "Basic Midwifery",
            "integrated_carryover_processor": "Carryover & Resit Files",
        }.get(script_name, "Files")
        
        files = request.files.getlist("files")
        candidate_files = request.files.getlist("candidate_files") if script_name == "utme" else []
        course_files = request.files.getlist("course_files") if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] else []
        set_name = request.form.get("nd_set") or request.form.get("nursing_set") or request.form.get("midwifery_set")
        
        if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] and set_name:
            raw_dir = get_raw_directory(script_name, program, set_name)
            logger.info(f"Uploading to exam processor raw directory: {raw_dir}")
        elif script_name == "integrated_carryover_processor" and set_name:
            raw_dir = os.path.join(BASE_DIR, program, set_name, "RAW_RESULTS", "CARRYOVER")
            logger.info(f"Uploading to carryover raw directory: {raw_dir}")
        else:
            raw_dir = get_raw_directory(script_name)
            logger.info(f"Uploading to other script raw directory: {raw_dir}")
        
        os.makedirs(raw_dir, exist_ok=True)
        saved_files = []
        
        if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] and course_files:
            course_dir = os.path.join(BASE_DIR, program, f"{program}-COURSES")
            os.makedirs(course_dir, exist_ok=True)
            for file in course_files:
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(course_dir, filename)
                    file.save(file_path)
                    saved_files.append(f"course: {filename}")
                    logger.info(f"Saved course file: {file_path}")
        
        for file in files + candidate_files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(raw_dir, filename)
                file.save(file_path)
                saved_files.append(filename)
                
                if filename.lower().endswith(".zip"):
                    try:
                        with zipfile.ZipFile(file_path, "r") as z:
                            z.extractall(raw_dir)
                        os.remove(file_path)
                        flash(f"Extracted ZIP: {filename}", "success")
                    except zipfile.BadZipFile:
                        flash(f"Invalid ZIP file: {filename}", "error")
                        return redirect(url_for("upload_center"))
                        
        if saved_files:
            flash(f"Uploaded files to {raw_dir}: {', '.join(saved_files)}", "success")
        else:
            flash("No valid files uploaded.", "error")
            return redirect(url_for("upload_center"))
        
        flash(f"Files uploaded successfully to RAW directory. You can now process them from the dashboard.", "success")
        return redirect(url_for("dashboard"))
        
    except Exception as e:
        app.logger.error(f"Upload files error: {e}")
        flash(f"Upload failed: {str(e)}", "error")
        return redirect(url_for("upload_center"))

@app.route("/download/<path:filename>")
@login_required
def download(filename):
    try:
        safe_name = os.path.basename(filename)
        for root, _, files in os.walk(BASE_DIR):
            if safe_name in files:
                return send_from_directory(root, safe_name, as_attachment=True)
        flash(f"File '{safe_name}' not found.", "error")
        return redirect(url_for("download_center"))
    except Exception as e:
        app.logger.error(f"Download error: {e}")
        flash(f"Download failed: {str(e)}", "error")
        return redirect(url_for("download_center"))

@app.route("/download_file/<path:filename>")
@login_required
def download_file(filename):
    return download(filename)

@app.route("/download_zip/<set_name>")
@login_required
def download_zip(set_name):
    try:
        zip_path = None
        for program in ["ND", "BN", "BM"]:
            sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
            if set_name in sets:
                clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
                if os.path.exists(clean_dir):
                    zip_files = [f for f in os.listdir(clean_dir) 
                               if f.startswith(f"{set_name}_RESULT-") and f.endswith('.zip')]
                    
                    if zip_files:
                        latest_zip = sorted(zip_files)[-1]
                        zip_path = os.path.join(clean_dir, latest_zip)
                        break
        
        if zip_path and os.path.exists(zip_path):
            file_size = os.path.getsize(zip_path)
            if file_size > 100:
                return send_file(zip_path, as_attachment=True)
            else:
                flash(f"ZIP file for '{set_name}' is empty or corrupted.", "error")
                return redirect(url_for("download_center"))
        else:
            flash(f"Set '{set_name}' not found or has no results.", "error")
            return redirect(url_for("download_center"))
    except Exception as e:
        app.logger.error(f"Download ZIP error: {e}")
        flash(f"Failed to create ZIP for {set_name}: {str(e)}", "error")
        return redirect(url_for("download_center"))

# ============================================================================
# FIXED: delete function - CORRECTED VERSION
# ============================================================================
@app.route("/delete/<path:filename>", methods=["POST"])
@login_required
def delete(filename):
    try:
        critical_dirs = [SCRIPT_DIR, TEMPLATE_DIR, STATIC_DIR, PROJECT_ROOT]
        
        file_path = None
        for root, _, files in os.walk(BASE_DIR):
            if os.path.basename(filename) in files:
                candidate = os.path.join(root, os.path.basename(filename))
                if os.path.exists(candidate):
                    file_path = candidate
                    break
        
        if not file_path:
            for root, dirs, _ in os.walk(BASE_DIR):
                if os.path.basename(filename) in dirs:
                    candidate = os.path.join(root, os.path.basename(filename))
                    if os.path.exists(candidate):
                        file_path = candidate
                        break
        
        if not file_path:
            flash(f"Path '{filename}' not found.", "error")
            logger.warning(f"Deletion failed: Path not found - {filename}")
            return redirect(request.referrer or url_for("file_browser"))
        
        abs_file_path = os.path.abspath(file_path)
        for critical_dir in critical_dirs:
            abs_critical_dir = os.path.abspath(critical_dir)
            if (abs_file_path == abs_critical_dir or 
                os.path.dirname(abs_file_path) == abs_critical_dir):
                flash(f"Cannot delete critical system path: {filename}", "error")
                logger.warning(f"Deletion blocked: Attempted to delete critical system path - {filename}")
                return redirect(request.referrer or url_for("file_browser"))
        
        if request.form.get("confirm") != "true":
            flash(f"Deletion of '{filename}' requires confirmation.", "warning")
            logger.info(f"Deletion of {filename} requires confirmation")
            return redirect(request.referrer or url_for("file_browser"))
        
        if os.path.isfile(file_path):
            os.remove(file_path)
            flash(f"File '{filename}' deleted successfully.", "success")
            logger.info(f"Deleted file: {file_path}")
        elif os.path.isdir(file_path):
            shutil.rmtree(file_path, ignore_errors=True)
            flash(f"Folder '{filename}' deleted successfully.", "success")
            logger.info(f"Deleted folder: {file_path}")
        else:
            flash(f"Path '{filename}' does not exist.", "error")
            logger.warning(f"Deletion failed: Path does not exist - {file_path}")
        
        return redirect(request.referrer or url_for("file_browser"))
    except Exception as e:
        app.logger.error(f"Delete error: {e}")
        logger.error(f"Delete error for {filename}: {e}")
        flash(f"Failed to delete '{filename}': {str(e)}", "error")
        return redirect(request.referrer or url_for("file_browser"))

@app.route("/delete_file/<path:filename>", methods=["POST"])
@login_required
def delete_file(filename):
    return delete(filename)

@app.route("/logout")
@login_required
def logout():
    session.pop("logged_in", None)
    flash("You have been logged out.", "success")
    return redirect(url_for("login"))

@app.route("/run_script/integrated_carryover_processor", methods=["GET"])
@login_required
def handle_invalid_carryover_access():
    """Handle GET requests to integrated_carryover_processor with redirect and flash message"""
    flash("This endpoint requires a POST request. Please use the carryover form.", "error")
    return redirect(url_for("carryover"))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    mode = "local" if is_local_environment() else "cloud"
    logger.info(f"Starting Flask app in {mode.upper()} mode on port {port}...")
    app.run(host="0.0.0.0", port=port, debug=True)