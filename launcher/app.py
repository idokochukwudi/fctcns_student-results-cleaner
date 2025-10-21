import os
import subprocess
import re
import sys
import zipfile
import shutil
import time
import socket
import logging
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
BASE_DIR = os.path.join(PROJECT_ROOT, "EXAMS_INTERNAL")

# Launcher-specific directories
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "templates")
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")

# Log paths for verification
logger.info(f"🔍 PROJECT_ROOT: {PROJECT_ROOT}")
logger.info(f"🔍 SCRIPT_DIR: {SCRIPT_DIR}")
logger.info(f"🔍 BASE_DIR (EXAMS_INTERNAL): {BASE_DIR}")
logger.info(f"🔍 TEMPLATE_DIR: {TEMPLATE_DIR}")
logger.info(f"🔍 STATIC_DIR: {STATIC_DIR}")
logger.info(f"🔍 Template dir exists: {os.path.exists(TEMPLATE_DIR)}")
logger.info(f"🔍 Static dir exists: {os.path.exists(STATIC_DIR)}")

# Verify templates
if os.path.exists(TEMPLATE_DIR):
    templates = os.listdir(TEMPLATE_DIR)
    logger.info(f"🔍 Templates found: {templates}")
    if "login.html" in templates:
        logger.info(f"✅ login.html found in {TEMPLATE_DIR}")
    else:
        logger.warning(f"❌ login.html NOT found in {TEMPLATE_DIR}")
else:
    logger.error(f"❌ Template directory not found: {TEMPLATE_DIR}")

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
        logger.error(f"⚠️ Error checking environment: {e}")
        return True

# Check if BASE_DIR exists
if not os.path.exists(BASE_DIR):
    logger.info(f"📁 Creating BASE_DIR: {BASE_DIR}")
    os.makedirs(BASE_DIR, exist_ok=True)
else:
    logger.info(f"✅ BASE_DIR already exists: {BASE_DIR}")

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
            logger.info(f"📁 Created subdirectory: {dir_path}")
        except Exception as e:
            logger.error(f"⚠️ Could not create {dir_path}: {e}")
    else:
        logger.info(f"✅ Directory already exists: {dir_path}")

# Script mapping
SCRIPT_MAP = {
    "utme": "utme_result.py",
    "caosce": "caosce_result.py",
    "clean": "obj_results.py",
    "split": "split_names.py",
    "exam_processor_nd": "exam_result_processor.py",
    "exam_processor_bn": "exam_processor_bn.py",
    "exam_processor_bm": "exam_processor_bm.py",
}

# Success indicators for detecting processed files
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
        r"Processing: (Set2025-.*?\.xlsx)",
        r"✅ Cleaned CSV saved in.*?cleaned_(Set2025-.*?\.csv)",
        r"🎉 Master CSV saved in.*?master_cleaned_results\.csv",
        r"✅ All processing completed successfully!",
    ],
    "split": [r"Saved processed file: (clean_jamb_DB_.*?\.csv)"],
    "exam_processor_nd": [
        r"PROCESSING SEMESTER: (ND-[A-Za-z0-9\s\-]+)",
        r"✅ Successfully processed .*",
        r"✅ Mastersheet saved:.*",
        r"📁 Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"✅ Processing complete",
        r"✅ ND Examination Results Processing completed successfully",
        r"🔄 Applying upgrade rule:.*→ 50",
        r"✅ Upgraded \d+ scores from.*to 50",
        r"✅ Identified \d+ carryover students",
        r"✅ Carryover records saved:.*",
        r"🔄 Processing resit results for.*",
        r"✅ Updated \d+ scores for \d+ students",
    ],
    "exam_processor_bn": [
        r"PROCESSING SET: (SET47|SET48)",
        r"✅ Successfully processed .*",
        r"✅ Mastersheet saved:.*",
        r"📁 Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"✅ Processing complete",
        r"✅ Basic Nursing Examination Results Processing completed successfully",
        r"🔄 Applying upgrade rule:.*→ 50",
        r"✅ Upgraded \d+ scores from.*to 50",
        r"✅ Identified \d+ carryover students",
        r"✅ Carryover records saved:.*",
        r"🔄 Processing resit results for.*",
        r"✅ Updated \d+ scores for \d+ students",
    ],
    "exam_processor_bm": [
        r"PROCESSING SET: (SET2023|SET2024|SET2025)",
        r"✅ Successfully processed .*",
        r"✅ Mastersheet saved:.*",
        r"📁 Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"✅ Processing complete",
        r"✅ Basic Midwifery Examination Results Processing completed successfully",
        r"🔄 Applying upgrade rule:.*→ 50",
        r"✅ Upgraded \d+ scores from.*to 50",
        r"✅ Identified \d+ carryover students",
        r"✅ Carryover records saved:.*",
        r"🔄 Processing resit results for.*",
        r"✅ Updated \d+ scores for \d+ students",
    ],
}

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv", "zip", "pdf"}

# Helper Functions
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def get_raw_directory(script_name, program=None, set_name=None):
    """Get the RAW_RESULTS directory for a specific script/program/set"""
    logger.info(f"🔍 Getting raw directory for: script={script_name}, program={program}, set={set_name}")
    
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
        if program and set_name:
            raw_dir = os.path.join(BASE_DIR, program, set_name, "RAW_RESULTS")
            logger.info(f"🔍 Exam processor raw directory: {raw_dir}")
            return raw_dir
        return BASE_DIR
    
    raw_paths = {
        "utme": os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_PUTME_RESULT"),
        "caosce": os.path.join(BASE_DIR, "CAOSCE_RESULT", "RAW_CAOSCE_RESULT"),
        "clean": os.path.join(BASE_DIR, "OBJ_RESULT", "RAW_OBJ"),
        "split": os.path.join(BASE_DIR, "JAMB_DB", "RAW_JAMB_DB"),
    }
    raw_dir = raw_paths.get(script_name, BASE_DIR)
    logger.info(f"🔍 Other script raw directory: {raw_dir}")
    return raw_dir

def get_clean_directory(script_name, program=None, set_name=None):
    """Get the CLEAN_RESULTS directory for a specific script/program/set"""
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
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

def get_input_directory(script_name, program=None, set_name=None):
    """Returns the correct input directory for raw results"""
    logger.info(f"🔍 Getting input directory for: {script_name}, program={program}, set={set_name}")
    
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
        if program and set_name and set_name != "all":
            input_dir = os.path.join(BASE_DIR, program, set_name, "RAW_RESULTS")
            logger.info(f"🔍 Exam processor specific input directory: {input_dir}")
            return input_dir
        # If no specific set or "all", return the program directory to scan all sets
        input_dir = os.path.join(BASE_DIR, program) if program else BASE_DIR
        logger.info(f"🔍 Exam processor general input directory: {input_dir}")
        return input_dir
    
    input_dir = get_raw_directory(script_name)
    logger.info(f"🔍 Other script input directory: {input_dir}")
    return input_dir

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "logged_in" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

# FIXED: Improved file detection function with better status checking
def check_exam_processor_files(input_dir, program, selected_set=None):
    """Check if exam processor files exist, optionally filtering by selected set"""
    logger.info(f"🔍 Checking exam processor files in: {input_dir}")
    logger.info(f"🔍 Program: {program}, Selected Set: {selected_set}")
    
    if not os.path.isdir(input_dir):
        logger.error(f"❌ Input directory doesn't exist: {input_dir}")
        return False

    # Check course file - FIXED: More flexible course file detection
    course_dir = os.path.join(BASE_DIR, program, f"{program}-COURSES")
    course_file = None
    course_file_found = False
    
    if os.path.exists(course_dir):
        # Look for any course file with common patterns
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
                    logger.info(f"✅ Found course file: {course_file}")
                    break
            if course_file:
                break
    
    if not course_file_found:
        logger.warning(f"⚠️ Course file not found in: {course_dir}")
        # Don't return False here - let the processor handle missing course files

    program_dir = os.path.join(BASE_DIR, program)
    if not os.path.exists(program_dir):
        logger.error(f"❌ Program directory not found: {program_dir}")
        return False

    program_sets = []
    valid_sets = BN_SETS if program == "BN" else (BM_SETS if program == "BM" else ND_SETS)
    
    # FIXED: Better set filtering logic
    if selected_set and selected_set != "all":
        if selected_set in valid_sets:
            program_sets = [selected_set]
            logger.info(f"🔍 Processing specific set: {selected_set}")
        else:
            logger.error(f"❌ Invalid set selected: {selected_set}")
            return False
    else:
        # Process all valid sets
        for item in os.listdir(program_dir):
            item_path = os.path.join(program_dir, item)
            if os.path.isdir(item_path) and item in valid_sets:
                program_sets.append(item)
        logger.info(f"🔍 Processing all sets: {program_sets}")

    if not program_sets:
        logger.error(f"❌ No {program} sets found in {program_dir} (valid sets: {valid_sets})")
        return False

    total_files_found = 0
    files_found = []
    
    for program_set in program_sets:
        raw_results_path = os.path.join(program_dir, program_set, "RAW_RESULTS")
        logger.info(f"🔍 Checking raw results path: {raw_results_path}")
        
        if not os.path.exists(raw_results_path):
            logger.warning(f"⚠️ RAW_RESULTS not found in {raw_results_path}")
            continue
            
        # FIXED: More accurate file detection - only count valid Excel files
        files = []
        for f in os.listdir(raw_results_path):
            file_path = os.path.join(raw_results_path, f)
            # Check if it's a file (not directory) and has allowed extension
            if (os.path.isfile(file_path) and 
                f.lower().endswith((".xlsx", ".xls")) and  # FIXED: Only Excel files for processing
                not f.startswith("~$") and
                not f.startswith(".")):
                files.append(f)
                files_found.append(f"{program_set}/{f}")
        
        total_files_found += len(files)
        
        if files:
            logger.info(f"📁 Found {len(files)} files in {raw_results_path}: {files}")
        else:
            logger.warning(f"⚠️ No Excel files found in {raw_results_path}")

    logger.info(f"📊 Total Excel files found for {program}: {total_files_found}")
    logger.info(f"📄 Files found: {files_found}")
    
    # FIXED: Return True only if we have actual Excel files to process
    return total_files_found > 0

def check_putme_files(input_dir):
    if not os.path.isdir(input_dir):
        return False
    excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".xlsx", ".xls")) and "PUTME" in f.upper()]
    candidate_batches_dir = os.path.join(os.path.dirname(input_dir), "RAW_CANDIDATE_BATCHES")
    batch_files = [f for f in os.listdir(candidate_batches_dir) if f.lower().endswith(".csv") and "BATCH" in f.upper()] if os.path.isdir(candidate_batches_dir) else []
    return len(excel_files) > 0 and len(batch_files) > 0

def check_internal_exam_files(input_dir):
    if not os.path.isdir(input_dir):
        return False
    excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".xlsx", ".xls")) and f.startswith("Set")]
    return len(excel_files) > 0

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
    logger.info(f"🔍 Checking input files for {script_name} in {input_dir} (selected_set: {selected_set})")
    
    if not os.path.isdir(input_dir):
        logger.error(f"❌ Input directory doesn't exist: {input_dir}")
        return False
        
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
        program = script_name.split("_")[-1].upper()
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

# NEW: Function to get detailed file status for templates
def get_exam_processor_status(program, selected_set=None):
    """Get detailed status for exam processor including file counts"""
    logger.info(f"🔍 Getting status for {program}, set: {selected_set}")
    
    status_info = {
        'ready': False,
        'course_file': False,
        'raw_files_count': 0,
        'raw_files_list': [],
        'sets_ready': {}
    }
    
    # Check course file
    course_dir = os.path.join(BASE_DIR, program, f"{program}-COURSES")
    if os.path.exists(course_dir):
        for file in os.listdir(course_dir):
            if file.lower().endswith('.xlsx') and any(keyword in file.lower() for keyword in ['course', 'credit']):
                status_info['course_file'] = True
                break
    
    program_dir = os.path.join(BASE_DIR, program)
    valid_sets = BN_SETS if program == "BN" else (BM_SETS if program == "BM" else ND_SETS)
    
    # Determine which sets to check
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
    
    # Overall ready status: must have course file AND at least one raw file
    status_info['ready'] = status_info['course_file'] and total_files > 0
    
    logger.info(f"📊 Status for {program}: course_file={status_info['course_file']}, raw_files={total_files}, ready={status_info['ready']}")
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
                if script_name in ["exam_processor_bn", "exam_processor_bm", "exam_processor_nd"]:
                    if "PROCESSING SEMESTER:" in line.upper() or "PROCESSING SET:" in line.upper():
                        set_name = match.group(1)
                        if set_name:
                            processed_files_set.add(f"Set: {set_name}")
                            logger.info(f"🔍 DETECTED SET: {set_name}")
                    elif "Mastersheet saved:" in line:
                        file_name = match.group(0).split("Mastersheet saved:")[-1].strip()
                        processed_files_set.add(f"Saved: {file_name}")
                    elif "✅ Identified" in line and "carryover students" in line:
                        processed_files_set.add("Carryover students identified")
                    elif "✅ Carryover records saved:" in line:
                        processed_files_set.add("Carryover records saved")
                    elif "🔄 Processing resit results for" in line:
                        processed_files_set.add("Resit processing started")
                    elif "✅ Updated" in line and "scores for" in line and "students" in line:
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
                    elif "✅ Cleaned CSV saved" in line:
                        file_name = match.group(1) if match.groups() else "cleaned_file"
                        processed_files_set.add(f"Cleaned: {file_name}")
                    elif "🎉 Master CSV saved" in line:
                        processed_files_set.add("Master file created")
                    elif "✅ All processing completed successfully!" in line:
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
        if any("✅ All processing completed successfully!" in line for line in output_lines):
            return f"Successfully processed internal examination results! Generated master file and individual cleaned files."
        return f"Processed {processed_files} internal examination file(s)."
    elif script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
        program = script_name.split("_")[-1].upper()
        program_name = {"ND": "ND", "BN": "Basic Nursing", "BM": "Basic Midwifery"}.get(program, program)
        upgrade_info = ""
        upgrade_count = ""
        carryover_info = ""
        resit_info = ""
        
        for line in output_lines:
            if "🔄 Applying upgrade rule:" in line:
                upgrade_match = re.search(r"🔄 Applying upgrade rule: (\d+)–49 → 50", line)
                if upgrade_match:
                    upgrade_info = f" Upgrade rule applied: {upgrade_match.group(1)}-49 → 50"
                    break
            elif "✅ Upgraded" in line:
                upgrade_count_match = re.search(r"✅ Upgraded (\d+) scores", line)
                if upgrade_count_match:
                    upgrade_count = f" Upgraded {upgrade_count_match.group(1)} scores"
                    break
            elif "✅ Identified" in line and "carryover students" in line:
                carryover_match = re.search(r"✅ Identified (\d+) carryover students", line)
                if carryover_match:
                    carryover_info = f" Identified {carryover_match.group(1)} carryover students"
                    break
            elif "🔄 Processing resit results for" in line:
                resit_info = " Resit processing completed"
                break
        
        if any(f"✅ {program_name} Examination Results Processing completed successfully" in line for line in output_lines):
            return f"{program_name} Examination processing completed successfully! Processed {processed_files} set(s).{upgrade_info}{upgrade_count}{carryover_info}{resit_info}"
        elif any("✅ Processing complete" in line for line in output_lines):
            return f"{program_name} Examination processing completed! Processed {processed_files} set(s).{upgrade_info}{upgrade_count}{carryover_info}{resit_info}"
        return f"Processed {processed_files} {program_name} examination set(s).{upgrade_info}{upgrade_count}{carryover_info}{resit_info}"
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
    logger.info(f"🔍 Script path for {script_name}: {script_path}")
    if not os.path.isfile(script_path):
        logger.error(f"Script not found: {script_path}")
        raise FileNotFoundError(f"Script {script_name} not found at {script_path}")
    return script_path

def get_files_by_category():
    """Get files organized by category and semester/set from existing RAW and CLEAN folders"""
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

    # Search in program-specific CLEAN_RESULTS folders with semester segmentation
    for program in ["ND", "BN", "BM"]:
        program_dir = os.path.join(BASE_DIR, program)
        if not os.path.exists(program_dir):
            continue
        
        sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
        for set_name in sets:
            clean_dir = os.path.join(program_dir, set_name, "CLEAN_RESULTS")
            if not os.path.exists(clean_dir):
                continue
            
            # Initialize set in category
            category = f"{program.lower()}_results"
            if set_name not in files_by_category[category]:
                files_by_category[category][set_name] = []
            
            for root, _, files in os.walk(clean_dir):
                for file in files:
                    if not file.lower().endswith((".xlsx", ".csv", ".pdf", ".zip")):
                        continue
                    file_path = os.path.join(root, file)
                    try:
                        relative_path = os.path.relpath(file_path, BASE_DIR)
                        folder = os.path.basename(os.path.dirname(file_path))
                        
                        # Detect semester from filename
                        semester = "Unknown"
                        if "first" in file.lower() and "semester" in file.lower():
                            semester = "First Semester"
                        elif "second" in file.lower() and "semester" in file.lower():
                            semester = "Second Semester"
                        elif "FIRST-YEAR-FIRST-SEMESTER" in file.upper():
                            semester = "First Year - First Semester"
                        elif "FIRST-YEAR-SECOND-SEMESTER" in file.upper():
                            semester = "First Year - Second Semester"
                        elif "SECOND-YEAR-FIRST-SEMESTER" in file.upper():
                            semester = "Second Year - First Semester"
                        elif "SECOND-YEAR-SECOND-SEMESTER" in file.upper():
                            semester = "Second Year - Second Semester"
                        
                        file_info = FileInfo(
                            name=file,
                            relative_path=relative_path,
                            folder=folder,
                            size=os.path.getsize(file_path),
                            modified=os.path.getmtime(file_path),
                            semester=semester,
                            set_name=set_name
                        )
                        files_by_category[category][set_name].append(file_info)
                        logger.info(f"✅ Categorized as {category}/{set_name}: {file}")
                    except Exception as e:
                        logger.error(f"⚠️ Error processing file {file}: {e}")

    # Search in other result CLEAN folders
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
        
        for root, _, files in os.walk(clean_dir):
            for file in files:
                if not file.lower().endswith((".xlsx", ".csv", ".pdf", ".zip")):
                    continue
                file_path = os.path.join(root, file)
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
                    logger.info(f"✅ Categorized as {category}: {file}")
                except Exception as e:
                    logger.error(f"⚠️ Error processing file {file}: {e}")

    for category, files in files_by_category.items():
        if isinstance(files, dict):
            total_files = sum(len(files_in_set) for files_in_set in files.values())
            logger.info(f"📊 {category}: {total_files} files across {len(files)} sets")
        else:
            logger.info(f"📊 {category}: {len(files)} files")
    
    return files_by_category

def get_sets_and_folders():
    """Get sets and their result folders from CLEAN_RESULTS directories"""
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
    logger.info(f"🔍 Scanning BASE_DIR: {BASE_DIR}")

    # Scan program-specific sets
    for program in ["ND", "BN", "BM"]:
        program_dir = os.path.join(BASE_DIR, program)
        if not os.path.exists(program_dir):
            logger.warning(f"⚠️ Program directory not found: {program_dir}")
            continue
        logger.info(f"🔍 Scanning program: {program_dir}")
        valid_sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
        for set_name in os.listdir(program_dir):
            set_path = os.path.join(program_dir, set_name)
            if not os.path.isdir(set_path) or set_name not in valid_sets:
                logger.warning(f"⚠️ Skipping invalid set {set_name} for program {program}")
                continue
            logger.info(f"🔍 Scanning set: {set_path}")
            folders = []
            clean_results_path = os.path.join(set_path, "CLEAN_RESULTS")
            if os.path.exists(clean_results_path):
                logger.info(f"🔍 Found CLEAN_RESULTS: {clean_results_path}")
                try:
                    # Check for subdirectories in CLEAN_RESULTS
                    for item in os.listdir(clean_results_path):
                        item_path = os.path.join(clean_results_path, item)
                        if os.path.isdir(item_path):
                            # This is a folder containing files
                            files = []
                            for file in os.listdir(item_path):
                                if file.lower().endswith((".xlsx", ".csv", ".pdf", ".zip")):
                                    file_path = os.path.join(item_path, file)
                                    try:
                                        relative_path = os.path.relpath(file_path, BASE_DIR)
                                        files.append(FileInfo(
                                            name=file,
                                            relative_path=relative_path,
                                            size=os.path.getsize(file_path),
                                            modified=os.path.getmtime(file_path)
                                        ))
                                        logger.info(f"✅ Found file: {file} in folder {item}")
                                    except Exception as e:
                                        logger.error(f"⚠️ Error processing file {file}: {e}")
                            if files:
                                folders.append(FolderInfo(name=item, files=files))
                                logger.info(f"✅ Added folder {item} with {len(files)} files")
                except Exception as e:
                    logger.error(f"⚠️ Error scanning {clean_results_path}: {e}")
            if folders:
                set_key = f"{program}_{set_name}"
                sets[set_key] = folders
                logger.info(f"✅ Added set {set_key} with {len(folders)} folders")

    # Scan other result types - FIXED MAPPINGS FOR INTERNAL_RESULT
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
        logger.info(f"🔍 Scanning {result_type} at: {result_path}")
        
        if not os.path.exists(result_path):
            logger.warning(f"⚠️ Directory not found: {result_path}")
            continue
            
        folders = []
        try:
            # First, check if there are direct files in the clean directory
            direct_files = []
            for file in os.listdir(result_path):
                file_path = os.path.join(result_path, file)
                if os.path.isfile(file_path) and file.lower().endswith((".xlsx", ".csv", ".pdf", ".zip")):
                    try:
                        relative_path = os.path.relpath(file_path, BASE_DIR)
                        direct_files.append(FileInfo(
                            name=file,
                            relative_path=relative_path,
                            size=os.path.getsize(file_path),
                            modified=os.path.getmtime(file_path)
                        ))
                        logger.info(f"✅ Found direct file: {file} in {result_path}")
                    except Exception as e:
                        logger.error(f"⚠️ Error processing file {file}: {e}")
            
            if direct_files:
                folders.append(FolderInfo(name="Results", files=direct_files))
                logger.info(f"✅ Added direct files folder with {len(direct_files)} files")
            
            # Then check for subdirectories
            for folder in os.listdir(result_path):
                folder_path = os.path.join(result_path, folder)
                if not os.path.isdir(folder_path):
                    continue
                    
                files = []
                for file in os.listdir(folder_path):
                    if not file.lower().endswith((".xlsx", ".csv", ".pdf", ".zip")):
                        continue
                    file_path = os.path.join(folder_path, file)
                    try:
                        relative_path = os.path.relpath(file_path, BASE_DIR)
                        files.append(FileInfo(
                            name=file,
                            relative_path=relative_path,
                            size=os.path.getsize(file_path),
                            modified=os.path.getmtime(file_path)
                        ))
                        logger.info(f"✅ Found file: {file} in {folder_path}")
                    except Exception as e:
                        logger.error(f"⚠️ Error processing file {file}: {e}")
                if files:
                    folders.append(FolderInfo(name=folder, files=files))
                    logger.info(f"✅ Added folder {folder} with {len(files)} files")
                    
        except Exception as e:
            logger.error(f"⚠️ Error scanning {result_path}: {e}")
            
        if folders:
            sets[result_type] = folders
            logger.info(f"✅ Added result type {result_type} with {len(folders)} folders")
        else:
            logger.info(f"ℹ️ No folders found for {result_type}")

    logger.info(f"📊 Total sets found: {len(sets)}")
    for set_name, folders in sets.items():
        total_files = sum(len(folder.files) for folder in folders)
        logger.info(f"📁 {set_name}: {total_files} files in {len(folders)} folders")
        
    return sets

# NEW: Carryover and Resit Management Functions
def get_carryover_records(program, set_name, semester_key=None):
    """Get carryover records for a specific program, set, and semester"""
    try:
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            return []
        
        # Look for the most recent timestamped folder
        timestamp_folders = [f for f in os.listdir(clean_dir) 
                           if f.startswith(f"{set_name}_RESULT-") and os.path.isdir(os.path.join(clean_dir, f))]
        
        if not timestamp_folders:
            return []
        
        latest_folder = sorted(timestamp_folders)[-1]
        output_dir = os.path.join(clean_dir, latest_folder)
        carryover_dir = os.path.join(output_dir, "CARRYOVER_RECORDS")
        
        if not os.path.exists(carryover_dir):
            return []
        
        carryover_files = []
        for file in os.listdir(carryover_dir):
            if file.startswith("co_student_") and file.endswith(".json"):
                # Extract semester from filename
                file_semester = None
                if "FIRST-YEAR-FIRST-SEMESTER" in file:
                    file_semester = "ND-FIRST-YEAR-FIRST-SEMESTER"
                elif "FIRST-YEAR-SECOND-SEMESTER" in file:
                    file_semester = "ND-FIRST-YEAR-SECOND-SEMESTER"
                elif "SECOND-YEAR-FIRST-SEMESTER" in file:
                    file_semester = "ND-SECOND-YEAR-FIRST-SEMESTER"
                elif "SECOND-YEAR-SECOND-SEMESTER" in file:
                    file_semester = "ND-SECOND-YEAR-SECOND-SEMESTER"
                
                # Filter by semester if specified
                if semester_key and file_semester != semester_key:
                    continue
                    
                file_path = os.path.join(carryover_dir, file)
                try:
                    with open(file_path, 'r') as f:
                        data = json.load(f)
                        carryover_files.append({
                            'filename': file,
                            'semester': file_semester,
                            'data': data,
                            'count': len(data),
                            'file_path': file_path
                        })
                except Exception as e:
                    logger.error(f"Error loading carryover file {file}: {e}")
        
        return carryover_files
    except Exception as e:
        logger.error(f"Error getting carryover records: {e}")
        return []

def get_carryover_summary(program, set_name):
    """Get summary of carryover students"""
    carryover_records = get_carryover_records(program, set_name)
    summary = {
        'total_students': 0,
        'total_courses': 0,
        'by_semester': {},
        'recent_semester': None,
        'recent_count': 0
    }
    
    for record in carryover_records:
        semester = record['semester']
        if semester not in summary['by_semester']:
            summary['by_semester'][semester] = 0
        
        summary['by_semester'][semester] += record['count']
        summary['total_students'] += record['count']
        summary['total_courses'] += sum(len(student['failed_courses']) for student in record['data'])
    
    # Find the most recent semester
    if summary['by_semester']:
        summary['recent_semester'] = max(summary['by_semester'].keys(), 
                                       key=lambda x: summary['by_semester'][x])
        summary['recent_count'] = summary['by_semester'][summary['recent_semester']]
    
    return summary

# Routes
@app.route("/", methods=["GET"])
def index():
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    try:
        app.logger.info(f"Login route accessed - Method: {request.method}")
        logger.info(f"🔍 Attempting to load template: {os.path.join(TEMPLATE_DIR, 'login.html')}")
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
        logger.error(f"❌ Template error: {e} at {TEMPLATE_DIR}")
        return f"<h1>Template Error</h1><p>Template not found: login.html at {TEMPLATE_DIR}</p><p>Available templates: {os.listdir(TEMPLATE_DIR) if os.path.exists(TEMPLATE_DIR) else 'None'}</p>", 500
    except Exception as e:
        app.logger.error(f"Login error: {e}")
        logger.error(f"❌ Login error: {e}")
        return f"<h1>Server Error</h1><p>{str(e)}</p>", 500

@app.route("/dashboard")
@login_required
def dashboard():
    try:
        # Get carryover summaries for each program
        carryover_summaries = {}
        for program in ["ND", "BN", "BM"]:
            program_sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
            for set_name in program_sets:
                summary = get_carryover_summary(program, set_name)
                if summary['total_students'] > 0:
                    key = f"{program}_{set_name}"
                    carryover_summaries[key] = summary
        
        return render_template(
            "dashboard.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            nd_sets=ND_SETS,
            bn_sets=BN_SETS,
            bm_sets=BM_SETS,
            programs=PROGRAMS,
            carryover_summaries=carryover_summaries
        )
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
    except TemplateNotFound:
        flash("Debug paths template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"Debug paths error: {e}")
        flash(f"Error loading debug page: {str(e)}", "error")
        return redirect(url_for("dashboard"))

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
    except Exception as e:
        app.logger.error(f"Upload center error: {e}")
        flash(f"Error loading upload center: {str(e)}", "error")
        return redirect(url_for("dashboard"))

# NEW ROUTE: Handle uploads from the upload center
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
        
        # Map program to script_name for directory structure
        program_map = {
            "nd": ("exam_processor_nd", "ND"),
            "bn": ("exam_processor_bn", "BN"), 
            "bm": ("exam_processor_bm", "BM")
        }
        
        if program not in program_map:
            flash("Invalid program selected.", "error")
            return redirect(url_for("upload_center"))
            
        script_name, program_code = program_map[program]
        
        # Get the set name based on program
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
        
        # Get the target directory
        raw_dir = get_raw_directory(script_name, program_code, set_name)
        logger.info(f"🔍 Uploading to directory: {raw_dir}")
        
        # Create directory if it doesn't exist
        os.makedirs(raw_dir, exist_ok=True)
        
        saved_files = []
        skipped_files = []
        
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(raw_dir, filename)
                
                # Save the file
                file.save(file_path)
                saved_files.append(filename)
                logger.info(f"✅ Saved file: {file_path}")
                
                # Handle ZIP files - extract them
                if filename.lower().endswith(".zip"):
                    try:
                        with zipfile.ZipFile(file_path, "r") as zip_ref:
                            zip_ref.extractall(raw_dir)
                        # Optionally remove the zip file after extraction
                        # os.remove(file_path)
                        logger.info(f"📦 Extracted ZIP file: {filename}")
                        saved_files.append(f"{filename} (extracted)")
                    except zipfile.BadZipFile:
                        logger.error(f"❌ Invalid ZIP file: {filename}")
                        skipped_files.append(f"{filename} (invalid ZIP)")
                    except Exception as e:
                        logger.error(f"❌ Error extracting ZIP {filename}: {e}")
                        skipped_files.append(f"{filename} (extraction error)")
            else:
                skipped_files.append(file.filename if file.filename else "unknown file")
        
        # Prepare success message
        if saved_files:
            success_msg = f"✅ Successfully uploaded {len(saved_files)} file(s) to {program_code}/{set_name}/RAW_RESULTS"
            if skipped_files:
                success_msg += f" (Skipped {len(skipped_files)} invalid files)"
            flash(success_msg, "success")
            logger.info(f"✅ Upload completed: {success_msg}")
        else:
            flash("No valid files were uploaded. Please check file formats.", "error")
        
        if skipped_files:
            flash(f"Skipped files: {', '.join(skipped_files)}", "warning")
            
        return redirect(url_for("upload_center"))
        
    except Exception as e:
        app.logger.error(f"Upload error: {e}")
        flash(f"Upload failed: {str(e)}", "error")
        return redirect(url_for("upload_center"))

# NEW ROUTE: Handle resit file uploads
@app.route("/handle_resit_upload", methods=["POST"])
@login_required
def handle_resit_upload():
    """Handle resit file uploads for carryover processing"""
    try:
        program = request.form.get("resit_program")
        set_name = request.form.get("resit_set")
        semester_key = request.form.get("resit_semester")
        resit_files = request.files.getlist("resit_files")
        
        if not program or not set_name or not semester_key:
            flash("Please select program, set, and semester for resit processing.", "error")
            return redirect(url_for("upload_center"))
            
        if not resit_files or all(file.filename == '' for file in resit_files):
            flash("Please select at least one resit file to upload.", "error")
            return redirect(url_for("upload_center"))
        
        # Create resit directory structure
        resit_dir = os.path.join(BASE_DIR, program, set_name, "RESIT_RESULTS")
        os.makedirs(resit_dir, exist_ok=True)
        
        # Create semester-specific subdirectory
        semester_dir = os.path.join(resit_dir, semester_key)
        os.makedirs(semester_dir, exist_ok=True)
        
        saved_files = []
        
        for file in resit_files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(semester_dir, filename)
                
                # Save the file
                file.save(file_path)
                saved_files.append(filename)
                logger.info(f"✅ Saved resit file: {file_path}")
        
        if saved_files:
            flash(f"✅ Successfully uploaded {len(saved_files)} resit file(s) for {program}/{set_name}/{semester_key}", "success")
            logger.info(f"✅ Resit upload completed for {program}/{set_name}/{semester_key}")
        else:
            flash("No valid resit files were uploaded.", "error")
            
        return redirect(url_for("upload_center"))
        
    except Exception as e:
        app.logger.error(f"Resit upload error: {e}")
        flash(f"Resit upload failed: {str(e)}", "error")
        return redirect(url_for("upload_center"))

@app.route("/download_center")
@login_required
def download_center():
    try:
        files_by_category = get_files_by_category()
        for category, files in files_by_category.items():
            if isinstance(files, dict):
                app.logger.info(f"Download center - {category}: {sum(len(f) for f in files.values())} files across {len(files)} sets")
            else:
                app.logger.info(f"Download center - {category}: {len(files)} files")
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
    except Exception as e:
        app.logger.error(f"Download center error: {e}")
        flash(f"Error loading download center: {str(e)}", "error")
        return redirect(url_for("dashboard"))

@app.route("/file_browser")
@login_required
def file_browser():
    try:
        sets = get_sets_and_folders()
        
        # Extract set names for each program
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
        
        # Debug: Log the structure of sets
        for set_key, folders in sets.items():
            total_files = sum(len(folder.files) for folder in folders)
            app.logger.info(f"Set '{set_key}': {total_files} files across {len(folders)} folders")
        
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
            processed_dir=BASE_DIR  # Add this for debug info
        )
    except Exception as e:
        app.logger.error(f"File browser error: {e}")
        flash(f"Error loading file browser: {str(e)}", "error")
        return redirect(url_for("dashboard"))

# NEW ROUTE: Carryover Management Dashboard
@app.route("/carryover_management")
@login_required
def carryover_management():
    """Carryover student management dashboard"""
    try:
        carryover_data = {}
        
        # Get carryover records for all programs and sets
        for program in ["ND", "BN", "BM"]:
            program_sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
            carryover_data[program] = {}
            
            for set_name in program_sets:
                records = get_carryover_records(program, set_name)
                if records:
                    carryover_data[program][set_name] = {
                        'records': records,
                        'total_students': sum(record['count'] for record in records),
                        'total_semesters': len(records)
                    }
        
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
    except Exception as e:
        app.logger.error(f"Carryover management error: {e}")
        flash(f"Error loading carryover management: {str(e)}", "error")
        return redirect(url_for("dashboard"))

# NEW ROUTE: Process Resit Results
@app.route("/process_resit", methods=["POST"])
@login_required
def process_resit():
    """Process resit results for carryover students"""
    try:
        program = request.form.get("resit_program")
        set_name = request.form.get("resit_set")
        semester_key = request.form.get("resit_semester")
        resit_file_path = request.form.get("resit_file_path")
        
        if not program or not set_name or not semester_key:
            flash("Please select program, set, and semester for resit processing.", "error")
            return redirect(url_for("carryover_management"))
        
        # Find the resit file
        resit_dir = os.path.join(BASE_DIR, program, set_name, "RESIT_RESULTS", semester_key)
        if not os.path.exists(resit_dir):
            flash(f"Resit directory not found: {resit_dir}", "error")
            return redirect(url_for("carryover_management"))
        
        resit_files = [f for f in os.listdir(resit_dir) if f.lower().endswith(('.xlsx', '.xls'))]
        if not resit_files:
            flash(f"No resit files found in {resit_dir}", "error")
            return redirect(url_for("carryover_management"))
        
        # Use the first resit file found
        resit_file = resit_files[0]
        resit_file_path = os.path.join(resit_dir, resit_file)
        
        # Find the clean results directory
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            flash(f"Clean results directory not found: {clean_dir}", "error")
            return redirect(url_for("carryover_management"))
        
        # Find the most recent timestamped folder
        timestamp_folders = [f for f in os.listdir(clean_dir) 
                           if f.startswith(f"{set_name}_RESULT-") and os.path.isdir(os.path.join(clean_dir, f))]
        
        if not timestamp_folders:
            flash(f"No result folders found in {clean_dir}", "error")
            return redirect(url_for("carryover_management"))
        
        latest_folder = sorted(timestamp_folders)[-1]
        output_dir = os.path.join(clean_dir, latest_folder)
        
        # Set up environment for resit processing
        env = os.environ.copy()
        env["BASE_DIR"] = BASE_DIR
        env["PROCESS_RESIT"] = "true"
        env["RESIT_FILE_PATH"] = resit_file_path
        env["SELECTED_SET"] = set_name
        env["SELECTED_SEMESTERS"] = semester_key
        env["PASS_THRESHOLD"] = "50.0"  # Default pass threshold
        
        # Run the exam processor script for resit processing
        script_name = f"exam_processor_{program.lower()}"
        script_path = _get_script_path(script_name)
        
        logger.info(f"🔄 Processing resit results for {program}/{set_name}/{semester_key}")
        logger.info(f"📁 Resit file: {resit_file_path}")
        logger.info(f"📁 Output directory: {output_dir}")
        
        result = subprocess.run(
            [sys.executable, script_path],
            env=env,
            text=True,
            capture_output=True,
            timeout=600,
        )
        
        logger.info("=== RESIT PROCESSING STDOUT ===")
        for line in result.stdout.splitlines():
            logger.info(line)
        logger.info("=== RESIT PROCESSING STDERR ===")
        for line in result.stderr.splitlines():
            logger.info(line)
            
        output_lines = result.stdout.splitlines()
        
        if result.returncode == 0:
            if any("✅ Updated" in line and "scores for" in line and "students" in line for line in output_lines):
                flash(f"✅ Resit processing completed successfully for {program}/{set_name}/{semester_key}", "success")
            else:
                flash(f"ℹ️ Resit processing completed but no scores were updated for {program}/{set_name}/{semester_key}", "info")
        else:
            flash(f"❌ Resit processing failed for {program}/{set_name}/{semester_key}: {result.stderr}", "error")
        
        return redirect(url_for("carryover_management"))
        
    except Exception as e:
        app.logger.error(f"Resit processing error: {e}")
        flash(f"Resit processing failed: {str(e)}", "error")
        return redirect(url_for("carryover_management"))

# FIXED: Enhanced run_script function with better file detection and status reporting
@app.route("/run/<script_name>", methods=["GET", "POST"])
@login_required
def run_script(script_name):
    try:
        if script_name not in SCRIPT_MAP:
            flash("Invalid script requested.", "error")
            return redirect(url_for("dashboard"))
            
        program = script_name.split("_")[-1].upper() if script_name.startswith("exam_processor") else None
        script_desc = {
            "utme": "PUTME Examination Results",
            "caosce": "CAOSCE Examination Results",
            "clean": "Internal Examination Results",
            "split": "JAMB Candidate Name Split",
            "exam_processor_nd": "ND Examination Results Processing",
            "exam_processor_bn": "Basic Nursing Examination Results Processing",
            "exam_processor_bm": "Basic Midwifery Examination Results Processing",
        }.get(script_name, "Script")
        
        script_path = _get_script_path(script_name)
        
        # Get selected set for exam processors
        selected_set = None
        if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
            selected_set = request.form.get(
                "selected_set" if script_name == "exam_processor_nd" else
                "nursing_set" if script_name == "exam_processor_bn" else
                "midwifery_set", "all"
            )
            logger.info(f"🔍 Selected set for {script_name}: {selected_set}")
        
        # Get the correct input directory based on script and selected set
        input_dir = get_input_directory(script_name, program, selected_set)
        logger.info(f"🔍 Final input directory for {script_name}: {input_dir}")
        
        # FIXED: Better file checking with detailed status reporting
        if not check_input_files(input_dir, script_name, selected_set):
            # For exam processors, provide more specific guidance with file listing
            if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
                status_info = get_exam_processor_status(program, selected_set)
                
                error_msg = f"❌ Cannot process {script_desc}: "
                issues = []
                
                if not status_info['course_file']:
                    issues.append("course file not found")
                
                if status_info['raw_files_count'] == 0:
                    issues.append("no Excel files found in RAW_RESULTS folders")
                
                if issues:
                    error_msg += " and ".join(issues)
                    
                    # Add specific file details
                    if selected_set and selected_set != "all":
                        expected_path = os.path.join(BASE_DIR, program, selected_set, "RAW_RESULTS")
                        error_msg += f". Please upload Excel files to: {expected_path}"
                    else:
                        error_msg += f". Please upload Excel files to the appropriate set folders in: {os.path.join(BASE_DIR, program)}"
                    
                    flash(error_msg, "error")
                else:
                    flash(f"❌ Unexpected error: No files found for {script_desc}", "error")
            else:
                flash(f"❌ No valid input files found in {input_dir} for {script_desc}.", "error")
            return redirect(url_for("dashboard"))
            
        if request.method == "GET":
            template_map = {
                "utme": "utme_form.html",
                "exam_processor_nd": "exam_processor_form.html",
                "exam_processor_bn": "basic_nursing_form.html",
                "exam_processor_bm": "basic_midwifery_form.html",
            }
            template = template_map.get(script_name)
            if template:
                # FIXED: Get detailed status information for templates
                status_info = {}
                if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
                    status_info = get_exam_processor_status(program)
                
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
                    status_info=status_info  # FIXED: Pass status info to template
                )
                
        if request.method == "POST":
            if script_name == "utme":
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
                )
                output_lines = result.stdout.splitlines()
                processed_files = count_processed_files(output_lines, script_name)
                success_msg = get_success_message(script_name, processed_files, output_lines)
                if success_msg:
                    flash(success_msg, "success")
                else:
                    flash(f"No files processed for {script_desc}. Check input files in {input_dir}.", "error")
                return redirect(url_for("dashboard"))
                
            elif script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
                selected_set = request.form.get(
                    "selected_set" if script_name == "exam_processor_nd" else
                    "nursing_set" if script_name == "exam_processor_bn" else
                    "midwifery_set", "all"
                )
                pass_threshold = request.form.get("pass_threshold", "50.0")
                upgrade_threshold = request.form.get("upgrade_threshold", "0")
                generate_pdf = "generate_pdf" in request.form
                track_withdrawn = "track_withdrawn" in request.form
                
                # Get selected semesters for ND processor
                selected_semesters = request.form.getlist('selected_semesters')
                if not selected_semesters or 'all' in selected_semesters:
                    selected_semesters = ['all']
                
                env = os.environ.copy()
                env["BASE_DIR"] = BASE_DIR
                env["SELECTED_SET"] = selected_set
                env["SELECTED_SEMESTERS"] = ','.join(selected_semesters)
                env["PASS_THRESHOLD"] = pass_threshold
                if upgrade_threshold and upgrade_threshold.strip() and upgrade_threshold != "0":
                    env["UPGRADE_THRESHOLD"] = upgrade_threshold.strip()
                env["GENERATE_PDF"] = str(generate_pdf)
                env["TRACK_WITHDRAWN"] = str(track_withdrawn)
                
                logger.info(f"🚀 Running {script_name} with environment:")
                logger.info(f"  BASE_DIR: {env['BASE_DIR']}")
                logger.info(f"  SELECTED_SET: {env['SELECTED_SET']}")
                logger.info(f"  SELECTED_SEMESTERS: {env['SELECTED_SEMESTERS']}")
                logger.info(f"  PASS_THRESHOLD: {env['PASS_THRESHOLD']}")
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
                
                if result.returncode == 0:
                    if success_msg:
                        flash(success_msg, "success")
                    else:
                        flash(f"{script_desc} completed but no files were processed.", "warning")
                    
                    # FIXED: Enhanced ZIP creation with better error handling and file verification
                    clean_dir = get_clean_directory(script_name, program, selected_set)
                    if os.path.exists(clean_dir):
                        # Look for timestamped result folders
                        result_dirs = [d for d in os.listdir(clean_dir) 
                                     if d.startswith(f"{selected_set}_RESULT-") and os.path.isdir(os.path.join(clean_dir, d))]
                        
                        if result_dirs:
                            latest = sorted(result_dirs)[-1]  # Get the most recent
                            folder_path = os.path.join(clean_dir, latest)
                            
                            # Verify the folder has actual files before creating ZIP
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
                                            # Use relative path within the ZIP
                                            arcname = os.path.relpath(file_path, folder_path)
                                            zip_f.write(file_path, arcname)
                                    
                                    # Verify the ZIP was created successfully
                                    if os.path.exists(zip_path) and os.path.getsize(zip_path) > 100:  # At least 100 bytes
                                        flash(f"📦 Zipped results ready: {zip_filename} ({len(all_files)} files)", "success")
                                        logger.info(f"✅ Successfully created ZIP: {zip_path} with {len(all_files)} files")
                                    else:
                                        logger.warning(f"⚠️ ZIP file created but appears empty: {zip_path}")
                                        flash(f"⚠️ ZIP created but appears empty for {selected_set}", "warning")
                                        
                                except Exception as e:
                                    logger.error(f"❌ Failed to create ZIP: {e}")
                                    flash(f"Failed to create ZIP for {selected_set}: {str(e)}", "error")
                            else:
                                logger.warning(f"⚠️ No files found to zip in: {folder_path}")
                                flash(f"No output files generated for {selected_set}", "warning")
                else:
                    flash(f"Script failed: {result.stderr or 'Unknown error'}", "error")
                return redirect(url_for("dashboard"))
                
        # Default processing for other scripts
        result = subprocess.run(
            [sys.executable, script_path],
            text=True,
            capture_output=True,
            check=True,
            timeout=300,
        )
        output_lines = result.stdout.splitlines()
        processed_files = count_processed_files(output_lines, script_name)
        success_msg = get_success_message(script_name, processed_files, output_lines)
        if success_msg:
            flash(success_msg, "success")
        else:
            flash(f"No files processed for {script_desc}. Check input files in {input_dir}.", "error")
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
        if script_name not in ["utme", "caosce", "clean", "split", "exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
            flash("Invalid script requested.", "error")
            return redirect(url_for("upload_center"))
            
        program = script_name.split("_")[-1].upper() if script_name.startswith("exam_processor") else None
        script_desc = {
            "utme": "PUTME Results",
            "caosce": "CAOSCE Results",
            "clean": "Internal Examinations",
            "split": "JAMB Database",
            "exam_processor_nd": "ND Examinations",
            "exam_processor_bn": "Basic Nursing",
            "exam_processor_bm": "Basic Midwifery",
        }.get(script_name, "Files")
        
        files = request.files.getlist("files")
        candidate_files = request.files.getlist("candidate_files") if script_name == "utme" else []
        course_files = request.files.getlist("course_files") if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] else []
        set_name = request.form.get("nd_set") or request.form.get("nursing_set") or request.form.get("midwifery_set")
        
        # Save to appropriate RAW directory
        if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] and set_name:
            raw_dir = get_raw_directory(script_name, program, set_name)
            logger.info(f"🔍 Uploading to exam processor raw directory: {raw_dir}")
        else:
            raw_dir = get_raw_directory(script_name)
            logger.info(f"🔍 Uploading to other script raw directory: {raw_dir}")
        
        os.makedirs(raw_dir, exist_ok=True)
        saved_files = []
        
        # Handle course files separately for exam processors
        if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] and course_files:
            course_dir = os.path.join(BASE_DIR, program, f"{program}-COURSES")
            os.makedirs(course_dir, exist_ok=True)
            for file in course_files:
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(course_dir, filename)
                    file.save(file_path)
                    saved_files.append(f"course: {filename}")
                    logger.info(f"✅ Saved course file: {file_path}")
        
        # Handle regular files
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
                        flash(f"📦 Extracted ZIP: {filename}", "success")
                    except zipfile.BadZipFile:
                        flash(f"❌ Invalid ZIP file: {filename}", "error")
                        return redirect(url_for("upload_center"))
                        
        if saved_files:
            flash(f"✅ Uploaded files to {raw_dir}: {', '.join(saved_files)}", "success")
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
        # Search in BASE_DIR for the file
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
    try:
        safe_name = os.path.basename(filename)
        # Search in BASE_DIR for the file
        for root, _, files in os.walk(BASE_DIR):
            if safe_name in files:
                return send_from_directory(root, safe_name, as_attachment=True)
        flash(f"File '{safe_name}' not found.", "error")
        return redirect(url_for("download_center"))
    except Exception as e:
        app.logger.error(f"Download error: {e}")
        flash(f"Download failed: {str(e)}", "error")
        return redirect(url_for("download_center"))

@app.route("/download_zip/<set_name>")
@login_required
def download_zip(set_name):
    try:
        # Find the set's CLEAN_RESULTS directory
        zip_path = None
        for program in ["ND", "BN", "BM"]:
            sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
            if set_name in sets:
                clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
                if os.path.exists(clean_dir):
                    # Look for ZIP files in the clean directory
                    zip_files = [f for f in os.listdir(clean_dir) 
                               if f.startswith(f"{set_name}_RESULT-") and f.endswith('.zip')]
                    
                    if zip_files:
                        # Get the most recent ZIP file
                        latest_zip = sorted(zip_files)[-1]
                        zip_path = os.path.join(clean_dir, latest_zip)
                        break
        
        if zip_path and os.path.exists(zip_path):
            # Verify the ZIP file has reasonable size
            file_size = os.path.getsize(zip_path)
            if file_size > 100:  # At least 100 bytes
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

@app.route("/delete/<path:filename>", methods=["POST"])
@login_required
def delete(filename):
    try:
        # Only protect critical system directories
        critical_dirs = [SCRIPT_DIR, TEMPLATE_DIR, STATIC_DIR, PROJECT_ROOT]
        
        file_path = None
        # Search for file in BASE_DIR
        for root, _, files in os.walk(BASE_DIR):
            if os.path.basename(filename) in files:
                candidate = os.path.join(root, os.path.basename(filename))
                if os.path.exists(candidate):
                    file_path = candidate
                    break
        
        if not file_path:
            # Check if filename is a directory
            for root, dirs, _ in os.walk(BASE_DIR):
                if os.path.basename(filename) in dirs:
                    candidate = os.path.join(root, os.path.basename(filename))
                    if os.path.exists(candidate):
                        file_path = candidate
                        break
        
        if not file_path:
            flash(f"Path '{filename}' not found.", "error")
            logger.warning(f"⚠️ Deletion failed: Path not found - {filename}")
            return redirect(request.referrer or url_for("file_browser"))
        
        # Check if path is critical
        abs_file_path = os.path.abspath(file_path)
        for critical_dir in critical_dirs:
            abs_critical_dir = os.path.abspath(critical_dir)
            if (abs_file_path == abs_critical_dir or 
                os.path.dirname(abs_file_path) == abs_critical_dir):
                flash(f"Cannot delete critical system path: {filename}", "error")
                logger.warning(f"⚠️ Deletion blocked: Attempted to delete critical system path - {filename}")
                return redirect(request.referrer or url_for("file_browser"))
        
        # Confirm deletion
        if request.form.get("confirm") != "true":
            flash(f"Deletion of '{filename}' requires confirmation.", "warning")
            logger.info(f"🛑 Deletion of {filename} requires confirmation")
            return redirect(request.referrer or url_for("file_browser"))
        
        # Perform deletion
        if os.path.isfile(file_path):
            os.remove(file_path)
            flash(f"File '{filename}' deleted successfully.", "success")
            logger.info(f"✅ Deleted file: {file_path}")
        elif os.path.isdir(file_path):
            shutil.rmtree(file_path, ignore_errors=True)
            flash(f"Folder '{filename}' deleted successfully.", "success")
            logger.info(f"✅ Deleted folder: {file_path}")
        else:
            flash(f"Path '{filename}' does not exist.", "error")
            logger.warning(f"⚠️ Deletion failed: Path does not exist - {file_path}")
        
        return redirect(request.referrer or url_for("file_browser"))
    except Exception as e:
        app.logger.error(f"Delete error: {e}")
        logger.error(f"❌ Delete error for {filename}: {e}")
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

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    mode = "local" if is_local_environment() else "cloud"
    logger.info(f"🚀 Starting Flask app in {mode.upper()} mode on port {port}...")
    app.run(host="0.0.0.0", port=port, debug=True)