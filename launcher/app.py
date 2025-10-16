import os
import subprocess
import re
import sys
from flask import Flask, request, redirect, url_for, render_template, flash, session, send_from_directory, send_file
from functools import wraps
from dotenv import load_dotenv
from jinja2 import TemplateNotFound, UndefinedError
import logging
import zipfile
import glob
from datetime import datetime
from werkzeug.utils import secure_filename

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

is_railway = 'RAILWAY_ENVIRONMENT' in os.environ

# Railway-specific logging
if os.getenv('RAILWAY_ENVIRONMENT'):
    # Enable more verbose logging on Railway
    logging.basicConfig(level=logging.DEBUG)
    logger.setLevel(logging.DEBUG)
    
    # Log important startup information
    logger.info("=== Railway Startup Debug ===")
    logger.info(f"Current directory: {os.getcwd()}")
    logger.info(f"Directory contents: {os.listdir('.')}")
    
    # Check for launcher directory
    if os.path.exists('launcher'):
        logger.info(f"Launcher contents: {os.listdir('launcher')}")
    
    # Check for scripts directory
    if os.path.exists('scripts'):
        logger.info(f"Scripts contents: {os.listdir('scripts')}")
        
app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "default_secret_key_1234567890")
load_dotenv()

PASSWORD = os.getenv("STUDENT_CLEANER_PASSWORD", "admin")
COLLEGE = os.getenv("COLLEGE_NAME", "FCT College of Nursing Sciences, Gwagwalada")
DEPARTMENT = os.getenv("DEPARTMENT", "Examinations Office")

# Allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv', 'zip'}
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def is_running_on_railway():
    """Check if running on Railway platform"""
    return 'RAILWAY_ENVIRONMENT' in os.environ

# Railway-compatible directory setup
def setup_railway_directories():
    """Create necessary directories - works for both local and Railway"""
    # Simple approach - use environment variables or fallbacks
    is_railway = 'RAILWAY_ENVIRONMENT' in os.environ
    
    if is_railway:
        base_dir = os.getenv('BASE_DIR', '/app/EXAMS_INTERNAL')
        upload_dir = os.getenv('UPLOAD_DIR', '/tmp/uploads')
        processed_dir = os.getenv('PROCESSED_DIR', '/tmp/processed')
    else:
        base_dir = os.getenv('BASE_DIR', "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/EXAMS_INTERNAL")
        upload_dir = os.getenv('UPLOAD_DIR', "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/uploads")
        processed_dir = os.getenv('PROCESSED_DIR', "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/processed")
    
    # Create directories
    for directory in [base_dir, upload_dir, processed_dir]:
        os.makedirs(directory, exist_ok=True)
        logger.info(f"Ensured directory exists: {directory}")
    
    return base_dir, upload_dir, processed_dir

# Setup directories on startup - with error handling
try:
    BASE_DIR, UPLOAD_DIR, PROCESSED_DIR = setup_railway_directories()
    logger.info(f"Directories initialized: BASE_DIR={BASE_DIR}")
except Exception as e:
    logger.error(f"Failed to setup directories: {e}")
    # Fallback to local directories
    BASE_DIR = os.path.join(os.path.dirname(__file__), '..', 'EXAMS_INTERNAL')
    UPLOAD_DIR = os.path.join(os.path.dirname(__file__), '..', 'uploads') 
    PROCESSED_DIR = os.path.join(os.path.dirname(__file__), '..', 'processed')
    os.makedirs(BASE_DIR, exist_ok=True)
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(PROCESSED_DIR, exist_ok=True)

SCRIPT_MAP = {
    "utme": "scripts/utme_result.py",
    "caosce": "scripts/caosce_result.py",
    "clean": "scripts/clean_results.py",
    "split": "scripts/split_names.py",
    "exam_processor": "scripts/exam_result_processor.py"
}

SUCCESS_INDICATORS = {
    "utme": [
        r"Processing: (PUTME 2025-Batch\d+[A-Z] Post-UTME Quiz-grades\.xlsx)",
        r"Saved processed file: (UTME_RESULT_.*?\.csv)",
        r"Saved processed file: (UTME_RESULT_.*?\.xlsx)",
        r"Saved processed file: (PUTME_COMBINE_RESULT_.*?\.xlsx)"
    ],
    "caosce": [
        r"Processed (CAOSCE SET2023A.*?|VIVA \([0-9]+\)\.xlsx) \(\d+ rows read\)",
        r"Saved processed file: (CAOSCE_RESULT_.*?\.csv)"
    ],
    "clean": [
        r"Processing: (Set2025-.*?\.xlsx)",
        r"✅ Cleaned CSV saved in.*?cleaned_(Set2025-.*?\.csv)",
        r"🎉 Master CSV saved in.*?master_cleaned_results\.csv",
        r"✅ All processing completed successfully!"
    ],
    "split": [r"Saved processed file: (clean_jamb_DB_.*?\.csv)"],
    "exam_processor": [
        r"(?i)processing\s*semester[: ]+\s*(nd[-a-z0-9]+-semester)",
        r"(?i)processing\s*nd[-a-z0-9]+-semester",
        r"(?i)✅\s*successfully\s*processed.*\.xlsx",
        r"(?i)✅\s*mastersheet\s*saved[: ]+.*\.xlsx",
        r"(?i)✅\s*nd\s*examination\s*results\s*processing\s*completed\s*successfully",
        r"(?i)📁\s*found\s*\d+\s*raw\s*files",
        r"(?i)processing[: ]+.*\.xlsx",
        r"(?i)🔄\s*applying\s*upgrade\s*rule[: ].*→\s*50",
        r"(?i)✅\s*upgraded\s*\d+\s*scores\s*from.*to\s*50"
    ]
}

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "logged_in" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

@app.route("/", methods=["GET"])
def index():
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        password = request.form.get("password")
        if password == PASSWORD:
            session["logged_in"] = True
            flash("Successfully logged in!")
            return redirect(url_for("dashboard"))
        else:
            flash("Invalid password. Please try again.")
    
    # Detect environment for display
    environment = "Railway Production" if is_running_on_railway() else "Local Development"
    return render_template("login.html", college=COLLEGE, environment=environment)

@app.route("/dashboard")
@login_required
def dashboard():
    environment = "Railway Production" if is_running_on_railway() else "Local Development"
    return render_template("dashboard.html", college=COLLEGE, DEPARTMENT=DEPARTMENT, environment=environment)

# NEW: Organized File Management Routes
@app.route("/upload-center")
@login_required
def upload_center():
    """Central upload hub for all file types"""
    return render_template("upload_center.html", college=COLLEGE, department=DEPARTMENT)

@app.route("/download-center")
@login_required
def download_center():
    """Central download hub with organized files"""
    files_by_category = get_organized_files()
    return render_template("download_center.html", 
                         files_by_category=files_by_category,
                         college=COLLEGE, 
                         department=DEPARTMENT)

@app.route("/file-browser")
@login_required
def file_browser():
    """Browse files by category and type"""
    file_structure = get_file_structure()
    return render_template("file_browser.html",
                         file_structure=file_structure,
                         college=COLLEGE,
                         department=DEPARTMENT)

@app.route("/download-zip/<category>/<folder_name>")
@login_required
def download_zip(category, folder_name):
    """Download entire folder as ZIP"""
    try:
        # Security check
        if category not in ['nd_sets', 'processed', 'raw']:
            flash("Invalid category")
            return redirect(url_for('download_center'))
        
        if category == 'nd_sets':
            zip_path = create_nd_set_zip(folder_name)
        else:
            zip_path = create_category_zip(category, folder_name)
        
        if zip_path and os.path.exists(zip_path):
            return send_file(zip_path, as_attachment=True)
        else:
            flash("Could not create zip file")
            return redirect(url_for('download_center'))
            
    except Exception as e:
        flash(f"Error creating zip: {str(e)}")
        return redirect(url_for('download_center'))

# NEW: File Organization Functions
def get_organized_files():
    """Get all files organized by category and timestamp"""
    files_by_category = {
        'nd_results': [],
        'putme_results': [],
        'caosce_results': [],
        'internal_results': [],
        'jamb_results': [],
        'raw_files': []
    }
    
    # Find ND processed files
    nd_pattern = os.path.join(BASE_DIR, "**", "*.xlsx")
    for file_path in glob.glob(nd_pattern, recursive=True):
        if 'CLEAN_RESULTS' in file_path or 'PROCESSED' in file_path or 'MASTERSHEET' in file_path:
            rel_path = os.path.relpath(file_path, BASE_DIR)
            files_by_category['nd_results'].append({
                'name': os.path.basename(file_path),
                'path': file_path,
                'relative_path': rel_path,
                'size': os.path.getsize(file_path),
                'modified': os.path.getmtime(file_path),
                'folder': os.path.dirname(rel_path)
            })
    
    # Find other processed files
    other_patterns = [
        (os.path.join(BASE_DIR, "**", "*UTME*.*"), 'putme_results'),
        (os.path.join(BASE_DIR, "**", "*CAOSCE*.*"), 'caosce_results'),
        (os.path.join(BASE_DIR, "**", "*CLEAN*.*"), 'internal_results'),
        (os.path.join(BASE_DIR, "**", "*JAMB*.*"), 'jamb_results'),
    ]
    
    for pattern, category in other_patterns:
        for file_path in glob.glob(pattern, recursive=True):
            if not file_path.startswith('~'):  # Skip temporary files
                rel_path = os.path.relpath(file_path, BASE_DIR)
                files_by_category[category].append({
                    'name': os.path.basename(file_path),
                    'path': file_path,
                    'relative_path': rel_path,
                    'size': os.path.getsize(file_path),
                    'modified': os.path.getmtime(file_path),
                    'folder': os.path.dirname(rel_path)
                })
    
    # Sort by modification time (newest first)
    for category in files_by_category:
        files_by_category[category].sort(key=lambda x: x['modified'], reverse=True)
    
    return files_by_category

def get_file_structure():
    """Get complete file structure"""
    structure = {
        'nd_sets': [],
        'processed_files': [],
        'raw_files': []
    }
    
    # ND Sets with CLEAN_RESULTS
    if os.path.exists(BASE_DIR):
        for item in os.listdir(BASE_DIR):
            item_path = os.path.join(BASE_DIR, item)
            if os.path.isdir(item_path) and item.startswith('ND-'):
                clean_path = os.path.join(item_path, "CLEAN_RESULTS")
                raw_path = os.path.join(item_path, "RAW_RESULTS")
                
                nd_set = {
                    'name': item,
                    'clean_results': [],
                    'raw_files': []
                }
                
                # Get clean results
                if os.path.exists(clean_path):
                    for file in os.listdir(clean_path):
                        if file.endswith(('.xlsx', '.csv', '.pdf')) and not file.startswith('~'):
                            file_path = os.path.join(clean_path, file)
                            nd_set['clean_results'].append({
                                'name': file,
                                'path': file_path,
                                'size': os.path.getsize(file_path),
                                'modified': os.path.getmtime(file_path)
                            })
                
                # Get raw files
                if os.path.exists(raw_path):
                    for file in os.listdir(raw_path):
                        if file.endswith(('.xlsx', '.xls', '.csv')) and not file.startswith('~'):
                            file_path = os.path.join(raw_path, file)
                            nd_set['raw_files'].append({
                                'name': file,
                                'path': file_path,
                                'size': os.path.getsize(file_path),
                                'modified': os.path.getmtime(file_path)
                            })
                
                # Sort files by modification time
                nd_set['clean_results'].sort(key=lambda x: x['modified'], reverse=True)
                nd_set['raw_files'].sort(key=lambda x: x['modified'], reverse=True)
                
                structure['nd_sets'].append(nd_set)
    
    return structure

def create_nd_set_zip(nd_set_name):
    """Create ZIP for entire ND set"""
    nd_set_path = os.path.join(BASE_DIR, nd_set_name)
    if not os.path.exists(nd_set_path):
        return None
    
    zip_filename = f"{nd_set_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
    zip_path = os.path.join(PROCESSED_DIR, zip_filename)
    
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        # Add CLEAN_RESULTS
        clean_path = os.path.join(nd_set_path, "CLEAN_RESULTS")
        if os.path.exists(clean_path):
            for root, dirs, files in os.walk(clean_path):
                for file in files:
                    if not file.startswith('~'):  # Skip temp files
                        file_path = os.path.join(root, file)
                        arcname = os.path.join(nd_set_name, "PROCESSED_RESULTS", file)
                        zipf.write(file_path, arcname)
        
        # Add RAW_RESULTS
        raw_path = os.path.join(nd_set_path, "RAW_RESULTS")
        if os.path.exists(raw_path):
            for root, dirs, files in os.walk(raw_path):
                for file in files:
                    if not file.startswith('~'):  # Skip temp files
                        file_path = os.path.join(root, file)
                        arcname = os.path.join(nd_set_name, "RAW_FILES", file)
                        zipf.write(file_path, arcname)
    
    return zip_path

def create_category_zip(category, folder_name):
    """Create ZIP for specific category"""
    # This function can be expanded for other categories
    zip_filename = f"{category}_{folder_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
    zip_path = os.path.join(PROCESSED_DIR, zip_filename)
    
    # For now, return None - can be implemented for other categories
    return None

# Update the existing upload function to create timestamped folders
@app.route("/upload/<script_name>", methods=["GET", "POST"])
@login_required
def upload_files(script_name):
    """Handle file uploads with timestamped organization"""
    if request.method == "POST":
        if 'file' not in request.files:
            flash('No file selected')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('No file selected')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            upload_path = get_upload_directory(script_name)
            
            # Create timestamped folder
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            timestamped_path = os.path.join(upload_path, timestamp)
            os.makedirs(timestamped_path, exist_ok=True)
            
            file_path = os.path.join(timestamped_path, filename)
            file.save(file_path)
            
            flash(f'File "{filename}" uploaded successfully to {timestamp} folder!')
            return redirect(url_for('upload_center'))
        else:
            flash('Invalid file type. Allowed: xlsx, xls, csv, zip')
    
    script_descriptions = {
        "utme": "PUTME Examination Results",
        "caosce": "CAOSCE Examination Results", 
        "clean": "Objective Examination Results",
        "split": "JAMB Candidate Database",
        "exam_processor": "ND Examination Results"
    }
    
    script_desc = script_descriptions.get(script_name, "Script")
    return render_template("upload_form.html", 
                         script_name=script_name, 
                         script_desc=script_desc,
                         college=COLLEGE, 
                         department=DEPARTMENT)

def get_upload_directory(script_name):
    """Get the appropriate upload directory for each script type"""
    is_railway = 'RAILWAY_ENVIRONMENT' in os.environ
    
    if is_railway:
        # Railway paths
        base_dir = '/app/EXAMS_INTERNAL'
        if script_name == "exam_processor":
            # For exam processor, create or use ND set directory
            nd_sets = [d for d in os.listdir(base_dir) if d.startswith('ND-') and os.path.isdir(os.path.join(base_dir, d))]
            if nd_sets:
                # Use the first ND set found
                nd_set = nd_sets[0]
                upload_path = os.path.join(base_dir, nd_set, "RAW_RESULTS")
                os.makedirs(upload_path, exist_ok=True)
                return upload_path
            else:
                # Create a default ND set
                default_set = "ND-2024"
                upload_path = os.path.join(base_dir, default_set, "RAW_RESULTS")
                os.makedirs(upload_path, exist_ok=True)
                return upload_path
        elif script_name == "utme":
            upload_path = os.path.join(base_dir, "PUTME_RESULT", "RAW_PUTME_RESULT")
        elif script_name == "caosce":
            upload_path = os.path.join(base_dir, "CAOSCE_RESULT", "RAW_CAOSCE_RESULT")
        elif script_name == "clean":
            upload_path = os.path.join(base_dir, "INTERNAL_RESULT", "RAW_INTERNAL_RESULT")
        elif script_name == "split":
            upload_path = os.path.join(base_dir, "JAMB_DB", "RAW_JAMB_DB")
        else:
            upload_path = '/tmp/uploads'
    else:
        # Local development paths
        if script_name == "exam_processor":
            upload_path = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/EXAMS_INTERNAL/ND-2024/RAW_RESULTS"
        elif script_name == "utme":
            upload_path = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/PUTME_RESULT/RAW_PUTME_RESULT"
        elif script_name == "caosce":
            upload_path = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/CAOSCE_RESULT/RAW_CAOSCE_RESULT"
        elif script_name == "clean":
            upload_path = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/INTERNAL_RESULT/RAW_INTERNAL_RESULT"
        elif script_name == "split":
            upload_path = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/JAMB_DB/RAW_JAMB_DB"
        else:
            upload_path = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/uploads"
    
    # Ensure the directory exists
    os.makedirs(upload_path, exist_ok=True)
    return upload_path

# Keep existing file checking functions
def check_exam_processor_files(input_dir):
    """Check for ND examination files - Railway compatible"""
    logger.info(f"Checking for exam files in: {input_dir}")
    
    if not os.path.isdir(input_dir):
        logger.warning(f"Input directory not found: {input_dir}")
        return False
    
    try:
        nd_sets = []
        for item in os.listdir(input_dir):
            item_path = os.path.join(input_dir, item)
            if os.path.isdir(item_path) and item.startswith('ND-') and item != 'ND-COURSES':
                nd_sets.append(item)
        
        logger.info(f"Found ND sets: {nd_sets}")
        
        if not nd_sets:
            return False
        
        total_files_found = 0
        for nd_set in nd_sets:
            set_path = os.path.join(input_dir, nd_set)
            if not os.path.isdir(set_path):
                continue
            
            raw_results_path = os.path.join(set_path, "RAW_RESULTS")
            if not os.path.isdir(raw_results_path):
                logger.warning(f"RAW_RESULTS not found in: {set_path}")
                continue
                
            excel_files = [f for f in os.listdir(raw_results_path) 
                         if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~')]
            
            total_files_found += len(excel_files)
            logger.info(f"Found {len(excel_files)} files in {nd_set}")
        
        logger.info(f"Total files found: {total_files_found}")
        return total_files_found > 0
        
    except Exception as e:
        logger.error(f"Error checking exam files: {e}")
        return False

def check_putme_files(input_dir):
    """Check for PUTME examination files"""
    if not os.path.isdir(input_dir):
        return False
    
    excel_files = [f for f in os.listdir(input_dir) 
                  if f.lower().endswith(('.xlsx', '.xls')) and 'PUTME' in f.upper()]
    
    candidate_batches_dir = os.path.join(os.path.dirname(input_dir), "RAW_CANDIDATE_BATCHES")
    batch_files = []
    if os.path.isdir(candidate_batches_dir):
        batch_files = [f for f in os.listdir(candidate_batches_dir) 
                      if f.lower().endswith('.csv') and 'BATCH' in f.upper()]
    
    return len(excel_files) > 0 and len(batch_files) > 0

def check_internal_exam_files(input_dir):
    """Check for internal exam files"""
    if not os.path.isdir(input_dir):
        return False
    
    excel_files = [f for f in os.listdir(input_dir) 
                  if f.lower().endswith(('.xlsx', '.xls')) and f.startswith('Set')]
    
    return len(excel_files) > 0

def check_caosce_files(input_dir):
    """Check for CAOSCE exam files"""
    if not os.path.isdir(input_dir):
        return False
    
    excel_files = [f for f in os.listdir(input_dir) 
                  if f.lower().endswith(('.xlsx', '.xls')) and 'CAOSCE' in f.upper()]
    
    return len(excel_files) > 0

def check_split_files(input_dir):
    """Check for JAMB split files"""
    if not os.path.isdir(input_dir):
        return False
    
    valid_files = [f for f in os.listdir(input_dir) 
                  if f.lower().endswith(('.csv', '.xlsx', '.xls')) and not f.startswith('~')]
    
    return len(valid_files) > 0

def check_input_files(input_dir, script_name):
    """Check for input files based on script type"""
    if not os.path.isdir(input_dir):
        return False
    
    if script_name == "exam_processor":
        return check_exam_processor_files(input_dir)
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
        valid_extensions = ('.csv', '.xlsx', '.xls')
        input_files = [f for f in dir_contents if f.lower().endswith(valid_extensions) and not f.startswith('~')]
        return len(input_files) > 0
    except Exception:
        return False

def count_processed_files(output_lines, script_name, selected_semesters=None):
    """Count processed semesters based on script output"""
    success_indicators = SUCCESS_INDICATORS.get(script_name, [])
    processed_files_set = set()
    
    # Log all output lines for debugging
    print(f"Raw output lines for {script_name}:")
    for output_line in output_lines:
        if output_line.strip():
            print(f"  OUTPUT: {output_line}")
    
    for output_line in output_lines:
        for indicator in success_indicators:
            match = re.search(indicator, output_line, re.IGNORECASE)
            if match:
                if script_name == "utme":
                    if "Processing:" in output_line:
                        file_name = match.group(1)
                        processed_files_set.add(f"Processed: {file_name}")
                    elif "Saved processed file:" in output_line:
                        file_name = match.group(1)
                        processed_files_set.add(f"Saved: {file_name}")
                elif script_name == "clean":
                    if "Processing:" in output_line:
                        file_name = match.group(1)
                        processed_files_set.add(f"Processed: {file_name}")
                    elif "✅ Cleaned CSV saved" in output_line:
                        file_name = match.group(1) if match.groups() else "cleaned_file"
                        processed_files_set.add(f"Cleaned: {file_name}")
                    elif "🎉 Master CSV saved" in output_line:
                        processed_files_set.add("Master file created")
                    elif "✅ All processing completed successfully!" in output_line:
                        processed_files_set.add("Processing completed")
                elif script_name == "exam_processor":
                    # Only count semesters explicitly processed
                    if "PROCESSING SEMESTER:" in output_line:
                        semester = match.group(1)
                        processed_files_set.add(f"Semester: {semester}")
                else:
                    file_name = match.group(1) if match.groups() else output_line
                    processed_files_set.add(file_name)
    
    # For exam_processor in manual mode, strictly validate against selected semesters
    if script_name == "exam_processor" and selected_semesters:
        semester_mapping = {
            'first_first': 'ND-FIRST-YEAR-FIRST-SEMESTER',
            'first_second': 'ND-FIRST-YEAR-SECOND-SEMESTER',
            'second_first': 'ND-SECOND-YEAR-FIRST-SEMESTER', 
            'second_second': 'ND-SECOND-YEAR-SECOND-SEMESTER'
        }
        expected_semesters = {f"Semester: {semester_mapping.get(sem, sem)}" for sem in selected_semesters}
        processed_files_set = processed_files_set.intersection(expected_semesters)
        print(f"Expected semesters: {expected_semesters}")
        if len(processed_files_set) != len(expected_semesters):
            print(f"WARNING: Processed semester count ({len(processed_files_set)}) does not match expected ({len(expected_semesters)})")
    
    print(f"Processed items for {script_name}: {processed_files_set}")
    if script_name == "exam_processor":
        print("🔍 Detected semester lines:")
        for item in processed_files_set:
            print(f"   {item}")
    return len(processed_files_set)

def get_success_message(script_name, processed_files, output_lines, selected_semesters=None):
    """Generate appropriate success message based on script and output"""
    if processed_files == 0:
        return None
    
    if script_name == "clean":
        if any("✅ All processing completed successfully!" in line for line in output_lines):
            return f"Successfully processed internal examination results! Generated master file and individual cleaned files."
        else:
            return f"Processed {processed_files} internal examination file(s)."
    
    elif script_name == "exam_processor":
        upgrade_info = ""
        upgrade_count = ""
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
        
        # Check for mismatch in manual mode
        if selected_semesters and len(selected_semesters) != processed_files:
            return f"Warning: Expected to process {len(selected_semesters)} semester(s), but processed {processed_files}. Please check logs.{upgrade_info}{upgrade_count}"
        
        if any("✅ ND Examination Results Processing completed successfully" in line for line in output_lines):
            return f"ND Examination processing completed successfully! Processed {processed_files} semester(s).{upgrade_info}{upgrade_count}"
        elif any("✅ Processing complete" in line for line in output_lines):
            return f"ND Examination processing completed! Processed {processed_files} semester(s).{upgrade_info}{upgrade_count}"
        else:
            return f"Processed {processed_files} ND examination semester(s).{upgrade_info}{upgrade_count}"
    
    elif script_name == "utme":
        if any("Processing completed successfully" in line for line in output_lines):
            return f"PUTME processing completed successfully! Processed {processed_files} batch file(s)."
        else:
            return f"Processed {processed_files} PUTME batch file(s)."
    
    elif script_name == "caosce":
        if any("Processed" in line for line in output_lines):
            return f"CAOSCE processing completed! Processed {processed_files} file(s)."
        else:
            return f"Processed {processed_files} CAOSCE file(s)."
    
    elif script_name == "split":
        if any("Saved processed file:" in line for line in output_lines):
            return f"JAMB name splitting completed! Processed {processed_files} file(s)."
        else:
            return f"Processed {processed_files} JAMB file(s)."
    
    else:
        return f"Successfully processed {processed_files} file(s)."

def get_script_paths():
    """Get script paths that work in both local and Railway environments"""
    is_railway = 'RAILWAY_ENVIRONMENT' in os.environ
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    script_paths = {}
    
    for script_name, relative_path in SCRIPT_MAP.items():
        if is_railway:
            # Railway paths
            possible_paths = [
                os.path.join('/app', relative_path),  # Railway root
                os.path.join('/app', 'scripts', os.path.basename(relative_path)),  # Direct scripts folder
                os.path.join('/app', 'student_result_cleaner', relative_path),  # Railway subdirectory
            ]
        else:
            # Local development paths
            possible_paths = [
                os.path.join(project_root, relative_path),  # Local development
                os.path.join(project_root, 'scripts', os.path.basename(relative_path)),  # Local scripts folder
                os.path.join('/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT', 'scripts', os.path.basename(relative_path)),  # Absolute local path
            ]
        
        found_path = None
        for path in possible_paths:
            if os.path.exists(path):
                found_path = path
                logger.info(f"Found script {script_name} at: {path}")
                break
        
        if found_path:
            script_paths[script_name] = found_path
        else:
            logger.error(f"Script not found: {script_name} at any of {possible_paths}")
            script_paths[script_name] = f"MISSING: {relative_path}"
    
    return script_paths

@app.route("/run/<script_name>", methods=["GET", "POST"])
@login_required
def run_script(script_name):
    try:
        script_paths = get_script_paths()
        
        if script_name not in script_paths:
            logger.error(f"Script '{script_name}' not found in paths: {list(script_paths.keys())}")
            flash(f"Script '{script_name}' not found. Available scripts: {', '.join(script_paths.keys())}")
            return redirect(url_for("dashboard"))

        script_path = script_paths[script_name]
        
        # ===== DEBUGGING =====
        logger.info(f"=== RUN_SCRIPT DEBUG ===")
        logger.info(f"Script name: {script_name}")
        logger.info(f"Request method: {request.method}")
        logger.info(f"Available scripts: {list(script_paths.keys())}")
        logger.info(f"Exam processor path: {script_paths.get('exam_processor')}")
        logger.info(f"Script exists: {os.path.exists(script_path)}")
        
        # Check if the path is valid
        if script_path.startswith('MISSING:') or not os.path.exists(script_path):
            logger.error(f"Script path invalid: {script_path}")
            flash(f"Script file not found at: {script_path}")
            return redirect(url_for("dashboard"))

        script_desc = {
            "utme": "PUTME Examination Results",
            "caosce": "CAOSCE Examination Results", 
            "clean": "Objective Examination Results",
            "split": "JAMB Candidate Database",
            "exam_processor": "ND Examination Results Processor"
        }.get(script_name, "Script")

        logger.info(f"Running script: {script_name} from {script_path}")

        # Dynamic input directories that work for both local and Railway
        is_railway = 'RAILWAY_ENVIRONMENT' in os.environ

        if is_railway:
            # Railway paths
            input_dirs = {
                "utme": os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_PUTME_RESULT"),
                "caosce": os.path.join(BASE_DIR, "CAOSCE_RESULT", "RAW_CAOSCE_RESULT"),
                "clean": os.path.join(BASE_DIR, "INTERNAL_RESULT", "RAW_INTERNAL_RESULT"),
                "split": os.path.join(BASE_DIR, "JAMB_DB", "RAW_JAMB_DB"),
                "exam_processor": BASE_DIR
            }
        else:
            # Local development paths - your existing local paths
            input_dirs = {
                "utme": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/PUTME_RESULT/RAW_PUTME_RESULT",
                "caosce": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/CAOSCE_RESULT/RAW_CAOSCE_RESULT",
                "clean": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/INTERNAL_RESULT/RAW_INTERNAL_RESULT",
                "split": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/JAMB_DB/RAW_JAMB_DB",
                "exam_processor": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/EXAMS_INTERNAL"
            }
        
        input_dir = input_dirs.get(script_name, BASE_DIR)
        
        # Create input directory if it doesn't exist
        os.makedirs(input_dir, exist_ok=True)
        
        if not check_input_files(input_dir, script_name):
            flash(f"No input files found for {script_desc}. Please upload files to the appropriate directory.")
            return redirect(url_for("upload_center"))

        # Handle exam processor with form parameters
        if script_name == "exam_processor" and request.method == "POST":
            return handle_exam_processor(script_path, script_name, input_dir)
        
        # For GET requests on exam_processor, show the form
        if script_name == "exam_processor" and request.method == "GET":
            nd_sets = []
            if os.path.isdir(input_dir):
                for item in os.listdir(input_dir):
                    item_path = os.path.join(input_dir, item)
                    if os.path.isdir(item_path) and item.startswith('ND-') and item != 'ND-COURSES':
                        nd_sets.append(item)
            
            return render_template(
                "exam_processor_form.html", 
                college=COLLEGE,
                department=DEPARTMENT,
                nd_sets=nd_sets
            )
        
        # Run other scripts
        try:
            result = subprocess.run(
                [sys.executable, script_path],
                env=os.environ.copy(),
                text=True,
                capture_output=True,
                timeout=600,
                cwd=os.path.dirname(script_path)
            )
            
            output_lines = result.stdout.splitlines() if result.stdout else []
            processed_files = count_processed_files(output_lines, script_name)
            success_msg = get_success_message(script_name, processed_files, output_lines)
            
            if success_msg:
                flash(success_msg)
            else:
                flash(f"Script completed but no files processed for {script_desc}.")
                
        except subprocess.TimeoutExpired:
            flash(f"Script timed out but may still be running in background.")
        except subprocess.CalledProcessError as e:
            flash(f"Error processing {script_desc}: {e.stderr or str(e)}")
            
        return redirect(url_for("dashboard"))
        
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        logger.error(f"CRITICAL ERROR in run_script: {error_details}")
        
        return f"""
        <h1>Internal Server Error - Debug Info</h1>
        <h3>Error Details:</h3>
        <pre>{error_details}</pre>
        <h3>Script Info:</h3>
        <p>Script: {script_name}</p>
        <p>Path: {script_path if 'script_path' in locals() else 'NOT FOUND'}</p>
        <p>Method: {request.method}</p>
        <br>
        <a href="{url_for('dashboard')}">Back to Dashboard</a>
        """, 500

def handle_exam_processor(script_path, script_name, input_dir):
    """Handle exam processor execution with form parameters"""
    selected_set = request.form.get("selected_set", "all")
    processing_mode = request.form.get("processing_mode", "auto")
    selected_semesters = request.form.getlist("semesters")
    pass_threshold = request.form.get("pass_threshold", "50.0")
    upgrade_threshold = request.form.get("upgrade_threshold", "0")
    generate_pdf = "generate_pdf" in request.form
    track_withdrawn = "track_withdrawn" in request.form
    
    logger.info(f"Exam processor parameters: set={selected_set}, mode={processing_mode}, semesters={selected_semesters}")

    env = os.environ.copy()
    env.update({
        'SELECTED_SET': selected_set,
        'PROCESSING_MODE': processing_mode,
        'PASS_THRESHOLD': pass_threshold,
        'BASE_DIR': BASE_DIR,  # Ensure BASE_DIR is set
        'GENERATE_PDF': str(generate_pdf),
        'TRACK_WITHDRAWN': str(track_withdrawn)
    })
    
    if upgrade_threshold and upgrade_threshold.strip() and upgrade_threshold != "0":
        env['UPGRADE_THRESHOLD'] = upgrade_threshold.strip()
    
    if processing_mode == "manual" and selected_semesters:
        semester_mapping = {
            'first_first': 'ND-FIRST-YEAR-FIRST-SEMESTER',
            'first_second': 'ND-FIRST-YEAR-SECOND-SEMESTER',
            'second_first': 'ND-SECOND-YEAR-FIRST-SEMESTER', 
            'second_second': 'ND-SECOND-YEAR-SECOND-SEMESTER'
        }
        
        selected_semester_keys = [semester_mapping.get(sem, sem) for sem in selected_semesters]
        env['SELECTED_SEMESTERS'] = ','.join(selected_semester_keys)

    try:
        result = subprocess.run(
            [sys.executable, script_path],
            env=env,
            text=True,
            capture_output=True,
            timeout=600,
            cwd=os.path.dirname(script_path)
        )
        
        output_lines = result.stdout.splitlines() if result.stdout else []
        error_lines = result.stderr.splitlines() if result.stderr else []
        
        logger.info(f"Script output: {len(output_lines)} lines")
        logger.info(f"Script errors: {len(error_lines)} lines")
        
        if result.returncode == 0:
            processed_files = count_processed_files(output_lines, script_name, selected_semesters if processing_mode == "manual" else None)
            success_msg = get_success_message(script_name, processed_files, output_lines, selected_semesters if processing_mode == "manual" else None)
            
            if success_msg:
                flash(success_msg)
            else:
                flash("Processing completed but no specific success message detected.")
        else:
            error_msg = result.stderr or "Unknown error occurred"
            flash(f"Script failed: {error_msg[:200]}")
            
    except subprocess.TimeoutExpired:
        flash("Processing timed out. The operation took too long to complete.")
    except Exception as e:
        flash(f"Error running exam processor: {str(e)}")
    
    return redirect(url_for("dashboard"))

# Keep existing routes for backward compatibility
@app.route("/downloads")
@login_required
def list_downloads():
    """Legacy download route - redirect to new download center"""
    return redirect(url_for('download_center'))

@app.route("/file-explorer")
@login_required
def file_explorer():
    """Legacy file explorer route - redirect to new file browser"""
    return redirect(url_for('file_browser'))

@app.route("/download/<path:filename>")
@login_required
def download_file(filename):
    """Download a specific file"""
    try:
        # Security check - ensure the file is within allowed directories
        full_path = os.path.join(BASE_DIR, filename)
        if not os.path.commonpath([os.path.abspath(full_path), os.path.abspath(BASE_DIR)]) == os.path.abspath(BASE_DIR):
            flash("Access denied")
            return redirect(url_for('download_center'))
        
        return send_file(full_path, as_attachment=True)
    except Exception as e:
        flash(f"Error downloading file: {str(e)}")
        return redirect(url_for('download_center'))

@app.route("/logout")
@login_required
def logout():
    session.pop("logged_in", None)
    flash("You have been logged out.")
    return redirect(url_for("login"))

@app.route('/health')
def health_check():
    """Health check endpoint for Railway"""
    return {'status': 'healthy', 'environment': os.getenv('RAILWAY_ENVIRONMENT', 'unknown')}

@app.route('/debug/paths')
@login_required
def debug_paths():
    """Debug endpoint to check script paths"""
    script_paths = get_script_paths()
    path_info = []
    
    for name, path in script_paths.items():
        exists = os.path.exists(path)
        path_info.append({
            'script': name,
            'path': path,
            'exists': exists,
            'is_file': os.path.isfile(path) if exists else False
        })
    
    return render_template('debug_paths.html', paths=path_info)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)