import os
import subprocess
import re
import sys
from flask import Flask, request, redirect, url_for, render_template, flash, session
from functools import wraps
from dotenv import load_dotenv
from jinja2 import TemplateNotFound, UndefinedError

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "default_secret_key_1234567890")
load_dotenv()

PASSWORD = os.getenv("STUDENT_CLEANER_PASSWORD", "admin")
COLLEGE = os.getenv("COLLEGE_NAME", "FCT College of Nursing Sciences, Gwagwalada")
DEPARTMENT = os.getenv("DEPARTMENT", "Examinations Office")

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
    return render_template("login.html", college=COLLEGE)

@app.route("/dashboard")
@login_required
def dashboard():
    return render_template("dashboard.html", college=COLLEGE, DEPARTMENT=DEPARTMENT)

def check_exam_processor_files(input_dir):
    """Check for ND examination files in the actual directory structure"""
    if not os.path.isdir(input_dir):
        return False
    
    nd_sets = []
    for item in os.listdir(input_dir):
        item_path = os.path.join(input_dir, item)
        if os.path.isdir(item_path) and item.startswith('ND-') and item != 'ND-COURSES':
            nd_sets.append(item)
    
    if not nd_sets:
        return False
    
    total_files_found = 0
    for nd_set in nd_sets:
        set_path = os.path.join(input_dir, nd_set)
        if not os.path.isdir(set_path):
            continue
        
        raw_results_path = os.path.join(set_path, "RAW_RESULTS")
        if not os.path.isdir(raw_results_path):
            continue
            
        excel_files = [f for f in os.listdir(raw_results_path) 
                     if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~')]
        
        total_files_found += len(excel_files)
    
    return total_files_found > 0

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
    for line in output_lines:
        if line.strip():
            print(f"  OUTPUT: {line}")
    
    print(f"CHECKING LINE: {line}")
    for line in output_lines:
        for indicator in success_indicators:
            match = re.search(indicator, line, re.IGNORECASE)
            if match:
                if script_name == "utme":
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
                elif script_name == "exam_processor":
                    # Only count semesters explicitly processed
                    if "PROCESSING SEMESTER:" in line:
                        semester = match.group(1)
                        processed_files_set.add(f"Semester: {semester}")
                else:
                    file_name = match.group(1) if match.groups() else line
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

@app.route("/run/<script_name>", methods=["GET", "POST"])
@login_required
def run_script(script_name):
    try:
        if script_name not in SCRIPT_MAP:
            flash("Invalid script requested.")
            return redirect(url_for("dashboard"))

        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        script_path = os.path.join(project_root, SCRIPT_MAP[script_name])
        script_desc = {
            "utme": "PUTME Examination Results",
            "caosce": "CAOSCE Examination Results", 
            "clean": "Objective Examination Results",
            "split": "JAMB Candidate Database",
            "exam_processor": "ND Examination Results Processor"
        }.get(script_name, "Script")

        if not os.path.isfile(script_path):
            flash(f"Script file not found: {script_path}")
            return redirect(url_for("dashboard"))

        if not os.access(script_path, os.X_OK):
            try:
                os.chmod(script_path, 0o755)
            except Exception:
                flash(f"Script permissions error: {script_path}")
                return redirect(url_for("dashboard"))

        input_dir = {
            "utme": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/PUTME_RESULT/RAW_PUTME_RESULT",
            "caosce": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/CAOSCE_RESULT/RAW_CAOSCE_RESULT", 
            "clean": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/INTERNAL_RESULT/RAW_INTERNAL_RESULT",
            "split": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/JAMB_DB/RAW_JAMB_DB",
            "exam_processor": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/EXAMS_INTERNAL"
        }.get(script_name, "the input directory")
        
        if not check_input_files(input_dir, script_name):
            if script_name == "exam_processor":
                nd_sets_found = []
                if os.path.isdir(input_dir):
                    for item in os.listdir(input_dir):
                        item_path = os.path.join(input_dir, item)
                        if os.path.isdir(item_path) and item.startswith('ND-') and item != 'ND-COURSES':
                            raw_results_path = os.path.join(item_path, "RAW_RESULTS")
                            if os.path.isdir(raw_results_path):
                                excel_files = [f for f in os.listdir(raw_results_path) 
                                             if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~')]
                                nd_sets_found.append(f"{item} ({len(excel_files)} files in RAW_RESULTS)")
                            else:
                                nd_sets_found.append(f"{item} (no RAW_RESULTS folder)")
                
                if not nd_sets_found:
                    flash(f"No ND examination sets found in {input_dir}. Please ensure you have ND sets (ND-2024, ND-2025, etc.) with Excel files.")
                else:
                    flash(f"Found ND sets but issues with file detection: {', '.join(nd_sets_found)}. Please ensure Excel files are in the RAW_RESULTS subdirectory.")
                    
            elif script_name == "utme":
                flash(f"No PUTME examination files found in {input_dir}. Please ensure you have PUTME Excel files and candidate batch CSV files.")
            elif script_name == "clean":
                flash(f"No internal examination files found in {input_dir}. Please ensure you have Set2025 Excel files.")
            elif script_name == "caosce":
                flash(f"No CAOSCE examination files found in {input_dir}. Please ensure you have CAOSCE Excel files.")
            elif script_name == "split":
                files_found = []
                if os.path.isdir(input_dir):
                    files_found = [f for f in os.listdir(input_dir) 
                                 if f.lower().endswith(('.csv', '.xlsx', '.xls')) and not f.startswith('~')]
                
                if not files_found:
                    flash(f"No JAMB database files found in {input_dir}. Please ensure you have JAMB CSV or Excel files.")
                else:
                    flash(f"JAMB files found but may not be accessible: {', '.join(files_found)}. Please check file permissions and formats.")
            else:
                flash(f"No CSV or Excel files found in {input_dir}")
            return redirect(url_for("dashboard"))

        if script_name in ["utme", "exam_processor"]:
            if request.method == "GET":
                if script_name == "utme":
                    return render_template("utme_form.html", college=COLLEGE, department=DEPARTMENT)
                elif script_name == "exam_processor":
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

            if request.method == "POST":
                if script_name == "utme":
                    convert_value = request.form.get("convert_value", "").strip()
                    convert_column = request.form.get("convert_column", "n")

                    cmd = ["python3", script_path]
                    if convert_value:
                        cmd.extend(["--non-interactive", "--converted-score-max", convert_value])

                    try:
                        result = subprocess.run(
                            cmd,
                            input=f"{convert_column}\n",
                            text=True,
                            capture_output=True,
                            check=True,
                            timeout=300
                        )
                        output_lines = result.stdout.splitlines()
                        processed_files = count_processed_files(output_lines, script_name)
                        success_msg = get_success_message(script_name, processed_files, output_lines)
                        
                        if success_msg:
                            flash(success_msg)
                        else:
                            if "No CSV or Excel files found" in result.stdout:
                                flash(f"No CSV or Excel files found in {input_dir} for {script_desc}.")
                            else:
                                flash(f"No files processed for {script_desc}. Check input files in {input_dir}.")
                            
                    except subprocess.TimeoutExpired:
                        flash(f"Script timed out but may still be running in background.")
                    except subprocess.CalledProcessError as e:
                        output_lines = e.stdout.splitlines() if e.stdout else []
                        processed_files = count_processed_files(output_lines, script_name)
                        success_msg = get_success_message(script_name, processed_files, output_lines)
                        if success_msg:
                            flash(f"Partially completed: {success_msg}, but encountered an error: {e.stderr or str(e)}")
                        else:
                            flash(f"Error processing {script_desc}: {e.stderr or str(e)}")
                    return redirect(url_for("dashboard"))

                elif script_name == "exam_processor":
                    selected_set = request.form.get("selected_set", "all")
                    processing_mode = request.form.get("processing_mode", "auto")
                    selected_semesters = request.form.getlist("semesters")
                    pass_threshold = request.form.get("pass_threshold", "50.0")
                    upgrade_threshold = request.form.get("upgrade_threshold", "0")
                    generate_pdf = "generate_pdf" in request.form
                    track_withdrawn = "track_withdrawn" in request.form
                    
                    print(f"FORM DATA RECEIVED for exam_processor:")
                    print(f"   Selected Set: {selected_set}")
                    print(f"   Processing Mode: {processing_mode}") 
                    print(f"   Selected Semesters: {selected_semesters}")
                    print(f"   Pass Threshold: {pass_threshold}")
                    print(f"   Upgrade Threshold: '{upgrade_threshold}'")
                    print(f"   Generate PDF: {generate_pdf}")
                    print(f"   Track Withdrawn: {track_withdrawn}")

                    if not os.path.exists(script_path):
                        flash(f"Exam processor script not found at: {script_path}")
                        return redirect(url_for("dashboard"))
                    
                    if not os.access(script_path, os.R_OK):
                        flash(f"No read permission for script: {script_path}")
                        return redirect(url_for("dashboard"))

                    env = os.environ.copy()
                    env['SELECTED_SET'] = selected_set
                    env['PROCESSING_MODE'] = processing_mode
                    env['PASS_THRESHOLD'] = pass_threshold
                    
                    if upgrade_threshold and upgrade_threshold.strip() and upgrade_threshold != "0":
                        env['UPGRADE_THRESHOLD'] = upgrade_threshold.strip()
                        print(f"   UPGRADE THRESHOLD SET: {upgrade_threshold}")
                    else:
                        if 'UPGRADE_THRESHOLD' in env:
                            del env['UPGRADE_THRESHOLD']
                        print(f"   NO UPGRADE THRESHOLD SET")

                    env['GENERATE_PDF'] = str(generate_pdf)
                    env['TRACK_WITHDRAWN'] = str(track_withdrawn)
                    
                    if processing_mode == "manual" and selected_semesters:
                        semester_mapping = {
                            'first_first': 'ND-FIRST-YEAR-FIRST-SEMESTER',
                            'first_second': 'ND-FIRST-YEAR-SECOND-SEMESTER',
                            'second_first': 'ND-SECOND-YEAR-FIRST-SEMESTER', 
                            'second_second': 'ND-SECOND-YEAR-SECOND-SEMESTER'
                        }
                        
                        selected_semester_keys = []
                        for semester in selected_semesters:
                            if semester in semester_mapping:
                                selected_semester_keys.append(semester_mapping[semester])
                            elif semester.startswith('ND-'):
                                selected_semester_keys.append(semester)
                        
                        if selected_semester_keys:
                            env['SELECTED_SEMESTERS'] = ','.join(selected_semester_keys)
                            print(f"   SELECTED SEMESTERS SET: {selected_semester_keys}")
                        else:
                            env['SELECTED_SEMESTERS'] = ''
                            print(f"   NO SPECIFIC SEMESTERS SELECTED")
                    else:
                        env['SELECTED_SEMESTERS'] = ''
                        print(f"   AUTO MODE - PROCESSING ALL SEMESTERS")
                    
                    print(f"Environment variables set:")
                    for key in ['SELECTED_SET', 'PROCESSING_MODE', 'PASS_THRESHOLD', 'UPGRADE_THRESHOLD', 'GENERATE_PDF', 'TRACK_WITHDRAWN', 'SELECTED_SEMESTERS']:
                        print(f"   {key}: {env.get(key)}")

                    print(f"Starting exam processor script...")
                    result = subprocess.run(
                        [sys.executable, script_path],
                        env=env,
                        text=True,
                        capture_output=True,
                        timeout=600
                    )
                    
                    print(f"Script execution completed")
                    print(f"   Return code: {result.returncode}")
                    print(f"   stdout length: {len(result.stdout)}")
                    print(f"   stderr length: {len(result.stderr)}")
                    
                    output_lines = result.stdout.splitlines() if result.stdout else []
                    processed_files = count_processed_files(
                        output_lines,
                        script_name,
                        selected_semesters if processing_mode == "manual" else None
                    )
                    
                    if result.returncode == 0:
                        upgrade_applied = False
                        upgrade_details = ""
                        upgrade_count = ""
                        for line in output_lines:
                            if "🔄 Applying upgrade rule:" in line:
                                upgrade_match = re.search(r"🔄 Applying upgrade rule: (\d+)–49 → 50", line)
                                if upgrade_match:
                                    upgrade_applied = True
                                    upgrade_details = f" Upgrade rule applied: {upgrade_match.group(1)}-49 → 50"
                                    break
                            elif "✅ Upgraded" in line:
                                upgrade_count_match = re.search(r"✅ Upgraded (\d+) scores", line)
                                if upgrade_count_match:
                                    upgrade_count = f" Upgraded {upgrade_count_match.group(1)} scores"
                                    break
                        
                        success_msg = get_success_message(
                            script_name,
                            processed_files,
                            output_lines,
                            selected_semesters if processing_mode == "manual" else None
                        )
                        if success_msg:
                            flash(success_msg)
                        else:
                            flash(f"Script completed but no semesters processed.{upgrade_details}{upgrade_count}")
                    else:
                        error_msg = result.stderr or "No error output"
                        if "No module named" in error_msg:
                            flash("Missing Python dependencies. Please install: pandas openpyxl reportlab")
                        elif "FileNotFoundError" in error_msg:
                            flash("Required files not found. Check if ND sets have RAW_RESULTS folders with Excel files.")
                        elif "Permission denied" in error_msg:
                            flash("Permission error. Check file permissions in the exam directory.")
                        else:
                            flash(f"Script failed: {error_msg[:100]}...")
                            
                    return redirect(url_for("dashboard"))

        try:
            result = subprocess.run(
                [sys.executable, script_path],
                text=True,
                capture_output=True,
                check=True,
                timeout=300
            )
            output_lines = result.stdout.splitlines()
            processed_files = count_processed_files(output_lines, script_name)
            success_msg = get_success_message(script_name, processed_files, output_lines)
            
            if success_msg:
                flash(success_msg)
            else:
                if "No CSV or Excel files found" in result.stdout:
                    flash(f"No CSV or Excel files found in {input_dir} for {script_desc}.")
                elif "No valid files were processed" in result.stdout:
                    flash(f"No files processed for {script_desc}. Check input files for required columns in {input_dir}.")
                else:
                    flash(f"No files processed for {script_desc}. Check logs for details.")
                
        except subprocess.TimeoutExpired:
            flash(f"Script timed out but may still be running in background.")
        except subprocess.CalledProcessError as e:
            output_lines = e.stdout.splitlines() if e.stdout else []
            processed_files = count_processed_files(output_lines, script_name)
            success_msg = get_success_message(script_name, processed_files, output_lines)
            if success_msg:
                flash(f"Partially completed: {success_msg}, but encountered an error: {e.stderr or str(e)}")
            else:
                flash(f"Error processing {script_desc}: {e.stderr or str(e)}")
            
    except Exception as e:
        flash(f"Server error processing {script_desc}: {str(e)}")
    return redirect(url_for("dashboard"))

@app.route("/logout")
@login_required
def logout():
    session.pop("logged_in", None)
    flash("You have been logged out.")
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)