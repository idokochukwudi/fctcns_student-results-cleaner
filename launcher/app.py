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
    """Check for ND examination files in the actual directory structure - UPDATED to check RAW_RESULTS subdirectory"""
    if not os.path.isdir(input_dir):
        return False
    
    # Check for ND sets (ND-2024, ND-2025, etc.)
    nd_sets = []
    for item in os.listdir(input_dir):
        item_path = os.path.join(input_dir, item)
        if os.path.isdir(item_path) and item.startswith('ND-') and item != 'ND-COURSES':
            nd_sets.append(item)
    
    if not nd_sets:
        return False
    
    # Check each ND set for Excel files in the RAW_RESULTS subdirectory
    total_files_found = 0
    for nd_set in nd_sets:
        set_path = os.path.join(input_dir, nd_set)
        if not os.path.isdir(set_path):
            continue
        
        # Look for RAW_RESULTS subdirectory
        raw_results_path = os.path.join(set_path, "RAW_RESULTS")
        if not os.path.isdir(raw_results_path):
            continue
            
        # Look for Excel files in the RAW_RESULTS subdirectory
        excel_files = [f for f in os.listdir(raw_results_path) 
                     if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~')]
        
        total_files_found += len(excel_files)
    
    return total_files_found > 0

def check_putme_files(input_dir):
    """Check for PUTME examination files"""
    if not os.path.isdir(input_dir):
        return False
    
    # Check for Excel files in PUTME directory
    excel_files = [f for f in os.listdir(input_dir) 
                  if f.lower().endswith(('.xlsx', '.xls')) and 'PUTME' in f.upper()]
    
    # Also check for candidate batch files in RAW_CANDIDATE_BATCHES
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
    
    # Check for Excel files in internal exam directory
    excel_files = [f for f in os.listdir(input_dir) 
                  if f.lower().endswith(('.xlsx', '.xls')) and 'Set2025' in f]
    
    return len(excel_files) > 0

def check_caosce_files(input_dir):
    """Check for CAOSCE exam files"""
    if not os.path.isdir(input_dir):
        return False
    
    # Check for Excel files in CAOSCE directory
    excel_files = [f for f in os.listdir(input_dir) 
                  if f.lower().endswith(('.xlsx', '.xls')) and 'CAOSCE' in f.upper()]
    
    return len(excel_files) > 0

def check_split_files(input_dir):
    """Check for JAMB split files - FIXED VERSION to handle both CSV and Excel files"""
    if not os.path.isdir(input_dir):
        return False
    
    # Check for ANY CSV or Excel files in JAMB directory (not just those with 'JAMB' in name)
    valid_files = [f for f in os.listdir(input_dir) 
                  if f.lower().endswith(('.csv', '.xlsx', '.xls')) and not f.startswith('~')]
    
    return len(valid_files) > 0

def check_input_files(input_dir, script_name):
    """Check for input files based on script type"""
    if not os.path.isdir(input_dir):
        return False
    
    # Special handling for different script types
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
    
    # For other scripts, check for CSV/Excel files directly in the directory
    try:
        dir_contents = os.listdir(input_dir)
        valid_extensions = ('.csv', '.xlsx', '.xls')
        input_files = [f for f in dir_contents if f.lower().endswith(valid_extensions) and not f.startswith('~')]
        return len(input_files) > 0
    except Exception:
        return False

def count_processed_semesters(output_lines):
    """Count actual semesters processed - FIXED VERSION"""
    processed_semesters = set()
    
    for line in output_lines:
        # Look for semester processing indicators
        if "PROCESSING SEMESTER:" in line:
            # Extract semester name
            semester_match = re.search(r"PROCESSING SEMESTER:\s*(ND-[A-Za-z\-]+)", line)
            if semester_match:
                semester_name = semester_match.group(1)
                processed_semesters.add(semester_name)
                print(f"✅ Found processed semester: {semester_name}")
        
        # Also look for successful semester completion
        elif "✅ Successfully processed" in line and "ND-" in line:
            # Extract semester from context
            for word in line.split():
                if word.startswith("ND-"):
                    processed_semesters.add(word)
                    print(f"✅ Found completed semester: {word}")
    
    print(f"📊 Total unique semesters processed: {len(processed_semesters)}")
    if processed_semesters:
        print(f"📋 Semesters: {list(processed_semesters)}")
    
    return len(processed_semesters)

def get_success_message(script_name, processed_count, output_lines):
    """Generate appropriate success message based on script and output"""
    if processed_count == 0:
        return None
    
    if script_name == "clean":
        if any("✅ All processing completed successfully!" in line for line in output_lines):
            return f"Successfully processed internal examination results! Generated master file and individual cleaned files."
        else:
            return f"Processed {processed_count} internal examination file(s)."
    
    elif script_name == "exam_processor":
        # Check for upgrade information
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
        
        if any("✅ ND Examination Results Processing completed successfully" in line for line in output_lines):
            return f"ND Examination processing completed successfully! Processed {processed_count} semester(s).{upgrade_info}{upgrade_count}"
        elif any("✅ Processing complete" in line for line in output_lines):
            return f"ND Examination processing completed! Processed {processed_count} semester(s).{upgrade_info}{upgrade_count}"
        elif processed_count > 0:
            return f"Successfully processed {processed_count} semester(s).{upgrade_info}{upgrade_count}"
        else:
            return f"Script completed but no specific success indicators found.{upgrade_info}{upgrade_count}"
    
    elif script_name == "utme":
        if any("Processing completed successfully" in line for line in output_lines):
            return f"PUTME processing completed successfully! Processed {processed_count} batch file(s)."
        else:
            return f"Processed {processed_count} PUTME batch file(s)."
    
    elif script_name == "caosce":
        if any("Processed" in line for line in output_lines):
            return f"CAOSCE processing completed! Processed {processed_count} file(s)."
        else:
            return f"Processed {processed_count} CAOSCE file(s)."
    
    elif script_name == "split":
        if any("Saved processed file:" in line for line in output_lines):
            return f"JAMB name splitting completed! Processed {processed_count} file(s)."
        else:
            return f"Processed {processed_count} JAMB file(s)."
    
    else:
        return f"Successfully processed {processed_count} file(s)."

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

        # Ensure script is executable
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
        
        # Check for input files with proper directory structure handling
        if not check_input_files(input_dir, script_name):
            if script_name == "exam_processor":
                # Provide more detailed information about what's missing
                nd_sets_found = []
                if os.path.isdir(input_dir):
                    for item in os.listdir(input_dir):
                        item_path = os.path.join(input_dir, item)
                        if os.path.isdir(item_path) and item.startswith('ND-') and item != 'ND-COURSES':
                            # Check if this set has RAW_RESULTS subdirectory and Excel files
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
                # Provide more detailed info for JAMB database
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

        # Handle scripts that need forms
        if script_name in ["utme", "exam_processor"]:
            if request.method == "GET":
                if script_name == "utme":
                    return render_template("utme_form.html", college=COLLEGE, department=DEPARTMENT)
                elif script_name == "exam_processor":
                    # Get available ND sets for the form
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
                        processed_count = count_processed_semesters(output_lines)
                        success_msg = get_success_message(script_name, processed_count, output_lines)
                        
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
                        # Even if there's an error, check if any files were processed
                        output_lines = e.stdout.splitlines() if e.stdout else []
                        processed_count = count_processed_semesters(output_lines)
                        success_msg = get_success_message(script_name, processed_count, output_lines)
                        if success_msg:
                            flash(f"Partially completed: {success_msg}, but encountered an error: {e.stderr or str(e)}")
                        else:
                            flash(f"Error processing {script_desc}: {e.stderr or str(e)}")
                    return redirect(url_for("dashboard"))

                elif script_name == "exam_processor":
                    # Get form parameters for exam processor
                    selected_set = request.form.get("selected_set", "all")
                    processing_mode = request.form.get("processing_mode", "auto")
                    selected_semesters = request.form.getlist("semesters")
                    pass_threshold = request.form.get("pass_threshold", "50.0")
                    upgrade_threshold = request.form.get("upgrade_threshold", "0")  # Get upgrade threshold
                    generate_pdf = "generate_pdf" in request.form
                    track_withdrawn = "track_withdrawn" in request.form
                    
                    print(f"FORM DATA RECEIVED for exam_processor:")
                    print(f"   Selected Set: {selected_set}")
                    print(f"   Processing Mode: {processing_mode}") 
                    print(f"   Selected Semesters: {selected_semesters}")
                    print(f"   Pass Threshold: {pass_threshold}")
                    print(f"   Upgrade Threshold: '{upgrade_threshold}'")  # Show what we're getting
                    print(f"   Generate PDF: {generate_pdf}")
                    print(f"   Track Withdrawn: {track_withdrawn}")

                    # Verify script exists and is accessible
                    script_path = os.path.join(project_root, SCRIPT_MAP[script_name])
                    print(f"Script path: {script_path}")
                    
                    if not os.path.exists(script_path):
                        flash(f"Exam processor script not found at: {script_path}")
                        return redirect(url_for("dashboard"))
                    
                    if not os.access(script_path, os.R_OK):
                        flash(f"No read permission for script: {script_path}")
                        return redirect(url_for("dashboard"))

                    # Set environment variables for non-interactive mode
                    env = os.environ.copy()
                    env['SELECTED_SET'] = selected_set
                    env['PROCESSING_MODE'] = processing_mode
                    env['PASS_THRESHOLD'] = pass_threshold
                    
                    # FIXED: Properly handle upgrade threshold
                    if upgrade_threshold and upgrade_threshold.strip() and upgrade_threshold != "0":
                        env['UPGRADE_THRESHOLD'] = upgrade_threshold.strip()
                        print(f"   UPGRADE THRESHOLD SET: {upgrade_threshold}")
                    else:
                        # Don't set it at all if no upgrade threshold
                        if 'UPGRADE_THRESHOLD' in env:
                            del env['UPGRADE_THRESHOLD']
                        print(f"   NO UPGRADE THRESHOLD SET")

                    env['GENERATE_PDF'] = str(generate_pdf)
                    env['TRACK_WITHDRAWN'] = str(track_withdrawn)
                    
                    # FIXED: Handle semester selection properly
                    if processing_mode == "manual" and selected_semesters:
                        # Convert semester values to proper semester keys
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
                            else:
                                # If it's already a proper semester key, use it directly
                                if semester.startswith('ND-'):
                                    selected_semester_keys.append(semester)
                        
                        if selected_semester_keys:
                            env['SELECTED_SEMESTERS'] = ','.join(selected_semester_keys)
                            print(f"   SELECTED SEMESTERS SET: {selected_semester_keys}")
                        else:
                            env['SELECTED_SEMESTERS'] = ''
                            print(f"   NO SPECIFIC SEMESTERS SELECTED")
                    else:
                        # Auto mode or no specific semesters selected
                        env['SELECTED_SEMESTERS'] = ''
                        print(f"   AUTO MODE - PROCESSING ALL SEMESTERS")
                    
                    print(f"Environment variables set:")
                    for key in ['SELECTED_SET', 'PROCESSING_MODE', 'PASS_THRESHOLD', 'UPGRADE_THRESHOLD', 'GENERATE_PDF', 'TRACK_WITHDRAWN', 'SELECTED_SEMESTERS']:
                        print(f"   {key}: {env.get(key)}")

                    # Run the script with environment variables
                    print(f"Starting exam processor script...")
                    result = subprocess.run(
                        [sys.executable, script_path],
                        env=env,
                        text=True,
                        capture_output=True,
                        timeout=600  # 10 minutes
                    )
                    
                    print(f"Script execution completed")
                    print(f"   Return code: {result.returncode}")
                    print(f"   stdout length: {len(result.stdout)}")
                    print(f"   stderr length: {len(result.stderr)}")
                    
                    if result.stdout:
                        # Show first few lines for debugging
                        output_lines = result.stdout.splitlines()
                        for line in output_lines[:20]:
                            if line.strip():
                                print(f"STDOUT: {line}")
                    
                    if result.stderr:
                        print(f"STDERR: {result.stderr}")
                    
                    # Process results - FIXED: Use the new counting function
                    output_lines = result.stdout.splitlines() if result.stdout else []
                    processed_count = count_processed_semesters(output_lines)
                    
                    if result.returncode == 0:
                        # Check for upgrade information
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
                        
                        if any("✅ ND Examination Results Processing completed successfully" in line for line in output_lines):
                            flash(f"ND Examination processing completed successfully! Processed {processed_count} semester(s).{upgrade_details}{upgrade_count}")
                        elif any("✅ Processing complete" in line for line in output_lines):
                            flash(f"ND Examination processing completed! Processed {processed_count} semester(s).{upgrade_details}{upgrade_count}")
                        elif processed_count > 0:
                            flash(f"Successfully processed {processed_count} semester(s).{upgrade_details}{upgrade_count}")
                        else:
                            flash(f"Script completed but no specific success indicators found.{upgrade_details}{upgrade_count}")
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

        # Handle scripts that run directly (no form needed)
        try:
            result = subprocess.run(
                [sys.executable, script_path],
                text=True,
                capture_output=True,
                check=True,
                timeout=300
            )
            output_lines = result.stdout.splitlines()
            processed_count = count_processed_semesters(output_lines)
            success_msg = get_success_message(script_name, processed_count, output_lines)
            
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
            processed_count = count_processed_semesters(output_lines)
            success_msg = get_success_message(script_name, processed_count, output_lines)
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