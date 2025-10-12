import os
import subprocess
import re
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
        r"PROCESSING SEMESTER: (ND-.*)", 
        r"✅ Successfully processed .*", 
        r"✅ Mastersheet saved:.*",
        r"📁 Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"✅ Processing complete",
        r"✅ ND Examination Results Processing completed successfully"
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
    """Check for ND examination files in the actual directory structure - FIXED VERSION"""
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
    
    # Check each ND set for Excel files directly in the set directory
    total_files_found = 0
    for nd_set in nd_sets:
        set_path = os.path.join(input_dir, nd_set)
        if not os.path.isdir(set_path):
            continue
            
        # Look for Excel files directly in the set directory
        excel_files = [f for f in os.listdir(set_path) 
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

def count_processed_files(output_lines, script_name):
    """Count processed files based on script output"""
    success_indicators = SUCCESS_INDICATORS.get(script_name, [])
    processed_files_set = set()
    
    for line in output_lines:
        for indicator in success_indicators:
            match = re.search(indicator, line)
            if match:
                if script_name == "utme":
                    # For UTME, count unique file patterns
                    if "Processing:" in line:
                        file_name = match.group(1)
                        processed_files_set.add(f"Processed: {file_name}")
                    elif "Saved processed file:" in line:
                        file_name = match.group(1)
                        processed_files_set.add(f"Saved: {file_name}")
                elif script_name == "clean":
                    # For internal exam cleaning
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
                    # For exam processor
                    if "PROCESSING SEMESTER:" in line:
                        semester = match.group(1)
                        processed_files_set.add(f"Semester: {semester}")
                    elif "Processing:" in line:
                        file_name = match.group(1)
                        processed_files_set.add(f"Processed: {file_name}")
                    elif "✅ Successfully processed" in line:
                        processed_files_set.add("Successfully processed")
                    elif "📁 Found" in line:
                        # Extract number of files found
                        files_match = re.search(r"📁 Found (\d+) raw files", line)
                        if files_match:
                            processed_files_set.add(f"Files found: {files_match.group(1)}")
                    elif "✅ ND Examination Results Processing completed successfully" in line:
                        processed_files_set.add("Processing completed")
                else:
                    # For other scripts
                    file_name = match.group(1) if match.groups() else line
                    processed_files_set.add(file_name)
    
    return len(processed_files_set)

def get_success_message(script_name, processed_files, output_lines):
    """Generate appropriate success message based on script and output"""
    if processed_files == 0:
        return None
    
    if script_name == "clean":
        if any("✅ All processing completed successfully!" in line for line in output_lines):
            return f"Successfully processed internal examination results! Generated master file and individual cleaned files."
        else:
            return f"Processed {processed_files} internal examination file(s)."
    
    elif script_name == "exam_processor":
        if any("✅ ND Examination Results Processing completed successfully" in line for line in output_lines):
            return f"ND Examination processing completed successfully! Processed {processed_files} semester(s)/file(s)."
        elif any("✅ Processing complete" in line for line in output_lines):
            return f"ND Examination processing completed! Processed {processed_files} semester(s)/file(s)."
        else:
            return f"Processed {processed_files} ND examination file(s)/semester(s)."
    
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
                            # Check if this set has Excel files
                            excel_files = [f for f in os.listdir(item_path) 
                                         if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~')]
                            nd_sets_found.append(f"{item} ({len(excel_files)} files)")
                
                if not nd_sets_found:
                    flash(f"No ND examination sets found in {input_dir}. Please ensure you have ND sets (ND-2024, ND-2025, etc.) with Excel files.")
                else:
                    flash(f"ND sets found but no Excel files detected. Sets: {', '.join(nd_sets_found)}. Please ensure Excel files are in the set directories.")
                    
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
                        # Even if there's an error, check if any files were processed
                        output_lines = e.stdout.splitlines() if e.stdout else []
                        processed_files = count_processed_files(output_lines, script_name)
                        success_msg = get_success_message(script_name, processed_files, output_lines)
                        if success_msg:
                            flash(f"Partially completed: {success_msg}, but encountered an error: {e.stderr or str(e)}")
                        else:
                            flash(f"Error processing {script_desc}: {e.stderr or str(e)}")
                    return redirect(url_for("dashboard"))

                elif script_name == "exam_processor":
                    # Get form parameters for exam processor
                    selected_set = request.form.get("selected_set", "all")
                    
                    # Build command - use non-interactive mode with set selection
                    cmd = ["python3", script_path, "--non-interactive"]
                    
                    # Add set selection
                    if selected_set != "all":
                        cmd.extend(["--set", selected_set])
                    
                    try:
                        result = subprocess.run(
                            cmd,
                            text=True,
                            capture_output=True,
                            check=True,
                            timeout=600
                        )
                        
                        output_lines = result.stdout.splitlines()
                        processed_files = count_processed_files(output_lines, script_name)
                        
                        # Check for success indicators
                        if any("✅ ND Examination Results Processing completed successfully" in line for line in output_lines):
                            flash("ND Examination processing completed successfully!")
                        elif any("✅ Processing complete" in line for line in output_lines):
                            flash("ND Examination processing completed!")
                        elif processed_files > 0:
                            flash(f"Successfully processed {processed_files} semester(s)/file(s).")
                        elif result.returncode == 0:
                            flash("Exam processor completed successfully!")
                        else:
                            flash("Exam processor ran but may not have processed any files. Check the output for details.")
                            
                    except subprocess.TimeoutExpired:
                        flash(f"Script timed out after 10 minutes. The exam processor may still be running in background.")
                    except subprocess.CalledProcessError as e:
                        # Check if the error is due to unsupported parameters
                        if "unrecognized arguments" in e.stderr or "No such option" in e.stderr:
                            # Fallback: run without parameters
                            try:
                                fallback_cmd = ["python3", script_path]
                                fallback_result = subprocess.run(
                                    fallback_cmd,
                                    text=True,
                                    capture_output=True,
                                    check=True,
                                    timeout=600
                                )
                                flash("ND Examination processing completed (using default settings)!")
                            except subprocess.CalledProcessError as fallback_error:
                                flash(f"Error running exam processor: {fallback_error.stderr or str(fallback_error)}")
                        else:
                            flash(f"Error running exam processor: {e.stderr or str(e)}")
                    
                    return redirect(url_for("dashboard"))

        # Handle scripts that run directly (no form needed)
        try:
            result = subprocess.run(
                ["python3", script_path],
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