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
app.logger.info(f"Loaded environment variables - STUDENT_CLEANER_PASSWORD: {os.getenv('STUDENT_CLEANER_PASSWORD', 'Not found')}, COLLEGE_NAME: {os.getenv('COLLEGE_NAME', 'Not found')}, DEPARTMENT: {os.getenv('DEPARTMENT', 'Not found')}, FLASK_SECRET: {os.getenv('FLASK_SECRET', 'Not found')}")

PASSWORD = os.getenv("STUDENT_CLEANER_PASSWORD", "admin")
COLLEGE = os.getenv("COLLEGE_NAME", "FCT College of Nursing Sciences, Gwagwalada")
DEPARTMENT = os.getenv("DEPARTMENT", "Examinations Office")

SCRIPT_MAP = {
    "utme": "scripts/utme_result.py",
    "caosce": "scripts/caosce_result.py",
    "clean": "scripts/clean_results.py",
    "split": "scripts/split_names.py"
}
SUCCESS_INDICATORS = {
    "utme": [r"Processing: (2025-PBN-EEXAM-PBN-Batch\d+ Quiz-grades\.xlsx)"],
    "caosce": [r"Processed (CAOSCE SET2023A.*?|VIVA \([0-9]+\)\.xlsx) \(\d+ rows read\)"],
    "clean": [r"Cleaned CSV saved in Windows Documents: .*?cleaned_(Set2023A Class-[^\.]+)\.csv"],
    "split": [r"Saved processed file: (clean_jamb_DB_.*?\.csv)"]
}

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        app.logger.debug(f"Session data: {session}")
        if "logged_in" not in session:
            app.logger.warning("Session 'logged_in' not found, redirecting to login")
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
        app.logger.info(f"Attempted login with password: {password}, Expected: {PASSWORD}")
        if password == PASSWORD:
            session["logged_in"] = True
            app.logger.info(f"Session set: {session}")
            flash("Successfully logged in!")
            return redirect(url_for("dashboard"))
        else:
            app.logger.error(f"Login failed. Provided: {password}, Expected: {PASSWORD}")
            flash("Invalid password. Please try again.")
            return redirect(url_for("login"))
    return render_template("login.html", college=COLLEGE)

@app.route("/dashboard")
@login_required
def dashboard():
    try:
        return render_template("dashboard.html", college=COLLEGE, DEPARTMENT=DEPARTMENT)
    except TemplateNotFound as e:
        app.logger.error(f"Template not found: {str(e)}")
        flash(f"Template error: {str(e)}")
        return redirect(url_for("login"))
    except UndefinedError as e:
        app.logger.error(f"Template rendering error: {str(e)}")
        flash(f"Template rendering error: {str(e)}")
        return redirect(url_for("login"))
    except Exception as e:
        app.logger.error(f"Unexpected error in dashboard: {str(e)}")
        flash(f"Server error: {str(e)}")
        return redirect(url_for("login"))

@app.route("/run/<script_name>", methods=["GET", "POST"])
@login_required
def run_script(script_name):
    try:
        if script_name not in SCRIPT_MAP:
            app.logger.error(f"Invalid script requested: {script_name}")
            flash("Invalid script requested.")
            return redirect(url_for("dashboard"))

        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        script_path = os.path.join(project_root, SCRIPT_MAP[script_name])
        script_desc = {
            "utme": "PUTME Examination Results",
            "caosce": "CAOSCE Examination Results",
            "clean": "Objective Examination Results",
            "split": "JAMB Candidate Database"
        }.get(script_name, "Script")

        app.logger.debug(f"Checking script path: {script_path}")
        if not os.path.isfile(script_path):
            app.logger.error(f"Script file not found: {script_path}")
            flash(f"Script file not found: {script_path}")
            return redirect(url_for("dashboard"))

        if not os.access(script_path, os.X_OK):
            app.logger.warning(f"Script {script_path} is not executable. Attempting to fix permissions.")
            os.chmod(script_path, 0o755)

        input_dir = {
            "utme": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/PUTME_RESULT/RAW_PUTME_RESULT",
            "caosce": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/CAOSCE_RESULT/RAW_CAOSCE_RESULT",
            "clean": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/INTERNAL_RESULT/RAW_INTERNAL_RESULT",
            "split": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/JAMB_DB/RAW_JAMB_DB"
        }.get(script_name, "the input directory")

        app.logger.debug(f"Checking input directory: {input_dir}")
        if not os.path.isdir(input_dir):
            app.logger.error(f"Input directory not found: {input_dir}")
            flash(f"Input directory not found: {input_dir}")
            return redirect(url_for("dashboard"))

        try:
            dir_contents = os.listdir(input_dir)
            valid_extensions = ('.csv', '.xlsx', '.xls')
            input_files = [f for f in dir_contents if f.lower().endswith(valid_extensions)]
            app.logger.debug(f"Input directory contents: {dir_contents}")
            app.logger.debug(f"Valid input files: {input_files}")
            for item in input_files:
                item_path = os.path.join(input_dir, item)
                permissions = oct(os.stat(item_path).st_mode & 0o777)
                readable = os.access(item_path, os.R_OK)
                writable = os.access(item_path, os.W_OK)
                app.logger.debug(f"File: {item_path}, Permissions: {permissions}, Readable: {readable}, Writable: {writable}")
            if not input_files:
                app.logger.warning(f"No valid CSV/Excel files found in {input_dir}")
                flash(f"No CSV or Excel files found in {input_dir}")
                return redirect(url_for("dashboard"))
        except Exception as e:
            app.logger.error(f"Failed to list input directory {input_dir}: {str(e)}")
            flash(f"Cannot access input directory: {input_dir}")
            return redirect(url_for("dashboard"))

        if script_name == "utme":
            if request.method == "GET":
                try:
                    return render_template(
                        "utme_form.html",
                        college=COLLEGE,
                        DEPARTMENT=DEPARTMENT
                    )
                except TemplateNotFound as e:
                    app.logger.error(f"Template not found: {str(e)}")
                    flash(f"Template error: {str(e)}")
                    return redirect(url_for("dashboard"))
                except UndefinedError as e:
                    app.logger.error(f"Template rendering error: {str(e)}")
                    flash(f"Template rendering error: {str(e)}")
                    return redirect(url_for("dashboard"))

            if request.method == "POST":
                convert_value = request.form.get("convert_value", "").strip()
                convert_column = request.form.get("convert_column", "n")

                cmd = ["python3", script_path]
                if convert_value:
                    cmd.extend(["--non-interactive", "--converted-score-max", convert_value])

                app.logger.debug(f"Executing command: {cmd}")
                try:
                    result = subprocess.run(
                        cmd,
                        input=f"{convert_column}\n",
                        text=True,
                        capture_output=True,
                        check=True
                    )
                    output_lines = result.stdout.splitlines()
                    success_indicators = SUCCESS_INDICATORS.get(script_name, [r"Processing: (2025-PBN-EEXAM-PBN-Batch\d+ Quiz-grades\.xlsx)"])
                    processed_files_set = set()
                    for line in output_lines:
                        for indicator in success_indicators:
                            match = re.search(indicator, line)
                            if match:
                                file_name = match.group(1)
                                processed_files_set.add(file_name)
                                app.logger.debug(f"Matched file for {script_name}: {file_name} in line: {line}")
                    processed_files = len(processed_files_set)
                    skipped_files = sum(1 for line in output_lines if "Skipping" in line)
                    app.logger.info(f"Script {script_name} stdout: {result.stdout}")
                    app.logger.info(f"Script {script_name} stderr: {result.stderr}")
                    if processed_files == 0 and skipped_files > 0:
                        flash(f"No files processed for {script_desc}. {skipped_files} file(s) skipped.")
                    elif processed_files == 0:
                        flash(f"No files processed for {script_desc}. Check input files in {input_dir}.")
                    else:
                        flash(f"Successfully processed {processed_files} file(s) for {script_desc}. {skipped_files} file(s) skipped.")
                except subprocess.CalledProcessError as e:
                    app.logger.error(f"Subprocess error in {script_name}: {e.stderr}")
                    flash(f"Error processing {script_desc}: {e.stderr or str(e)}")
                return redirect(url_for("dashboard"))

        app.logger.debug(f"Executing command: python3 {script_path}")
        try:
            result = subprocess.run(
                ["python3", script_path],
                text=True,
                capture_output=True,
                check=True
            )
            output_lines = result.stdout.splitlines()
            success_indicators = SUCCESS_INDICATORS.get(script_name, [r"Saved processed file: .*?\.csv"])
            processed_files_set = set()
            for line in output_lines:
                for indicator in success_indicators:
                    match = re.search(indicator, line)
                    if match:
                        file_name = match.group(1)
                        processed_files_set.add(file_name)
                        app.logger.debug(f"Matched file for {script_name}: {file_name} in line: {line}")
            processed_files = len(processed_files_set)
            skipped_files = sum(1 for line in output_lines if "Skipping" in line)
            app.logger.info(f"Script {script_name} stdout: {result.stdout}")
            app.logger.info(f"Script {script_name} stderr: {result.stderr}")
            if "No CSV or Excel files found" in result.stdout or "No CSV or Excel files found" in result.stderr:
                flash(f"No CSV or Excel files found in {input_dir} for {script_desc}.")
            elif "No valid files were processed" in result.stdout or "No valid files were processed" in result.stderr:
                flash(f"No files processed for {script_desc}. Check input files for required columns in {input_dir}.")
            elif processed_files == 0:
                flash(f"No files processed for {script_desc}. {skipped_files} file(s) skipped. Check logs for details.")
            else:
                flash(f"Successfully processed {processed_files} file(s) for {script_desc}. {skipped_files} file(s) skipped.")
        except subprocess.CalledProcessError as e:
            app.logger.error(f"Subprocess error in {script_name}: {e.stderr}")
            flash(f"Error processing {script_desc}: {e.stderr or str(e)}")
    except Exception as e:
        app.logger.error(f"Unexpected error in {script_name}: {str(e)}")
        flash(f"Server error processing {script_desc}: {str(e)}")
    return redirect(url_for("dashboard"))

@app.route("/logout")
@login_required
def logout():
    session.pop("logged_in", None)
    app.logger.info(f"Session cleared: {session}")
    flash("You have been logged out.")
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)