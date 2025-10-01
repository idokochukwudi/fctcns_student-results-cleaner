import os
import subprocess
import traceback
from flask import Flask, render_template, request, redirect, url_for, flash, session
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

COLLEGE = os.getenv("COLLEGE_NAME", "Default College")
DEPARTMENT = os.getenv("DEPARTMENT", "Default Department")
APP_PASSWORD = os.getenv("STUDENT_CLEANER_PASSWORD", "admin")
FLASK_SECRET = os.getenv("FLASK_SECRET", "supersecretkey")

app = Flask(__name__)
app.secret_key = FLASK_SECRET

# Script paths (absolute)
SCRIPTS_DIR = "/home/ernest/student_result_cleaner/scripts"

SCRIPT_MAP = {
    "utme": os.path.join(SCRIPTS_DIR, "utme_result.py"),
    "caosce": os.path.join(SCRIPTS_DIR, "caosce_result.py"),
    "clean": os.path.join(SCRIPTS_DIR, "clean_results.py"),
    "split": os.path.join(SCRIPTS_DIR, "split_names.py"),
}

# Success indicators for counting processed input files
SUCCESS_INDICATORS = {
    "utme": ["Processing: "],  # Count input files processed
    "caosce": ["Processed "],  # Count input files processed
    "clean": ["Cleaned CSV saved in WSL", "Master CSV saved in WSL"],
    "split": ["Saved processed file"]
}

# Login + Auth
@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        password = request.form.get("password")
        if password == APP_PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("dashboard"))
        else:
            error = "Invalid password. Please try again."
    return render_template(
        "login.html",
        college=COLLEGE,
        department=DEPARTMENT,
        error=error
    )

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.before_request
def require_login():
    allowed_routes = {"login", "static"}
    if "logged_in" not in session and request.endpoint not in allowed_routes:
        return redirect(url_for("login"))

# Dashboard
@app.route("/")
def dashboard():
    scripts = {
        "utme": {"desc": "Process PUTME Examination Results"},
        "caosce": {"desc": "Process CAOSCE Examination Results"},
        "clean": {"desc": "Process Objective Examination Results"},
        "split": {"desc": "Process JAMB Candidate Database"},
    }
    return render_template(
        "dashboard.html",
        college=COLLEGE,
        department=DEPARTMENT,
        scripts=scripts
    )

# Run Scripts
@app.route("/run/<script_name>", methods=["GET", "POST"])
def run_script(script_name):
    if script_name not in SCRIPT_MAP:
        flash("Invalid script requested.")
        return redirect(url_for("dashboard"))

    script_path = SCRIPT_MAP[script_name]
    script_desc = {
        "utme": "PUTME Examination Results",
        "caosce": "CAOSCE Examination Results",
        "clean": "Objective Examination Results",
        "split": "JAMB Candidate Database"
    }.get(script_name, "Script")

    # Verify script exists
    if not os.path.isfile(script_path):
        flash(f"Script file not found: {script_path}")
        return redirect(url_for("dashboard"))

    # Define input directory for error messaging
    input_dir = {
        "utme": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/PUTME_RESULT/RAW_PUTME_RESULT",
        "caosce": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/CAOSCE_RESULT/RAW_CAOSCE_RESULT",
        "clean": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/INTERNAL_RESULT/RAW_INTERNAL_RESULT",
        "split": "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/JAMB_DB/RAW_JAMB_DB"
    }.get(script_name, "the input directory")

    try:
        if script_name == "utme":
            if request.method == "GET":
                return render_template(
                    "utme_form.html",
                    college=COLLEGE,
                    department=DEPARTMENT
                )

            if request.method == "POST":
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
                    check=True
                )

                # Parse output for summary
                output_lines = result.stdout.splitlines()
                success_indicators = SUCCESS_INDICATORS.get(script_name, ["Processing: "])
                processed_files = sum(1 for line in output_lines if any(indicator in line for indicator in success_indicators))
                skipped_files = sum(1 for line in output_lines if "Skipping" in line)
                # Log raw output for debugging
                print(f"Script {script_name} stdout: {result.stdout}")
                print(f"Script {script_name} stderr: {result.stderr}")
                if processed_files == 0 and skipped_files > 0:
                    flash(f"No files were processed for {script_desc}. {skipped_files} file(s) skipped due to missing required columns or invalid format.")
                else:
                    flash(f"Successfully processed {processed_files} input file(s) for {script_desc}. {skipped_files} file(s) skipped due to missing required columns or invalid format.")
                return redirect(url_for("dashboard"))

        # Other scripts
        result = subprocess.run(
            ["python3", script_path],
            text=True,
            capture_output=True,
            check=True
        )

        # Parse output for summary
        output_lines = result.stdout.splitlines()
        success_indicators = SUCCESS_INDICATORS.get(script_name, ["Saved processed file"])
        processed_files = sum(1 for line in output_lines if any(indicator in line for indicator in success_indicators))
        skipped_files = sum(1 for line in output_lines if "Skipping" in line)
        # Log raw output for debugging
        print(f"Script {script_name} stdout: {result.stdout}")
        print(f"Script {script_name} stderr: {result.stderr}")
        if "No CSV or Excel files found" in result.stdout or "No CSV or Excel files found" in result.stderr:
            flash(f"No files found to process for {script_desc}. Please add CSV or Excel files to the input directory: {input_dir}")
        elif "No valid files were processed" in result.stdout or "No valid files were processed" in result.stderr:
            flash(f"No files were processed for {script_desc}. Please check input files for required columns (e.g., Surname, First name, Grade/...).")
        elif processed_files == 0:
            flash(f"No files were processed for {script_desc}. {skipped_files} file(s) skipped due to missing required columns or invalid format.")
        else:
            flash(f"Successfully processed {processed_files} input file(s) for {script_desc}. {skipped_files} file(s) skipped due to missing required columns or invalid format.")
    except subprocess.CalledProcessError as e:
        print(f"Script {script_name} stdout: {e.stdout}")
        print(f"Script {script_name} stderr: {e.stderr}")
        if "No CSV or Excel files found" in e.stdout or "No CSV or Excel files found" in e.stderr:
            flash(f"No files found to process for {script_desc}. Please add CSV or Excel files to the input directory: {input_dir}")
        elif "No valid files were processed" in e.stdout or "No valid files were processed" in e.stderr:
            flash(f"No files were processed for {script_desc}. Please check input files for required columns (e.g., Surname, First name, Grade/...).")
        else:
            flash(f"Error processing {script_desc}: {e.stderr or str(e)}")
    except Exception as e:
        traceback.print_exc()
        flash(f"Unexpected error while processing {script_desc}: {str(e)}")

    return redirect(url_for("dashboard"))

# Start App
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)