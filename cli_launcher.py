#!/usr/bin/env python3
"""
cli_launcher.py

Centralized launcher for student result cleaning scripts.
FIXED VERSION - Enhanced support for all programs and better error handling.
"""

import os
import subprocess
import sys
from getpass import getpass  # Hide password input

# ---------------------------
# ANSI color codes (for console)
# ---------------------------
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
BLUE = "\033[94m"
RESET = "\033[0m"

# ---------------------------
# CORRECTED: Actual data directory (from your tree structure)
# ---------------------------
ACTUAL_DATA_DIR = os.path.expanduser("~/student_result_cleaner/EXAMS_INTERNAL")

# ---------------------------
# Base directory inside Windows Documents (for new uploads/processing)
# ---------------------------
BASE_DIR = os.path.expanduser("~/Documents/PROCESS_RESULT")

# ---------------------------
# Load environment variables from .env
# ---------------------------
try:
    from dotenv import load_dotenv, find_dotenv
except ModuleNotFoundError:
    print(
        f"{RED}❌ python-dotenv is not installed. Please install it in your venv:{RESET}"
    )
    print(f"{YELLOW}pip install python-dotenv{RESET}")
    sys.exit(1)

# Try multiple locations for .env file
dotenv_locations = [
    os.path.join(BASE_DIR, ".env"),
    os.path.join(ACTUAL_DATA_DIR, ".env"),
    os.path.join(os.path.dirname(__file__), ".env"),
]

dotenv_path = None
for location in dotenv_locations:
    if os.path.exists(location):
        dotenv_path = location
        break

if not dotenv_path:
    print(f"{RED}❌ .env file not found in any of these locations:{RESET}")
    for location in dotenv_locations:
        print(f"  - {location}")
    input("Press any key to exit . . .")
    sys.exit(1)

load_dotenv(dotenv_path)
PASSWORD = os.environ.get("STUDENT_CLEANER_PASSWORD")

if not PASSWORD:
    print(f"{RED}❌ STUDENT_CLEANER_PASSWORD not set in .env.{RESET}")
    input("Press any key to exit . . .")
    sys.exit(1)

# ---------------------------
# Password prompt
# ---------------------------
for attempt in range(3):
    entered = getpass("Enter the launcher password: ")
    if entered == PASSWORD:
        print(f"{GREEN}✅ Password correct!{RESET}\n")
        break
    else:
        print(
            f"{RED}❌ Incorrect password. Try again ({3 - attempt - 1} attempts left).{RESET}"
        )
else:
    print(f"{RED}❌ Too many failed attempts. Exiting.{RESET}")
    sys.exit(1)

# ---------------------------
# Define folders inside PROCESS_RESULT (for new processing)
# ---------------------------
FOLDERS = {
    "INTERNAL_RAW": os.path.join(BASE_DIR, "INTERNAL_RESULT/RAW_INTERNAL_RESULT"),
    "INTERNAL_CLEAN": os.path.join(BASE_DIR, "INTERNAL_RESULT/CLEAN_INTERNAL_RESULT"),
    "CAOSCE_RAW": os.path.join(BASE_DIR, "CAOSCE_RESULT/RAW_CAOSCE_RESULT"),
    "CAOSCE_CLEAN": os.path.join(BASE_DIR, "CAOSCE_RESULT/CLEAN_CAOSCE_RESULT"),
    "PUTME_RAW": os.path.join(BASE_DIR, "PUTME_RESULT/RAW_PUTME_RESULT"),
    "PUTME_RAW_BATCH": os.path.join(BASE_DIR, "PUTME_RESULT/RAW_CANDIDATE_BATCHES"),
    "PUTME_CLEAN": os.path.join(BASE_DIR, "PUTME_RESULT/CLEAN_PUTME_RESULT"),
    "JAMB_RAW": os.path.join(BASE_DIR, "JAMB_DB/RAW_JAMB_DB"),
    "JAMB_CLEAN": os.path.join(BASE_DIR, "JAMB_DB/CLEAN_JAMB_DB"),
    "SET48_RAW": os.path.join(BASE_DIR, "SET48_RESULTS/RAW_RESULTS"),
    "SET48_CLEAN": os.path.join(BASE_DIR, "SET48_RESULTS/CLEAN_RESULTS"),
    "ND_COURSES": os.path.join(BASE_DIR, "EXAMS_INTERNAL/ND-COURSES"),
    "BN_COURSES": os.path.join(BASE_DIR, "EXAMS_INTERNAL/BN-COURSES"),
    "BM_COURSES": os.path.join(BASE_DIR, "EXAMS_INTERNAL/BM-COURSES"),
}

# ---------------------------
# Auto-create folders (for new processing)
# ---------------------------
print(f"{YELLOW}🔹 Setting up required folders...{RESET}")
for name, path in FOLDERS.items():
    os.makedirs(path, exist_ok=True)
    print(f"{GREEN}✅ Folder ready: {path}{RESET}")

# ---------------------------
# Scripts paths - FIXED: Added BN and BM processors
# ---------------------------
SCRIPTS_DIR = os.path.join(os.path.dirname(__file__), "scripts")
SCRIPTS = {
    "1": (
        "CAOSCE Result Cleaning          ",
        os.path.join(SCRIPTS_DIR, "caosce_result.py"),
    ),
    "2": (
        "Internal Exam Cleaning        ",
        os.path.join(SCRIPTS_DIR, "clean_results.py"),
    ),
    "3": (
        "PUTME Result Cleaning         ",
        os.path.join(SCRIPTS_DIR, "utme_result.py"),
    ),
    "4": (
        "JAMB Candidate Name Split     ",
        os.path.join(SCRIPTS_DIR, "split_names.py"),
    ),
    "5": (
        "ND Examination Results        ",
        os.path.join(SCRIPTS_DIR, "exam_result_processor.py"),
    ),
    "6": (
        "Basic Nursing Results         ",
        os.path.join(SCRIPTS_DIR, "exam_result_processor.py"),  # Same processor, different program
    ),
    "7": (
        "Basic Midwifery Results       ",
        os.path.join(SCRIPTS_DIR, "exam_result_processor.py"),  # Same processor, different program
    ),
}

# ---------------------------
# Menu prompt with aligned descriptions
# ---------------------------
print(f"\n{BLUE}🎯 SELECT SCRIPT TO RUN:{RESET}")
max_desc_len = max(len(desc) for desc, _ in SCRIPTS.values())
for key, (desc, _) in SCRIPTS.items():
    print(f"{key}. {desc:<{max_desc_len}}")

while True:
    choice = input("\nEnter 1, 2, 3, 4, 5, 6, or 7: ").strip()
    if choice in SCRIPTS:
        script_name, script_to_run = SCRIPTS[choice]
        if not os.path.exists(script_to_run):
            print(f"{RED}❌ Script not found: {script_to_run}{RESET}")
            input("Press any key to exit . . .")
            sys.exit(1)
        break
    else:
        print(f"{RED}❌ Invalid choice, please enter 1-7.{RESET}")

# ---------------------------
# Confirm script execution
# ---------------------------
print(f"\n{YELLOW}🔹 You selected: {script_name.strip()}{RESET}")
confirm = input("Proceed with running this script? (y/n): ").strip().lower()
if confirm != "y":
    print(f"{YELLOW}⚠️ Script execution cancelled.{RESET}")
    input("Press any key to exit . . .")
    sys.exit(0)

# ---------------------------
# Special handling for exam processors
# ---------------------------
if choice in ["5", "6", "7"]:  # ND, Basic Nursing, Basic Midwifery
    program = {"5": "ND", "6": "BN", "7": "BM"}[choice]
    print(f"\n{BLUE}🎓 {program} EXAMINATION PROCESSOR SETUP{RESET}")
    print(f"{YELLOW}📚 This script includes FLEXIBLE UPGRADE RULE{RESET}")
    print(
        f"{YELLOW}🔹 You'll be prompted for each semester to choose score upgrades{RESET}"
    )
    print(
        f"{YELLOW}🔹 Options: 45, 46, 47, 48, 49 (upgrade range to 50) or 0 to skip{RESET}"
    )
    print(f"{YELLOW}🔹 Example: Enter '47' to upgrade scores 47-49 to 50{RESET}")

    print(f"\n{BLUE}🚀 Starting {program} Examination Results Processor...{RESET}")
    print(
        f"{YELLOW}Note: Follow the interactive prompts for set selection and semester processing.{RESET}"
    )
    print(f"{YELLOW}Using data directory: {ACTUAL_DATA_DIR}{RESET}\n")

# ---------------------------
# Run selected script using current venv Python
# CRITICAL FIX: Remove capture_output=True to allow interactive I/O
# ---------------------------
print(f"\n{YELLOW}🚀 Running {script_name.strip()} ...{RESET}\n")
try:
    # Set environment for the subprocess
    env = os.environ.copy()

    # For exam processors, use ACTUAL_DATA_DIR as BASE_DIR and set SELECTED_PROGRAM
    if choice in ["5", "6", "7"]:
        program = {"5": "ND", "6": "BN", "7": "BM"}[choice]
        env["BASE_DIR"] = ACTUAL_DATA_DIR
        env["SELECTED_PROGRAM"] = program
        env["PASS_THRESHOLD"] = "50.0"
        print(f"{BLUE}🔹 Processing {program} examination results{RESET}")
    else:
        # For other scripts, use BASE_DIR (for new uploads/processing)
        env["BASE_DIR"] = BASE_DIR

    # CRITICAL FIX: Don't capture output for interactive scripts
    # This allows the script to prompt for user input and display output in real-time
    result = subprocess.run(
        [sys.executable, script_to_run],
        env=env,
        check=False,  # Changed from check=True to allow custom error handling
        text=True
        # REMOVED: capture_output=True - this was blocking interactive I/O
        # REMOVED: timeout parameter - let it run as long as needed
    )

    print(f"\n{GREEN}{'='*60}{RESET}")
    
    # Check return code
    if result.returncode == 0:
        print(f"{GREEN}✅ {script_name.strip()} completed successfully!{RESET}")
        
        # Special success message for exam processors
        if choice in ["5", "6", "7"]:
            program = {"5": "ND", "6": "BN", "7": "BM"}[choice]
            print(f"{GREEN}🎉 {program} Examination processing finished!{RESET}")
            print(
                f"{YELLOW}📊 Check the CLEAN_RESULTS folder for mastersheets and PDFs{RESET}"
            )
            print(f"{YELLOW}📁 Location: {ACTUAL_DATA_DIR}/{{SET-NAME}}/CLEAN_RESULTS{RESET}")
    else:
        print(f"\n{RED}❌ Script exited with error code: {result.returncode}{RESET}")
        
        if choice in ["5", "6", "7"]:
            program = {"5": "ND", "6": "BN", "7": "BM"}[choice]
            print(f"{YELLOW}Note for {program} Exam Processor:{RESET}")
            print(
                f"{YELLOW}• Ensure 'course-code-creditUnit.xlsx' exists in {ACTUAL_DATA_DIR}/{program}-COURSES folder{RESET}"
            )
            print(f"{YELLOW}• Check that RAW_RESULTS folders contain Excel files{RESET}")
            print(f"{YELLOW}• Verify semester files follow naming conventions{RESET}")

            # Check if course file exists
            course_file = os.path.join(
                ACTUAL_DATA_DIR, f"{program}-COURSES", "course-code-creditUnit.xlsx"
            )
            if not os.path.exists(course_file):
                print(f"{RED}❌ Course file not found: {course_file}{RESET}")
            else:
                print(f"{GREEN}✅ Course file found: {course_file}{RESET}")

            # Check for program sets
            program_sets = []
            if os.path.exists(ACTUAL_DATA_DIR):
                for item in os.listdir(ACTUAL_DATA_DIR):
                    item_path = os.path.join(ACTUAL_DATA_DIR, item)
                    if os.path.isdir(item_path) and item.startswith(f"{program}-"):
                        program_sets.append(item)

            if program_sets:
                print(f"{GREEN}✅ Found {program} sets: {', '.join(program_sets)}{RESET}")
            else:
                print(f"{RED}❌ No {program} sets found in {ACTUAL_DATA_DIR}{RESET}")

        else:
            print(f"{YELLOW}Note: Check input files and folder structure.{RESET}")

except KeyboardInterrupt:
    print(f"\n{YELLOW}⚠️ Script execution interrupted by user.{RESET}")

except Exception as e:
    print(f"\n{RED}❌ Unexpected error: {e}{RESET}")
    import traceback
    traceback.print_exc()

print(f"\n{GREEN}{'='*60}{RESET}")
input("\nPress any key to exit . . .")