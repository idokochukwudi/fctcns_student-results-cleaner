#!/usr/bin/env python3
"""
launcher.py

Centralized launcher for student result cleaning scripts.

Features:
- Password protection using .env file (hidden input)
- Auto-create required folders in Windows Documents/PROCESS_RESULT
- Prompt to select which script to run
- Works with WSL and Python 3
- User-friendly prompts and messages
- Enhanced support for interactive exam processor with upgrade prompts
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
# Base directory inside Windows Documents
# ---------------------------
BASE_DIR = os.path.expanduser("~/Documents/PROCESS_RESULT")

# ---------------------------
# Load environment variables from .env
# ---------------------------
try:
    from dotenv import load_dotenv, find_dotenv
except ModuleNotFoundError:
    print(f"{RED}❌ python-dotenv is not installed. Please install it in your venv:{RESET}")
    print(f"{YELLOW}pip install python-dotenv{RESET}")
    sys.exit(1)

dotenv_path = os.path.join(BASE_DIR, ".env")
if not os.path.exists(dotenv_path):
    # fallback to launcher folder if BASE_DIR .env not found
    dotenv_path = os.path.join(os.path.dirname(__file__), ".env")

if not os.path.exists(dotenv_path):
    print(f"{RED}❌ .env file not found in {BASE_DIR} or launcher directory.{RESET}")
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
        print(f"{RED}❌ Incorrect password. Try again.{RESET}")
else:
    print(f"{RED}❌ Too many failed attempts. Exiting.{RESET}")
    sys.exit(1)

# ---------------------------
# Define folders inside PROCESS_RESULT
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
    "ND_COURSES": os.path.join(BASE_DIR, "EXAMS_INTERNAL/ND-COURSES")
}

# ---------------------------
# Auto-create folders
# ---------------------------
print(f"{YELLOW}🔹 Setting up required folders...{RESET}")
for name, path in FOLDERS.items():
    os.makedirs(path, exist_ok=True)
    print(f"{GREEN}✅ Folder ready: {path}{RESET}")

# ---------------------------
# Scripts paths
# ---------------------------
SCRIPTS_DIR = os.path.join(os.path.dirname(__file__), "scripts")
SCRIPTS = {
    "1": ("CAOSCE Result Cleaning", os.path.join(SCRIPTS_DIR, "caosce_result.py")),
    "2": ("Internal Exam Cleaning", os.path.join(SCRIPTS_DIR, "clean_results.py")),
    "3": ("PUTME Result Cleaning", os.path.join(SCRIPTS_DIR, "utme_result.py")),
    "4": ("JAMB Candidate Name Split", os.path.join(SCRIPTS_DIR, "split_names.py")),
    "5": ("ND Examination Results Processing", os.path.join(SCRIPTS_DIR, "exam_result_processor.py"))
}

# ---------------------------
# Menu prompt
# ---------------------------
print(f"\n{BLUE}🎯 SELECT SCRIPT TO RUN:{RESET}")
for key, (desc, _) in SCRIPTS.items():
    print(f"{key}. {desc}")

while True:
    choice = input("\nEnter 1, 2, 3, 4, or 5: ").strip()
    if choice in SCRIPTS:
        script_name, script_to_run = SCRIPTS[choice]
        if not os.path.exists(script_to_run):
            print(f"{RED}❌ Script not found: {script_to_run}{RESET}")
            sys.exit(1)
        break
    else:
        print(f"{RED}❌ Invalid selection. Please enter 1, 2, 3, 4, or 5.{RESET}")

# ---------------------------
# Special handling for exam processor
# ---------------------------
if choice == "5":  # ND Examination Results Processing
    print(f"\n{BLUE}🎓 ND EXAMINATION PROCESSOR SETUP{RESET}")
    print(f"{YELLOW}📚 This script now includes FLEXIBLE UPGRADE RULE{RESET}")
    print(f"{YELLOW}🔹 You'll be prompted for each semester to choose score upgrades{RESET}")
    print(f"{YELLOW}🔹 Options: 45, 46, 47, 48, 49 (upgrade range to 50) or 0 to skip{RESET}")
    print(f"{YELLOW}🔹 Example: Enter '47' to upgrade scores 47-49 to 50{RESET}")
    
    # Set default pass threshold via environment
    os.environ['PASS_THRESHOLD'] = '50.0'
    
    print(f"\n{BLUE}🚀 Starting ND Examination Results Processor...{RESET}")
    print(f"{YELLOW}Note: Follow the interactive prompts for set selection and semester processing.{RESET}\n")

# ---------------------------
# Run selected script using current venv Python
# ---------------------------
print(f"\n{YELLOW}🚀 Running {script_name} ...{RESET}\n")
try:
    if choice == "5":
        # For exam processor, run interactively (no timeout)
        result = subprocess.run([sys.executable, script_to_run], check=True)
    else:
        # For other scripts, use standard execution
        result = subprocess.run([sys.executable, script_to_run], check=True)
    
    print(f"\n{GREEN}✅ {script_name} completed successfully!{RESET}")
    
    # Special success message for exam processor
    if choice == "5":
        print(f"{GREEN}🎉 ND Examination processing finished!{RESET}")
        print(f"{YELLOW}📊 Check the CLEAN_RESULTS folder for mastersheets and PDFs{RESET}")
        
except subprocess.CalledProcessError as e:
    print(f"\n{RED}❌ An error occurred while running {script_name}.{RESET}")
    print(f"Command {e.cmd} returned non-zero exit status {e.returncode}.")
    
    if choice == "5":
        print(f"{YELLOW}Note for ND Processor:{RESET}")
        print(f"{YELLOW}• Ensure 'course-code-creditUnit.xlsx' exists in ND-COURSES folder{RESET}")
        print(f"{YELLOW}• Check that RAW_RESULTS folders contain Excel files{RESET}")
        print(f"{YELLOW}• Verify semester files follow naming conventions{RESET}")
    else:
        print(f"{YELLOW}Note: Check input files and folder structure.{RESET}")

except KeyboardInterrupt:
    print(f"\n{YELLOW}⚠️  Script execution interrupted by user.{RESET}")
    
except Exception as e:
    print(f"\n{RED}❌ Unexpected error: {e}{RESET}")

input("\nPress any key to exit . . .")