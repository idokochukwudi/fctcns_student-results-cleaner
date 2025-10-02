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
    "JAMB_CLEAN": os.path.join(BASE_DIR, "JAMB_DB/CLEAN_JAMB_DB")
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
    "4": ("JAMB Candidate Name Split", os.path.join(SCRIPTS_DIR, "split_names.py"))
}

# ---------------------------
# Menu prompt
# ---------------------------
print("\nSelect the script to run:")
for key, (desc, _) in SCRIPTS.items():
    print(f"{key}. {desc}")

while True:
    choice = input("Enter 1, 2, 3, or 4: ").strip()
    if choice in SCRIPTS:
        script_name, script_to_run = SCRIPTS[choice]
        if not os.path.exists(script_to_run):
            print(f"{RED}❌ Script not found: {script_to_run}{RESET}")
            sys.exit(1)
        break
    else:
        print(f"{RED}❌ Invalid selection. Please enter 1, 2, 3, or 4.{RESET}")

# ---------------------------
# Run selected script using current venv Python
# ---------------------------
print(f"\n{YELLOW}🚀 Running {script_name} ...{RESET}\n")
try:
    subprocess.run([sys.executable, script_to_run], check=True)
    print(f"\n{GREEN}✅ {script_name} completed successfully!{RESET}")
except subprocess.CalledProcessError as e:
    print(f"\n{RED}❌ An error occurred while running {script_name}.{RESET}")
    print(str(e))

input("\nPress any key to exit . . .")
