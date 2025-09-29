#!/usr/bin/env python3
"""
launcher.py

Centralized launcher for student result cleaning scripts.

Features:
- Password protection using .env file (hidden input)
- Auto-create required folders in Windows Documents
- Prompt to select which script to run
- Works with WSL and Python 3
- User-friendly prompts and messages
"""

import os
import subprocess
import sys
from dotenv import load_dotenv
from getpass import getpass  # Hide password input

# ---------------------------
# ANSI color codes (for console)
# ---------------------------
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
RESET = "\033[0m"

# ---------------------------
# Load environment variables
# ---------------------------
BASE_DIR = os.path.expanduser("~/student_result_cleaner")
dotenv_path = os.path.join(BASE_DIR, ".env")
load_dotenv(dotenv_path)
PASSWORD = os.environ.get("STUDENT_CLEANER_PASSWORD")

if not PASSWORD:
    print(f"{RED}❌ .env file missing or STUDENT_CLEANER_PASSWORD not set.{RESET}")
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
# Windows username detection
# ---------------------------
WINDOWS_USER = "MTECH COMPUTERS"  # Replace if your Windows username changes

# ---------------------------
# Define folders
# ---------------------------
FOLDERS = {
    "CAOSCE_RAW": f"/mnt/c/Users/{WINDOWS_USER}/Documents/CAOSCE_RAW",
    "CAOSCE_CLEAN": f"/mnt/c/Users/{WINDOWS_USER}/Documents/CAOSCE_CLEAN",
    "RAW_RESULTS": f"/mnt/c/Users/{WINDOWS_USER}/Documents/RAW_RESULTS",
    "CLEANED_RESULTS": f"/mnt/c/Users/{WINDOWS_USER}/Documents/CLEANED_RESULTS",
    "RAW_JAMB_DB": f"/mnt/c/Users/{WINDOWS_USER}/Documents/RAW_JAMB_DB",
    "CLEAN_JAMB_DB": f"/mnt/c/Users/{WINDOWS_USER}/Documents/CLEAN_JAMB_DB"
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
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
SCRIPTS = {
    "1": ("CAOSCE Result Cleaning", os.path.join(SCRIPTS_DIR, "caosce_result.py")),
    "2": ("Internal Exam Cleaning (UTME/clean_results)", os.path.join(SCRIPTS_DIR, "clean_results.py")),
    "3": ("UTME Result Cleaning (utme_result)", os.path.join(SCRIPTS_DIR, "utme_result.py")),
    "4": ("JAMB Name Split (split_names)", os.path.join(SCRIPTS_DIR, "split_names.py"))
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
# Run selected script
# ---------------------------
print(f"\n{YELLOW}🚀 Running {script_name} ...{RESET}\n")
try:
    subprocess.run(["python3", script_to_run], check=True)
    print(f"\n{GREEN}✅ {script_name} completed successfully!{RESET}")
except subprocess.CalledProcessError as e:
    print(f"\n{RED}❌ An error occurred while running {script_name}.{RESET}")
    print(str(e))

input("\nPress any key to exit . . .")
