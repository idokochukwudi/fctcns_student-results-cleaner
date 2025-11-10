import os
import subprocess
import re
import sys
import zipfile
import shutil
import time
import socket
import logging
import json
import glob
from pathlib import Path
from datetime import datetime
from flask import (
    Flask,
    request,
    redirect,
    url_for,
    render_template,
    flash,
    session,
    send_file,
    send_from_directory,
    jsonify,
)
from functools import wraps
from dotenv import load_dotenv
from jinja2 import TemplateNotFound
from werkzeug.utils import secure_filename
# Configure logging
logging.basicConfig(
    level=logging.DEBUG, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)
# Load environment variables
load_dotenv()
# Define directory structure relative to project root
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SCRIPT_DIR = os.path.join(PROJECT_ROOT, "scripts")
BASE_DIR = os.getenv("BASE_DIR", "/home/ernest/student_result_cleaner/EXAMS_INTERNAL")
# Launcher-specific directories
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "templates")
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
# Log paths for verification
logger.info(f"PROJECT_ROOT: {PROJECT_ROOT}")
logger.info(f"SCRIPT_DIR: {SCRIPT_DIR}")
logger.info(f"BASE_DIR (EXAMS_INTERNAL): {BASE_DIR}")
logger.info(f"TEMPLATE_DIR: {TEMPLATE_DIR}")
logger.info(f"STATIC_DIR: {STATIC_DIR}")
logger.info(f"Template dir exists: {os.path.exists(TEMPLATE_DIR)}")
logger.info(f"Static dir exists: {os.path.exists(STATIC_DIR)}")
# Verify templates
if os.path.exists(TEMPLATE_DIR):
    templates = os.listdir(TEMPLATE_DIR)
    logger.info(f"Templates found: {templates}")
    if "login.html" in templates:
        logger.info(f"login.html found in {TEMPLATE_DIR}")
    else:
        logger.warning(f"login.html NOT found in {TEMPLATE_DIR}")
else:
    logger.error(f"Template directory not found: {TEMPLATE_DIR}")
# Initialize Flask with explicit paths
app = Flask(__name__, template_folder=TEMPLATE_DIR, static_folder=STATIC_DIR)
app.logger.setLevel(logging.DEBUG)
app.secret_key = os.getenv("FLASK_SECRET", "default_secret_key_1234567890")
# Configuration
PASSWORD = os.getenv("STUDENT_CLEANER_PASSWORD", "admin")
COLLEGE = os.getenv("COLLEGE_NAME", "FCT College of Nursing Sciences, Gwagwalada")
DEPARTMENT = os.getenv("DEPARTMENT", "Examinations Office")
# ============================================================================
# FIX: Move login_required decorator to the top before any routes use it
# ============================================================================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "logged_in" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function
# Define sets for templates
ND_SETS = ["ND-2024", "ND-2025"]
BN_SETS = ["SET47", "SET48"]
BM_SETS = ["SET2023", "SET2024", "SET2025"]
PROGRAMS = ["ND", "BN", "BM"]
# ============================================================================
# NEW: Enhanced ZIP Creation Functions for Missing Scripts
# ============================================================================
def create_missing_zips():
    """Create ZIP files for scripts that don't automatically create them"""
    try:
        logger.info("🔄 Creating missing ZIP files for problematic scripts...")
       
        # Define script to directory mappings
        script_dirs = {
            "caosce": os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT"),
            "clean": os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ"),
            "split": os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB"),
            "utme": os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT")
        }
       
        created_zips = []
       
        for script_name, clean_dir in script_dirs.items():
            if not os.path.exists(clean_dir):
                logger.warning(f"⚠️ Clean directory not found: {clean_dir}")
                continue
               
            logger.info(f"🔍 Checking {script_name} directory: {clean_dir}")
           
            # Check for existing ZIP files
            existing_zips = [f for f in os.listdir(clean_dir) if f.endswith('.zip')]
           
            if existing_zips:
                logger.info(f"✅ {script_name} already has ZIP files: {existing_zips}")
                continue
               
            # Check for scattered files that need zipping
            scattered_files = []
            scattered_dirs = []
           
            for item in os.listdir(clean_dir):
                item_path = os.path.join(clean_dir, item)
                if os.path.isfile(item_path) and not item.startswith('~') and not item.endswith('.zip'):
                    scattered_files.append(item)
                elif os.path.isdir(item_path) and not item.startswith('CARRYOVER'):
                    scattered_dirs.append(item)
           
            if not scattered_files and not scattered_dirs:
                logger.info(f"ℹ️ No files to zip in {clean_dir}")
                continue
               
            logger.info(f"📦 Found {len(scattered_files)} files and {len(scattered_dirs)} directories to zip for {script_name}")
           
            # Create timestamp for ZIP filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"{script_name.upper()}_RESULTS_{timestamp}.zip"
            zip_path = os.path.join(clean_dir, zip_filename)
           
            # Create ZIP file
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Add files
                for file in scattered_files:
                    file_path = os.path.join(clean_dir, file)
                    zipf.write(file_path, file)
                    logger.info(f"➕ Added file to ZIP: {file}")
               
                # Add directory contents
                for dir_name in scattered_dirs:
                    dir_path = os.path.join(clean_dir, dir_name)
                    for root, dirs, files in os.walk(dir_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.join(dir_name, os.path.relpath(file_path, dir_path))
                            zipf.write(file_path, arcname)
                            logger.info(f"➕ Added directory file to ZIP: {arcname}")
           
            # Verify ZIP was created
            if os.path.exists(zip_path) and os.path.getsize(zip_path) > 0:
                logger.info(f"✅ Created ZIP: {zip_filename} ({os.path.getsize(zip_path)} bytes)")
                created_zips.append(zip_filename)
               
                # Clean up scattered files after successful ZIP creation
                for file in scattered_files:
                    try:
                        os.remove(os.path.join(clean_dir, file))
                        logger.info(f"🗑️ Removed scattered file: {file}")
                    except Exception as e:
                        logger.error(f"Error removing file {file}: {e}")
               
                for dir_name in scattered_dirs:
                    try:
                        shutil.rmtree(os.path.join(clean_dir, dir_name))
                        logger.info(f"🗑️ Removed scattered directory: {dir_name}")
                    except Exception as e:
                        logger.error(f"Error removing directory {dir_name}: {e}")
            else:
                logger.error(f"❌ Failed to create ZIP: {zip_path}")
       
        return created_zips
       
    except Exception as e:
        logger.error(f"❌ Error creating missing ZIPs: {e}")
        return []
def create_zip_from_scattered_content(directory, dirs_to_zip, files_to_zip):
    """Create a ZIP file from scattered directories and files"""
    try:
        if not dirs_to_zip and not files_to_zip:
            return None
           
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dir_name = os.path.basename(directory)
        zip_filename = f"{dir_name}_RESULTS_{timestamp}.zip"
        zip_path = os.path.join(directory, zip_filename)
       
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add individual files
            for file in files_to_zip:
                file_path = os.path.join(directory, file)
                zipf.write(file_path, file)
                print(f"➕ Added file: {file}")
           
            # Add directory contents
            for dir_name in dirs_to_zip:
                dir_path = os.path.join(directory, dir_name)
                for root, dirs, files in os.walk(dir_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.join(dir_name, os.path.relpath(file_path, dir_path))
                        zipf.write(file_path, arcname)
                        print(f"➕ Added directory file: {arcname}")
       
        # Clean up original files after successful ZIP creation
        for file in files_to_zip:
            try:
                os.remove(os.path.join(directory, file))
            except Exception as e:
                print(f"⚠️ Error removing file {file}: {e}")
               
        for dir_name in dirs_to_zip:
            try:
                shutil.rmtree(os.path.join(directory, dir_name))
            except Exception as e:
                print(f"⚠️ Error removing directory {dir_name}: {e}")
       
        print(f"✅ Created ZIP: {zip_filename}")
        return zip_path
       
    except Exception as e:
        print(f"❌ Error creating ZIP: {e}")
        return None
# ============================================================================
# ENHANCED: Universal ZIP Creation and Cleanup Functions - STRICT ENFORCEMENT
# ============================================================================
def create_zip_from_directory(source_dir, zip_filename, remove_original=True):
    """Create ZIP file from directory contents and remove ALL original files"""
    try:
        if not os.path.exists(source_dir):
            logger.error(f"Source directory doesn't exist: {source_dir}")
            return False
     
        zip_path = os.path.join(os.path.dirname(source_dir), zip_filename)
     
        # Collect ALL files to zip (no filtering)
        all_files = []
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                if not file.startswith('~') and not file.startswith('.'):
                    file_path = os.path.join(root, file)
                    all_files.append(file_path)
     
        if not all_files:
            logger.warning(f"No files found to zip in: {source_dir}")
            return False
     
        # Create ZIP file
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_f:
            for file_path in all_files:
                arcname = os.path.relpath(file_path, source_dir)
                zip_f.write(file_path, arcname)
                logger.info(f"Added to ZIP: {arcname}")
     
        # Verify ZIP file
        if os.path.exists(zip_path) and os.path.getsize(zip_path) > 100:
            logger.info(f"✅ Created ZIP: {zip_path} with {len(all_files)} files")
         
            # Remove ALL original files and folders
            if remove_original:
                cleanup_directory(source_dir, keep_files=[zip_filename])
         
            return True
        else:
            logger.error(f"❌ ZIP file created but appears empty: {zip_path}")
            return False
         
    except Exception as e:
        logger.error(f"❌ Failed to create ZIP: {e}")
        return False
def enforce_zip_only_policy(directory):
    """Enforce ZIP-only policy but PROTECT CARRYOVER_RECORDS and recent CARRYOVER_* directories"""
    print(f"🔒 Enforcing ZIP-only policy in: {directory}")
  
    if not os.path.exists(directory):
        return
  
    items = os.listdir(directory)
    zip_files = [f for f in items if f.endswith('.zip')]
    directories = [d for d in items if os.path.isdir(os.path.join(directory, d))]
    other_files = [f for f in items if not f.endswith('.zip') and not os.path.isdir(os.path.join(directory, f))]
  
    print(f"📊 ZIP enforcement - ZIPs: {len(zip_files)}, Dirs: {len(directories)}, Files: {len(other_files)}")
  
    # If we have content but no ZIPs, create a ZIP
    if (directories or other_files) and not zip_files:
        print(f"🔄 Creating ZIP from {len(directories)} directories and {len(other_files)} files")
        create_zip_from_scattered_content(directory, directories, other_files)
       
        # Refresh the items list after ZIP creation
        items = os.listdir(directory)
        zip_files = [f for f in items if f.endswith('.zip')]
        directories = [d for d in items if os.path.isdir(os.path.join(directory, d))]
        other_files = [f for f in items if not f.endswith('.zip') and not os.path.isdir(os.path.join(directory, f))]
  
    removed_dirs = 0
    removed_files = 0
  
    # PROTECTED: Skip CARRYOVER_RECORDS and recent CARRYOVER_* directories
    protected_dirs = []
    removable_dirs = []
  
    for d in directories:
        if d == "CARRYOVER_RECORDS":
            protected_dirs.append(d)
        elif d.startswith("CARRYOVER_"):
            # Check age - protect if less than 7 days old
            dir_path = os.path.join(directory, d)
            try:
                age_days = (datetime.now().timestamp() - os.path.getmtime(dir_path)) / 86400
                if age_days < 7:
                    protected_dirs.append(f"{d} (age: {age_days:.1f}d)")
                else:
                    removable_dirs.append(d)
            except:
                removable_dirs.append(d)
        else:
            removable_dirs.append(d)
  
    if protected_dirs:
        print(f"🛡️ PROTECTED directories (skipped): {protected_dirs}")
  
    # Remove non-protected directories
    for dir_name in removable_dirs:
        dir_path = os.path.join(directory, dir_name)
        try:
            shutil.rmtree(dir_path)
            print(f"🗑️ Removed directory: {dir_name}")
            removed_dirs += 1
        except Exception as e:
            print(f"⚠️ Error removing directory {dir_name}: {e}")
  
    # Remove non-ZIP files
    for file_name in other_files:
        file_path = os.path.join(directory, file_name)
        try:
            os.remove(file_path)
            print(f"🗑️ Removed file: {file_name}")
            removed_files += 1
        except Exception as e:
            print(f"⚠️ Error removing file {file_name}: {e}")
  
    print(f"🧹 CLEANUP completed: {removed_dirs} folders, {removed_files} files removed. CARRYOVER_RECORDS and recent CARRYOVER_* directories protected.")
  
    return len(zip_files), removed_dirs, removed_files
def enforce_zip_only_policy_legacy(clean_dir, zip_base_name):
    """ENFORCE ZIP-ONLY POLICY: Remove all non-ZIP files and create ZIP if needed"""
    try:
        if not os.path.exists(clean_dir):
            return False
         
        logger.info(f"🔒 Enforcing ZIP-only policy in: {clean_dir}")
     
        # Check for existing ZIP files
        zip_files = [f for f in os.listdir(clean_dir) if f.lower().endswith('.zip')]
     
        # Check for scattered files and directories
        scattered_files = []
        scattered_dirs = []
     
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
         
            if item.lower().endswith('.zip'):
                continue # Skip ZIP files
             
            if os.path.isdir(item_path):
                scattered_dirs.append(item)
            elif os.path.isfile(item_path) and not item.startswith('~'):
                scattered_files.append(item)
     
        logger.info(f"📊 ZIP enforcement - ZIPs: {len(zip_files)}, Dirs: {len(scattered_dirs)}, Files: {len(scattered_files)}")
     
        # If we have scattered content but no ZIP, create a ZIP
        if (scattered_dirs or scattered_files) and not zip_files:
            logger.info(f"🔄 Creating ZIP from scattered content")
         
            # Create a temporary directory to consolidate files
            temp_consolidate_dir = os.path.join(clean_dir, f"TEMP_CONSOLIDATE_{int(time.time())}")
            os.makedirs(temp_consolidate_dir, exist_ok=True)
         
            # Move all files and directories to temp directory
            moved_count = 0
            for item in scattered_files + scattered_dirs:
                src_path = os.path.join(clean_dir, item)
                dest_path = os.path.join(temp_consolidate_dir, item)
             
                try:
                    if os.path.isdir(src_path):
                        shutil.move(src_path, dest_path)
                    else:
                        shutil.move(src_path, dest_path)
                    moved_count += 1
                    logger.info(f"Moved to temp: {item}")
                except Exception as e:
                    logger.error(f"Error moving {item}: {e}")
         
            if moved_count > 0:
                # Create ZIP from temp directory
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                zip_filename = f"{zip_base_name}_{timestamp}.zip"
             
                if create_zip_from_directory(temp_consolidate_dir, zip_filename, remove_original=True):
                    # Remove temp directory
                    shutil.rmtree(temp_consolidate_dir, ignore_errors=True)
                    logger.info(f"✅ Successfully created ZIP and cleaned scattered files")
                else:
                    logger.error(f"❌ Failed to create ZIP from temp directory")
            else:
                # Clean up temp directory if no files were moved
                shutil.rmtree(temp_consolidate_dir, ignore_errors=True)
     
        # Final cleanup: Remove any remaining non-ZIP files
        cleanup_scattered_files_strict(clean_dir)
     
        return True
     
    except Exception as e:
        logger.error(f"❌ Error enforcing ZIP-only policy: {e}")
        return False
def cleanup_scattered_files_strict(clean_dir):
    """STRICT CLEANUP: Remove ALL non-ZIP files and folders"""
    try:
        files_removed = 0
        folders_removed = 0
     
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
         
            # Keep ONLY ZIP files
            if item.lower().endswith('.zip'):
                continue
             
            if os.path.isdir(item_path):
                shutil.rmtree(item_path, ignore_errors=True)
                folders_removed += 1
                logger.info(f"🗑️ Removed directory: {item}")
            elif os.path.isfile(item_path):
                os.remove(item_path)
                files_removed += 1
                logger.info(f"🗑️ Removed file: {item}")
     
        if files_removed > 0 or folders_removed > 0:
            logger.info(f"🧹 STRICT Cleanup completed: {folders_removed} folders, {files_removed} files removed. ONLY ZIP files remain.")
     
        return True
     
    except Exception as e:
        logger.error(f"❌ Error during strict cleanup: {e}")
        return False
def cleanup_directory(directory, keep_files=None):
    """Remove all files and subdirectories except specified files"""
    if keep_files is None:
        keep_files = []
 
    try:
        files_removed = 0
        folders_removed = 0
     
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
         
            # Skip files in keep list
            if item in keep_files:
                continue
             
            if os.path.isdir(item_path):
                shutil.rmtree(item_path, ignore_errors=True)
                folders_removed += 1
                logger.info(f"🗑️ Removed directory: {item}")
            elif os.path.isfile(item_path):
                os.remove(item_path)
                files_removed += 1
                logger.info(f"🗑️ Removed file: {item}")
     
        logger.info(f"🧹 Cleanup completed: {folders_removed} folders, {files_removed} files removed")
        return True
     
    except Exception as e:
        logger.error(f"❌ Error during directory cleanup: {e}")
        return False
def create_result_zip(clean_dir, set_name, result_folder):
    """Create ZIP file - FIXED with proper cleanup."""
    try:
        folder_path = os.path.join(clean_dir, result_folder)
     
        if not os.path.exists(folder_path):
            logger.error(f"Result folder doesn't exist: {folder_path}")
            return False
     
        # Collect all files
        all_files = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(('.xlsx', '.csv', '.pdf', '.txt', '.json')):
                    file_path = os.path.join(root, file)
                    all_files.append(file_path)
     
        if not all_files:
            logger.warning(f"⚠️ No files found to zip in: {folder_path}")
            return False
     
        zip_filename = f"{set_name}_RESULT-{result_folder.split('-')[-1]}.zip"
        zip_path = os.path.join(clean_dir, zip_filename)
     
        # Create ZIP with explicit close
        zip_file = None
        try:
            zip_file = zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)
            for file_path in all_files:
                arcname = os.path.relpath(file_path, folder_path)
                zip_file.write(file_path, arcname)
                logger.debug(f"Added: {arcname}")
        finally:
            if zip_file:
                zip_file.close()
     
        # Verify ZIP
        if os.path.exists(zip_path):
            zip_size = os.path.getsize(zip_path)
            if zip_size > 100:
                # Verify integrity
                try:
                    with zipfile.ZipFile(zip_path, 'r') as test_zip:
                        bad_file = test_zip.testzip()
                        if bad_file:
                            logger.error(f"❌ Corrupted file in ZIP: {bad_file}")
                            return False
                except zipfile.BadZipFile:
                    logger.error(f"❌ ZIP file is corrupted: {zip_path}")
                    return False
             
                logger.info(f"✅ Created ZIP: {zip_path} ({zip_size:,} bytes, {len(all_files)} files)")
             
                # Safe cleanup
                try:
                    shutil.rmtree(folder_path, ignore_errors=True)
                    logger.info(f"🗑️ Removed original folder: {result_folder}")
                except Exception as e:
                    logger.warning(f"⚠️ Could not remove folder: {e}")
             
                return True
            else:
                logger.error(f"❌ ZIP file too small: {zip_size} bytes")
                return False
        else:
            logger.error(f"❌ ZIP file not created: {zip_path}")
            return False
         
    except Exception as e:
        logger.error(f"❌ Failed to create ZIP: {e}")
        import traceback
        traceback.print_exc()
        return False
def cleanup_scattered_files(clean_dir, zip_filename):
    """Remove all scattered files and folders after successful zipping - ENHANCED"""
    try:
        files_removed = 0
        folders_removed = 0
     
        # Remove any result directories (except the ZIP we just created)
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
         
            # Skip the ZIP file we just created
            if item == zip_filename:
                continue
             
            if os.path.isdir(item_path) and ("RESULT" in item or "RESIT" in item):
                shutil.rmtree(item_path, ignore_errors=True)
                folders_removed += 1
                logger.info(f"🗑️ Removed scattered folder: {item}")
         
            # Remove individual result files (keep only ZIP files)
            elif os.path.isfile(item_path) and not item.lower().endswith('.zip'):
                # Skip course files and other important files
                if not any(pattern in item for pattern in ['course', 'credit', 'CARRYOVER']):
                    os.remove(item_path)
                    files_removed += 1
                    logger.info(f"🗑️ Removed scattered file: {item}")
             
        logger.info(f"🧹 Cleanup completed: {folders_removed} folders, {files_removed} files removed. Only ZIP files remain.")
        return True
     
    except Exception as e:
        logger.error(f"❌ Error during cleanup: {e}")
        return False
# ============================================================================
# ENHANCED: Universal File Processing with ZIP Enforcement
# ============================================================================
def ensure_zipped_results(clean_dir, script_name, set_name=None):
    """Ensure all results are properly zipped and scattered files are removed"""
    try:
        logger.info(f"🔍 Checking ZIP status in: {clean_dir}")
     
        if not os.path.exists(clean_dir):
            logger.warning(f"Clean directory doesn't exist: {clean_dir}")
            return False
     
        # Check for existing ZIP files
        zip_files = [f for f in os.listdir(clean_dir) if f.lower().endswith('.zip')]
     
        # Check for scattered result directories
        result_dirs = [d for d in os.listdir(clean_dir)
                     if os.path.isdir(os.path.join(clean_dir, d)) and
                     ("RESULT" in d or "RESIT" in d)]
     
        # Check for scattered files
        scattered_files = [f for f in os.listdir(clean_dir)
                          if os.path.isfile(os.path.join(clean_dir, f)) and
                          not f.lower().endswith('.zip') and
                          not f.startswith('~')]
     
        logger.info(f"📊 Cleanup status - ZIPs: {len(zip_files)}, Dirs: {len(result_dirs)}, Files: {len(scattered_files)}")
     
        # If we have result directories but no ZIPs, create ZIPs
        if result_dirs and not zip_files:
            logger.info(f"🔄 Creating ZIPs from {len(result_dirs)} result directories")
            for result_dir in result_dirs:
                if set_name:
                    create_result_zip(clean_dir, set_name, result_dir)
                else:
                    # Extract set name from directory name or use generic
                    dir_set_name = extract_set_from_directory(result_dir) or "RESULTS"
                    create_result_zip(clean_dir, dir_set_name, result_dir)
     
        # Clean up any remaining scattered files
        if scattered_files:
            logger.info(f"🧹 Cleaning up {len(scattered_files)} scattered files")
            cleanup_scattered_files(clean_dir, "dummy_zip.zip" if not zip_files else zip_files[0])
     
        return True
     
    except Exception as e:
        logger.error(f"❌ Error ensuring zipped results: {e}")
        return False
def extract_set_from_directory(dir_name):
    """Extract set name from directory name"""
    patterns = [
        r'(ND-\d{4})',
        r'(SET\d+)',
        r'(SET\d{4})',
        r'(BN-\w+)',
        r'(BM-\w+)'
    ]
 
    for pattern in patterns:
        match = re.search(pattern, dir_name)
        if match:
            return match.group(1)
 
    return None
# ============================================================================
# FIXED: Enhanced BM ZIP Creation with Proper Set Filtering
# ============================================================================
def create_result_zip(clean_dir, set_name, result_folder):
    """Create ZIP file for a specific result folder and clean up scattered files"""
    try:
        folder_path = os.path.join(clean_dir, result_folder)
     
        # Collect all files from the result folder
        all_files = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(('.xlsx', '.csv', '.pdf')):
                    all_files.append(os.path.join(root, file))
     
        if all_files:
            zip_filename = f"{set_name}_RESULT-{result_folder.split('-')[-1]}.zip"
            zip_path = os.path.join(clean_dir, zip_filename)
         
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_f:
                for file_path in all_files:
                    arcname = os.path.relpath(file_path, folder_path)
                    zip_f.write(file_path, arcname)
         
            # Verify ZIP file
            if os.path.exists(zip_path) and os.path.getsize(zip_path) > 100:
                logger.info(f"✅ Created ZIP: {zip_path} with {len(all_files)} files")
             
                # Clean up scattered files and folders
                cleanup_scattered_files(clean_dir, zip_filename)
                return True
            else:
                logger.warning(f"⚠️ ZIP file created but appears empty: {zip_path}")
                return False
        else:
            logger.warning(f"⚠️ No files found to zip in: {folder_path}")
            return False
         
    except Exception as e:
        logger.error(f"❌ Failed to create ZIP: {e}")
        return False
# ============================================================================
# FIXED: Enhanced cleanup_scattered_files function
# ============================================================================
def cleanup_scattered_files(clean_dir, zip_filename):
    """Remove all scattered files and folders after successful zipping - ENHANCED"""
    try:
        files_removed = 0
        folders_removed = 0
     
        # Remove any result directories (except the ZIP we just created)
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
         
            # Skip the ZIP file we just created
            if item == zip_filename:
                continue
             
            if os.path.isdir(item_path) and ("RESULT" in item or "RESIT" in item):
                shutil.rmtree(item_path, ignore_errors=True)
                folders_removed += 1
                logger.info(f"🗑️ Removed scattered folder: {item}")
         
            # Remove individual result files (keep only ZIP files)
            elif os.path.isfile(item_path) and not item.lower().endswith('.zip'):
                # Skip course files and other important files
                if not any(pattern in item for pattern in ['course', 'credit', 'CARRYOVER']):
                    os.remove(item_path)
                    files_removed += 1
                    logger.info(f"🗑️ Removed scattered file: {item}")
             
        logger.info(f"🧹 Cleanup completed: {folders_removed} folders, {files_removed} files removed. Only ZIP files remain.")
        return True
     
    except Exception as e:
        logger.error(f"❌ Error during cleanup: {e}")
        return False
# ============================================================================
# UPDATED: Route Names and Functions with Individual Semester Selection
# ============================================================================
# Update existing BN form route with individual semester selection
@app.route("/bn_regular_exam_processor")
@login_required
def bn_regular_exam_processor():
    """Basic Nursing regular exam processor form"""
    try:
        bn_sets = get_available_sets("BN")
     
        # Define BN semesters
        bn_semesters = [
            {"key": "N-FIRST-YEAR-FIRST-SEMESTER", "display": "Year 1 - First Semester"},
            {"key": "N-FIRST-YEAR-SECOND-SEMESTER", "display": "Year 1 - Second Semester"},
            {"key": "N-SECOND-YEAR-FIRST-SEMESTER", "display": "Year 2 - First Semester"},
            {"key": "N-SECOND-YEAR-SECOND-SEMESTER", "display": "Year 2 - Second Semester"},
            {"key": "N-THIRD-YEAR-FIRST-SEMESTER", "display": "Year 3 - First Semester"},
            {"key": "N-THIRD-YEAR-SECOND-SEMESTER", "display": "Year 3 - Second Semester"},
        ]
     
        return render_template(
            "bn_regular_exam_processor.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            bn_sets=bn_sets,
            bn_semesters=bn_semesters
        )
    except Exception as e:
        logger.error(f"BN regular exam processor form error: {e}")
        flash(f"Error loading BN exam processor: {str(e)}", "error")
        return redirect(url_for("dashboard"))
# Update existing BM form route with individual semester selection
@app.route("/bm_regular_exam_processor")
@login_required
def bm_regular_exam_processor():
    """Basic Midwifery regular exam processor form"""
    try:
        bm_sets = get_available_sets("BM")
     
        # Define BM semesters
        bm_semesters = [
            {"key": "M-FIRST-YEAR-FIRST-SEMESTER", "display": "Year 1 - First Semester"},
            {"key": "M-FIRST-YEAR-SECOND-SEMESTER", "display": "Year 1 - Second Semester"},
            {"key": "M-SECOND-YEAR-FIRST-SEMESTER", "display": "Year 2 - First Semester"},
            {"key": "M-SECOND-YEAR-SECOND-SEMESTER", "display": "Year 2 - Second Semester"},
            {"key": "M-THIRD-YEAR-FIRST-SEMESTER", "display": "Year 3 - First Semester"},
            {"key": "M-THIRD-YEAR-SECOND-SEMESTER", "display": "Year 3 - Second Semester"},
        ]
     
        return render_template(
            "bm_regular_exam_processor.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            bm_sets=bm_sets,
            bm_semesters=bm_semesters
        )
    except Exception as e:
        logger.error(f"BM regular exam processor form error: {e}")
        flash(f"Error loading BM exam processor: {str(e)}", "error")
        return redirect(url_for("dashboard"))
# ============================================================================
# FIX 1: ADDED MISSING ROUTE: ND Regular Exam Processor Form
# ============================================================================
@app.route("/nd_regular_exam_processor")
@login_required
def nd_regular_exam_processor():
    """National Diploma regular exam processor form"""
    try:
        nd_sets = get_available_sets("ND")
     
        # Define ND semesters
        nd_semesters = [
            {"key": "ND-FIRST-YEAR-FIRST-SEMESTER", "display": "Year 1 - First Semester"},
            {"key": "ND-FIRST-YEAR-SECOND-SEMESTER", "display": "Year 1 - Second Semester"},
            {"key": "ND-SECOND-YEAR-FIRST-SEMESTER", "display": "Year 2 - First Semester"},
            {"key": "ND-SECOND-YEAR-SECOND-SEMESTER", "display": "Year 2 - Second Semester"},
        ]
     
        return render_template(
            "nd_regular_exam_processor.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            nd_sets=nd_sets,
            nd_semesters=nd_semesters
        )
    except Exception as e:
        logger.error(f"ND regular exam processor form error: {e}")
        flash(f"Error loading ND exam processor: {str(e)}", "error")
        return redirect(url_for("dashboard"))
# ============================================================================
# FIXED: BN Carryover Processing Route - ACCEPTS BOTH NAMING CONVENTIONS
# ============================================================================
@app.route("/process_bn_resit", methods=["POST"])
@login_required
def process_bn_resit():
    """Process Basic Nursing carryover/resit results"""
    try:
        logger.info("BN RESIT: Route called")
        logger.info(f"Form data: {request.form}")
        logger.info(f"Files: {request.files}")
     
        # ✅ FIX: Accept both naming conventions
        resit_set = request.form.get("resit_set") or request.form.get("bn_resit_set")
        resit_semester = request.form.get("resit_semester") or request.form.get("bn_resit_semester")
        resit_file = request.files.get("resit_file") or request.files.get("bn_resit_file")
     
        logger.info(f"BN RESIT: Set={resit_set}, Semester={resit_semester}, File={resit_file}")
     
        # Validation
        if not all([resit_set, resit_semester, resit_file]):
            flash("All fields are required for carryover processing", "error")
            return redirect(url_for("bn_regular_exam_processor"))
     
        # ✅ FIX: Save uploaded file to correct directory - RAW_RESULTS/CARRYOVER
        upload_dir = os.path.join(BASE_DIR, "BN", resit_set, "RAW_RESULTS", "CARRYOVER")
        os.makedirs(upload_dir, exist_ok=True)
     
        filename = f"bn_resit_{resit_semester}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = os.path.join(upload_dir, filename)
        resit_file.save(file_path)
     
        # Process BN carryover results
        result = process_carryover_results("BN", resit_set, resit_semester, file_path)
     
        if result["success"]:
            flash(f"BN carryover results processed successfully! {result.get('message', '')}", "success")
        else:
            flash(f"BN carryover processing failed: {result.get('message', 'Unknown error')}", "error")
         
        return redirect(url_for("bn_carryover"))
     
    except Exception as e:
        logger.error(f"BN carryover processing error: {e}")
        flash(f"Error processing BN carryover results: {str(e)}", "error")
        return redirect(url_for("bn_regular_exam_processor"))
# ============================================================================
# FIXED: BM Carryover Processing Route - ACCEPTS BOTH NAMING CONVENTIONS
# ============================================================================
@app.route("/process_bm_resit", methods=["POST"])
@login_required
def process_bm_resit():
    """Process Basic Midwifery carryover/resit results"""
    try:
        logger.info("BM RESIT: Route called")
        logger.info(f"Form data: {request.form}")
        logger.info(f"Files: {request.files}")
     
        # ✅ FIX: Accept both naming conventions
        resit_set = request.form.get("resit_set") or request.form.get("bm_resit_set")
        resit_semester = request.form.get("resit_semester") or request.form.get("bm_resit_semester")
        resit_file = request.files.get("resit_file") or request.files.get("bm_resit_file")
     
        logger.info(f"BM RESIT: Set={resit_set}, Semester={resit_semester}, File={resit_file}")
     
        # Validation
        if not all([resit_set, resit_semester, resit_file]):
            flash("All fields are required for carryover processing", "error")
            return redirect(url_for("bm_regular_exam_processor"))
     
        # ✅ FIX: Save uploaded file to correct directory - RAW_RESULTS/CARRYOVER
        upload_dir = os.path.join(BASE_DIR, "BM", resit_set, "RAW_RESULTS", "CARRYOVER")
        os.makedirs(upload_dir, exist_ok=True)
     
        filename = f"bm_resit_{resit_semester}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = os.path.join(upload_dir, filename)
        resit_file.save(file_path)
     
        # Process BM carryover results
        result = process_carryover_results("BM", resit_set, resit_semester, file_path)
     
        if result["success"]:
            flash(f"BM carryover results processed successfully! {result.get('message', '')}", "success")
        else:
            flash(f"BM carryover processing failed: {result.get('message', 'Unknown error')}", "error")
         
        return redirect(url_for("bm_carryover"))
     
    except Exception as e:
        logger.error(f"BM carryover processing error: {e}")
        flash(f"Error processing BM carryover results: {str(e)}", "error")
        return redirect(url_for("bm_regular_exam_processor"))
def get_program_from_set(set_name):
    """Determine program from set name"""
    if set_name in ND_SETS:
        return "ND"
    elif set_name in BN_SETS:
        return "BN"
    elif set_name in BM_SETS:
        return "BM"
    return None
# ============================================================================
# FIX 1: Enhanced Program Detection Function
# ============================================================================
def detect_program_from_set(set_name):
    """Enhanced program detection from set name with fallback strategies"""
    if not set_name or set_name == "all":
        return None
 
    set_upper = set_name.upper()
 
    # Check against defined sets first
    if set_name in ND_SETS:
        return "ND"
    elif set_name in BN_SETS:
        return "BN"
    elif set_name in BM_SETS:
        return "BM"
 
    # Fallback: Check for explicit program indicators in set name
    if any(x in set_upper for x in ['BN', 'NURSING', 'N-']):
        return "BN"
    elif any(x in set_upper for x in ['BM', 'MIDWIFE', 'MIDWIFERY', 'M-']):
        return "BM"
    elif any(x in set_upper for x in ['ND', 'DIPLOMA']):
        return "ND"
 
    # Check SET number patterns
    if set_upper.startswith("SET4"): # SET47, SET48
        return "BN"
    elif set_upper.startswith("SET") and any(c.isdigit() for c in set_upper):
        # Assume other numbered SETs are BM unless proven otherwise
        return "BM"
    elif set_upper.startswith("ND-"):
        return "ND"
 
    # Final fallback
    logger.warning(f"⚠️ Could not determine program from set '{set_name}', defaulting to ND")
    return "ND"
# ============================================================================
# FIX B: UPDATED: Semester Key Standardization Function - ND-SPECIFIC VERSION
# ============================================================================
def standardize_semester_key_nd(semester_key):
    """Standardize semester key to canonical format for ND ONLY."""
    if not semester_key:
        return None
 
    key_upper = semester_key.upper()
    prefix = "ND-" # Always ND prefix for ND
 
    # Define canonical mappings for ND
    canonical_mappings = {
        # First Year First Semester variants
        ("FIRST", "YEAR", "FIRST", "SEMESTER"): f"{prefix}FIRST-YEAR-FIRST-SEMESTER",
        ("1ST", "YEAR", "1ST", "SEMESTER"): f"{prefix}FIRST-YEAR-FIRST-SEMESTER",
        ("YEAR", "1", "SEMESTER", "1"): f"{prefix}FIRST-YEAR-FIRST-SEMESTER",
        ("FIRST", "SEMESTER"): f"{prefix}FIRST-YEAR-FIRST-SEMESTER",
        ("SEMESTER", "1"): f"{prefix}FIRST-YEAR-FIRST-SEMESTER",
     
        # First Year Second Semester variants
        ("FIRST", "YEAR", "SECOND", "SEMESTER"): f"{prefix}FIRST-YEAR-SECOND-SEMESTER",
        ("1ST", "YEAR", "2ND", "SEMESTER"): f"{prefix}FIRST-YEAR-SECOND-SEMESTER",
        ("YEAR", "1", "SEMESTER", "2"): f"{prefix}FIRST-YEAR-SECOND-SEMESTER",
        ("SECOND", "SEMESTER"): f"{prefix}FIRST-YEAR-SECOND-SEMESTER",
        ("SEMESTER", "2"): f"{prefix}FIRST-YEAR-SECOND-SEMESTER",
     
        # Second Year First Semester variants
        ("SECOND", "YEAR", "FIRST", "SEMESTER"): f"{prefix}SECOND-YEAR-FIRST-SEMESTER",
        ("2ND", "YEAR", "1ST", "SEMESTER"): f"{prefix}SECOND-YEAR-FIRST-SEMESTER",
        ("YEAR", "2", "SEMESTER", "1"): f"{prefix}SECOND-YEAR-FIRST-SEMESTER",
     
        # Second Year Second Semester variants
        ("SECOND", "YEAR", "SECOND", "SEMESTER"): f"{prefix}SECOND-YEAR-SECOND-SEMESTER",
        ("2ND", "YEAR", "2ND", "SEMESTER"): f"{prefix}SECOND-YEAR-SECOND-SEMESTER",
        ("YEAR", "2", "SEMESTER", "2"): f"{prefix}SECOND-YEAR-SECOND-SEMESTER",
    }
 
    # Extract key components using regex
    patterns = [
        r'(FIRST|1ST|YEAR.?1).*?(FIRST|1ST|SEMESTER.?1)',
        r'(FIRST|1ST|YEAR.?1).*?(SECOND|2ND|SEMESTER.?2)',
        r'(SECOND|2ND|YEAR.?2).*?(FIRST|1ST|SEMESTER.?1)',
        r'(SECOND|2ND|YEAR.?2).*?(SECOND|2ND|SEMESTER.?2)',
    ]
 
    for pattern_idx, pattern in enumerate(patterns):
        if re.search(pattern, key_upper):
            if pattern_idx == 0:
                return f"{prefix}FIRST-YEAR-FIRST-SEMESTER"
            elif pattern_idx == 1:
                return f"{prefix}FIRST-YEAR-SECOND-SEMESTER"
            elif pattern_idx == 2:
                return f"{prefix}SECOND-YEAR-FIRST-SEMESTER"
            elif pattern_idx == 3:
                return f"{prefix}SECOND-YEAR-SECOND-SEMESTER"
 
    # Check for direct matches in canonical mappings
    for key_components, canonical_key in canonical_mappings.items():
        if all(component in key_upper for component in key_components):
            return canonical_key
 
    # If no match but contains ND prefix, return as-is
    if key_upper.startswith("ND-"):
        return semester_key.upper()
 
    # If no match, try to construct from known patterns
    if "FIRST" in key_upper and "FIRST" in key_upper:
        return f"{prefix}FIRST-YEAR-FIRST-SEMESTER"
    elif "FIRST" in key_upper and "SECOND" in key_upper:
        return f"{prefix}FIRST-YEAR-SECOND-SEMESTER"
    elif "SECOND" in key_upper and "FIRST" in key_upper:
        return f"{prefix}SECOND-YEAR-FIRST-SEMESTER"
    elif "SECOND" in key_upper and "SECOND" in key_upper:
        return f"{prefix}SECOND-YEAR-SECOND-SEMESTER"
 
    # If no match, return original with ND prefix
    logger.warning(f"Could not standardize ND semester key: {semester_key}, using prefix: {prefix}")
    return f"{prefix}{semester_key.replace('-', ' ').upper().replace(' ', '-')}"
# Original function kept for BN/BM compatibility
def standardize_semester_key(semester_key, program=None):
    """Universal semester key standardization."""
    if not semester_key:
        return None
 
    key_upper = semester_key.upper()
 
    # Determine program and prefix
    if program == "ND" or key_upper.startswith("ND-") or "-ND-" in key_upper:
        prefix = "ND-"
    elif program == "BN" or key_upper.startswith(("N-", "BN-")) or "NURSING" in key_upper:
        prefix = "N-"
    elif program == "BM" or key_upper.startswith(("M-", "BM-")) or "MIDWIFE" in key_upper:
        prefix = "M-"
    else:
        # Try to detect from content
        if "NURSING" in key_upper:
            prefix = "N-"
        elif "MIDWIFE" in key_upper:
            prefix = "M-"
        else:
            prefix = "ND-" # Default
 
    # Pattern matching for year and semester
    patterns = {
        r'FIRST.*FIRST': f'{prefix}FIRST-YEAR-FIRST-SEMESTER',
        r'FIRST.*SECOND': f'{prefix}FIRST-YEAR-SECOND-SEMESTER',
        r'SECOND.*FIRST': f'{prefix}SECOND-YEAR-FIRST-SEMESTER',
        r'SECOND.*SECOND': f'{prefix}SECOND-YEAR-SECOND-SEMESTER',
        r'THIRD.*FIRST': f'{prefix}THIRD-YEAR-FIRST-SEMESTER',
        r'THIRD.*SECOND': f'{prefix}THIRD-YEAR-SECOND-SEMESTER',
        r'(1|I).*1': f'{prefix}FIRST-YEAR-FIRST-SEMESTER',
        r'(1|I).*2': f'{prefix}FIRST-YEAR-SECOND-SEMESTER',
        r'(2|II).*1': f'{prefix}SECOND-YEAR-FIRST-SEMESTER',
        r'(2|II).*2': f'{prefix}SECOND-YEAR-SECOND-SEMESTER',
        r'(3|III).*1': f'{prefix}THIRD-YEAR-FIRST-SEMESTER',
        r'(3|III).*2': f'{prefix}THIRD-YEAR-SECOND-SEMESTER',
    }
 
    for pattern, result in patterns.items():
        if re.search(pattern, key_upper):
            logger.info(f"✅ Standardized '{semester_key}' → '{result}' (regex match)")
            return result
 
    # If already in correct format, return as-is
    if key_upper.startswith((prefix,)):
        return key_upper
 
    # Last resort: add prefix
    logger.warning(f"⚠️ Could not parse '{semester_key}', adding prefix: {prefix}")
    return f"{prefix}{key_upper.replace('-', ' ').replace('_', ' ').strip()}"
# NEW: Helper function for GPA tracking across all semesters
def get_previous_semesters_for_display(current_semester_key):
    """Get list of previous semesters for GPA display in mastersheet."""
    current_standard = standardize_semester_key(current_semester_key)
 
    semester_mapping = {
        "ND-FIRST-YEAR-FIRST-SEMESTER": [],
        "ND-FIRST-YEAR-SECOND-SEMESTER": ["Semester 1"],
        "ND-SECOND-YEAR-FIRST-SEMESTER": ["Semester 1", "Semester 2"],
        "ND-SECOND-YEAR-SECOND-SEMESTER": ["Semester 1", "Semester 2", "Semester 3"]
    }
 
    return semester_mapping.get(current_standard, [])
# Jinja2 filters
def datetimeformat(timestamp):
    try:
        dt = datetime.fromtimestamp(timestamp)
        return dt.strftime("%b %d, %Y %I:%M %p")
    except Exception:
        return "Unknown"
app.jinja_env.filters["datetimeformat"] = datetimeformat
def filesizeformat(size):
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"
app.jinja_env.filters["filesizeformat"] = filesizeformat
def is_local_environment():
    try:
        hostname = socket.gethostname()
        ip = socket.gethostbyname(hostname)
        if "railway" in hostname.lower() or ip.startswith("10.") or ip.startswith("172."):
            return False
        return os.path.exists(BASE_DIR)
    except Exception as e:
        logger.error(f"Error checking environment: {e}")
        return True
# Check if BASE_DIR exists
if not os.path.exists(BASE_DIR):
    logger.info(f"Creating BASE_DIR: {BASE_DIR}")
    os.makedirs(BASE_DIR, exist_ok=True)
else:
    logger.info(f"BASE_DIR already exists: {BASE_DIR}")
# Define required subdirectories structure - ONLY create if they don't exist
required_dirs = [
    # ND Structure
    os.path.join(BASE_DIR, "ND", "ND-COURSES"),
    os.path.join(BASE_DIR, "ND", "ND-2024", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "ND", "ND-2024", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "ND", "ND-2024", "CLEAN_RESULTS", "CARRYOVER_RECORDS"), # NEW!
    os.path.join(BASE_DIR, "ND", "ND-2025", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "ND", "ND-2025", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "ND", "ND-2025", "CLEAN_RESULTS", "CARRYOVER_RECORDS"), # NEW!
 
    # BN Structure
    os.path.join(BASE_DIR, "BN", "BN-COURSES"),
    os.path.join(BASE_DIR, "BN", "SET47", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BN", "SET47", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "BN", "SET47", "CLEAN_RESULTS", "CARRYOVER_RECORDS"), # NEW!
    os.path.join(BASE_DIR, "BN", "SET48", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BN", "SET48", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "BN", "SET48", "CLEAN_RESULTS", "CARRYOVER_RECORDS"), # NEW!
 
    # BM Structure (if applicable)
    os.path.join(BASE_DIR, "BM", "BM-COURSES"),
    os.path.join(BASE_DIR, "BM", "SET2023", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2023", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2023", "CLEAN_RESULTS", "CARRYOVER_RECORDS"), # NEW!
    # ... etc for other BM sets
    os.path.join(BASE_DIR, "BM", "SET2024", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2024", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2024", "CLEAN_RESULTS", "CARRYOVER_RECORDS"), # NEW!
    os.path.join(BASE_DIR, "BM", "SET2025", "RAW_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2025", "CLEAN_RESULTS"),
    os.path.join(BASE_DIR, "BM", "SET2025", "CLEAN_RESULTS", "CARRYOVER_RECORDS"), # NEW!
 
    # Other Results Structure
    os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_PUTME_RESULT"),
    os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT"),
    os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_CANDIDATE_BATCHES"),
    os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_UTME_CANDIDATES"),
    os.path.join(BASE_DIR, "CAOSCE_RESULT", "RAW_CAOSCE_RESULT"),
    os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT"),
    os.path.join(BASE_DIR, "OBJ_RESULT", "RAW_OBJ"),
    os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ"),
    os.path.join(BASE_DIR, "JAMB_DB", "RAW_JAMB_DB"),
    os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB"),
]
# Only create directories that don't exist
for dir_path in required_dirs:
    if not os.path.exists(dir_path):
        try:
            os.makedirs(dir_path, exist_ok=True)
            logger.info(f"Created subdirectory: {dir_path}")
        except Exception as e:
            logger.error(f"Could not create {dir_path}: {e}")
    else:
        logger.info(f"Directory already exists: {dir_path}")
# ============================================================================
# UPDATED: Script mapping - CHANGED: ND-specific processor
# ============================================================================
SCRIPT_MAP = {
    "utme": "utme_result.py",
    "caosce": "caosce_result.py",
    "clean": "obj_results.py",
    "split": "split_names.py",
    "exam_processor_nd": "exam_result_processor.py",
    "exam_processor_bn": "exam_processor_bn.py",
    "exam_processor_bm": "exam_processor_bm.py",
    "nd_carryover_processor": "nd_carryover_processor.py" # CHANGED: ND-specific
}
# ============================================================================
# UPDATED: Success indicators - CHANGED: ND-specific key
# ============================================================================
SUCCESS_INDICATORS = {
    "utme": [
        r"Processing: (PUTME 2025-Batch\d+[A-Z] Post-UTME Quiz-grades\.xlsx)",
        r"Saved processed file: (UTME_RESULT_.*?\.csv)",
        r"Saved processed file: (UTME_RESULT_.*?\.xlsx)",
        r"Saved processed file: (PUTME_COMBINE_RESULT_.*?\.xlsx)",
    ],
    "caosce": [
        r"Processed (CAOSCE SET2023A.*?|VIVA \([0-9]+\)\.xlsx) \(\d+ rows read\)",
        r"Saved processed file: (CAOSCE_RESULT_.*?\.csv)",
    ],
    "clean": [
        r"Processing: (Set2025-.*?\.xlsx|ND\d{4}-SET\d+.*?\.xlsx|.*ND.*SET.*\.xlsx)",
        r"Cleaned CSV saved in.*?cleaned_(Set2025-.*?\.csv|ND\d{4}-SET\d+.*?\.csv)",
        r"Master CSV saved in.*?master_cleaned_results\.csv",
        r"All processing completed successfully!",
    ],
    "split": [r"Saved processed file: (clean_jamb_DB_.*?\.csv)"],
    "exam_processor_nd": [
        r"PROCESSING SEMESTER: (ND-[A-Za-z0-9\s\-]+)",
        r"Successfully processed .*",
        r"Mastersheet saved:.*",
        r"Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"Processing complete",
        r"ND Examination Results Processing completed successfully",
        r"Applying upgrade rule:.*→ 50",
        r"Upgraded \d+ scores from.*to 50",
        r"Identified \d+ carryover students",
        r"Carryover records saved:.*",
        r"Processing resit results for.*",
        r"Updated \d+ scores for \d+ students",
    ],
    "exam_processor_bn": [
        r"PROCESSING SET: (SET47|SET48)",
        r"Successfully processed .*",
        r"Mastersheet saved:.*",
        r"Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"Processing complete",
        r"Basic Nursing Examination Results Processing completed successfully",
        r"Applying upgrade rule:.*→ 50",
        r"Upgraded \d+ scores from.*to 50",
        r"Identified \d+ carryover students",
        r"Carryover records saved:.*",
        r"Processing resit results for.*",
        r"Updated \d+ scores for \d+ students",
    ],
    "exam_processor_bm": [
        r"PROCESSING SET: (SET2023|SET2024|SET2025)",
        r"Successfully processed .*",
        r"Mastersheet saved:.*",
        r"Found \d+ raw files",
        r"Processing: (.*?\.xlsx)",
        r"Processing complete",
        r"Basic Midwifery Examination Results Processing completed successfully",
        r"Applying upgrade rule:.*→ 50",
        r"Upgraded \d+ scores from.*to 50",
        r"Identified \d+ carryover students",
        r"Carryover records saved:.*",
        r"Processing resit results for.*",
        r"Updated \d+ scores for \d+ students",
    ],
    "nd_carryover_processor": [ # CHANGED: ND-specific key
        r"Updated \d+ scores for \d+ students",
        r"ND CARRYOVER PROCESSING COMPLETED",
        r"Saved updated mastersheet",
    ],
}
ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv", "zip", "pdf"}
# Helper Functions
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
def get_raw_directory(script_name, program=None, set_name=None):
    """Get the RAW_RESULTS directory for a specific script/program/set"""
    logger.info(f"Getting raw directory for: script={script_name}, program={program}, set={set_name}")
 
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] or script_name == "nd_carryover_processor":
        if program and set_name:
            raw_dir = os.path.join(BASE_DIR, program, set_name, "RAW_RESULTS")
            logger.info(f"Exam processor raw directory: {raw_dir}")
            return raw_dir
        return BASE_DIR
 
    raw_paths = {
        "utme": os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_PUTME_RESULT"),
        "caosce": os.path.join(BASE_DIR, "CAOSCE_RESULT", "RAW_CAOSCE_RESULT"),
        "clean": os.path.join(BASE_DIR, "OBJ_RESULT", "RAW_OBJ"),
        "split": os.path.join(BASE_DIR, "JAMB_DB", "RAW_JAMB_DB"),
    }
    raw_dir = raw_paths.get(script_name, BASE_DIR)
    logger.info(f"Other script raw directory: {raw_dir}")
    return raw_dir
def get_clean_directory(script_name, program=None, set_name=None):
    """Get the CLEAN_RESULTS directory for a specific script/program/set"""
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"] or script_name == "nd_carryover_processor":
        if program and set_name:
            return os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        return BASE_DIR
 
    clean_paths = {
        "utme": os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT"),
        "caosce": os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT"),
        "clean": os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ"),
        "split": os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB"),
    }
    return clean_paths.get(script_name, BASE_DIR)
# ============================================================================
# UPDATED: get_input_directory function - CHANGED: ND-specific handling
# ============================================================================
def get_input_directory(script_name, program=None, set_name=None):
    """Returns the correct input directory for raw results"""
    logger.info(f"Getting input directory for: {script_name}, program={program}, set={set_name}")
 
    # ND carryover processor - specific handling
    if script_name == "nd_carryover_processor":
        if set_name and set_name != "all":
            carryover_dir = os.path.join(BASE_DIR, "ND", set_name, "RAW_RESULTS", "CARRYOVER")
            logger.info(f"ND carryover input directory: {carryover_dir}")
            return carryover_dir
        else:
            logger.error(f"ND carryover requires specific set, got: {set_name}")
            return os.path.join(BASE_DIR, "ND")
 
    # Regular exam processors
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
        if program and set_name and set_name != "all":
            input_dir = os.path.join(BASE_DIR, program, set_name, "RAW_RESULTS")
            logger.info(f"Exam processor input directory: {input_dir}")
            return input_dir
        input_dir = os.path.join(BASE_DIR, program) if program else BASE_DIR
        logger.info(f"Exam processor general input directory: {input_dir}")
        return input_dir
 
    # Other scripts
    input_dir = get_raw_directory(script_name)
    logger.info(f"Other script input directory: {input_dir}")
    return input_dir
def check_exam_processor_files(input_dir, program, selected_set=None):
    """Check if exam processor files exist, optionally filtering by selected set"""
    logger.info(f"Checking exam processor files in: {input_dir}")
    logger.info(f"Program: {program}, Selected Set: {selected_set}")
 
    if not os.path.isdir(input_dir):
        logger.error(f"Input directory doesn't exist: {input_dir}")
        return False
    course_dir = os.path.join(BASE_DIR, program, f"{program}-COURSES")
    course_file = None
    course_file_found = False
 
    if os.path.exists(course_dir):
        course_patterns = [
            "N-course-code-creditUnit.xlsx",
            "M-course-code-creditUnit.xlsx",
            "course-code-creditUnit.xlsx",
            "*course*.xlsx",
            "*credit*.xlsx"
        ]
     
        for pattern in course_patterns:
            for file in os.listdir(course_dir):
                if file.lower().endswith('.xlsx') and any(keyword in file.lower() for keyword in ['course', 'credit']):
                    course_file = os.path.join(course_dir, file)
                    course_file_found = True
                    logger.info(f"Found course file: {course_file}")
                    break
            if course_file:
                break
 
    if not course_file_found:
        logger.warning(f"Course file not found in: {course_dir}")
    program_dir = os.path.join(BASE_DIR, program)
    if not os.path.exists(program_dir):
        logger.error(f"Program directory not found: {program_dir}")
        return False
    valid_sets = BN_SETS if program == "BN" else (BM_SETS if program == "BM" else ND_SETS)
 
    if selected_set and selected_set != "all":
        if selected_set in valid_sets:
            program_sets = [selected_set]
            logger.info(f"Processing specific set: {selected_set}")
        else:
            logger.error(f"Invalid set selected: {selected_set}")
            return False
    else:
        program_sets = []
        for item in os.listdir(program_dir):
            item_path = os.path.join(program_dir, item)
            if os.path.isdir(item_path) and item in valid_sets:
                program_sets.append(item)
        logger.info(f"Processing all sets: {program_sets}")
    if not program_sets:
        logger.error(f"No {program} sets found in {program_dir} (valid sets: {valid_sets})")
        return False
    total_files_found = 0
    files_found = []
 
    for program_set in program_sets:
        raw_results_path = os.path.join(program_dir, program_set, "RAW_RESULTS")
        logger.info(f"Checking raw results path: {raw_results_path}")
     
        if not os.path.exists(raw_results_path):
            logger.warning(f"RAW_RESULTS not found in {raw_results_path}")
            continue
         
        files = []
        for f in os.listdir(raw_results_path):
            file_path = os.path.join(raw_results_path, f)
            if (os.path.isfile(file_path) and
                f.lower().endswith((".xlsx", ".xls")) and
                not f.startswith("~$") and
                not f.startswith(".")):
                files.append(f)
                files_found.append(f"{program_set}/{f}")
     
        total_files_found += len(files)
     
        if files:
            logger.info(f"Found {len(files)} files in {raw_results_path}: {files}")
        else:
            logger.warning(f"No Excel files found in {raw_results_path}")
    logger.info(f"Total Excel files found for {program}: {total_files_found}")
    logger.info(f"Files found: {files_found}")
 
    return total_files_found > 0
def check_putme_files(input_dir):
    if not os.path.isdir(input_dir):
        return False
    excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".xlsx", ".xls")) and "PUTME" in f.upper()]
    candidate_batches_dir = os.path.join(os.path.dirname(input_dir), "RAW_CANDIDATE_BATCHES")
    batch_files = [f for f in os.listdir(candidate_batches_dir) if f.lower().endswith(".csv") and "BATCH" in f.upper()] if os.path.isdir(candidate_batches_dir) else []
    return len(excel_files) > 0 and len(batch_files) > 0
def check_internal_exam_files(input_dir):
    """Check for internal exam files - UPDATED to recognize both Set pattern and new ND-SET pattern"""
    if not os.path.isdir(input_dir):
        logger.error(f"Directory doesn't exist: {input_dir}")
        return False
 
    # Get all valid files (same as processing script)
    csv_files = glob.glob(os.path.join(input_dir, "*.csv"))
    xls_files = glob.glob(os.path.join(input_dir, "*.xls")) + glob.glob(os.path.join(input_dir, "*.xlsx"))
    all_files = [f for f in (csv_files + xls_files) if not os.path.basename(f).startswith("~$")]
 
    # Also check for files with patterns that indicate internal exam results
    pattern_files = [
        f for f in os.listdir(input_dir)
        if f.lower().endswith((".xlsx", ".xls", ".csv"))
        and not f.startswith("~")
        and (
            f.startswith("Set") or # Original pattern
            "ND" in f.upper() and "SET" in f.upper() or # New pattern like ND2024-SET1
            "OBJ" in f.upper() or # Objective results
            "RESULT" in f.upper() # General result files
        )
    ]
 
    # Use whichever method gives us more files
    final_files = list(set(all_files + [os.path.join(input_dir, f) for f in pattern_files]))
 
    logger.info(f"Internal exam files found in {input_dir}: {len(final_files)} files")
 
    return len(final_files) > 0
def check_caosce_files(input_dir):
    if not os.path.isdir(input_dir):
        return False
    excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".xlsx", ".xls")) and "CAOSCE" in f.upper()]
    return len(excel_files) > 0
def check_split_files(input_dir):
    if not os.path.isdir(input_dir):
        return False
    valid_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".csv", ".xlsx", ".xls")) and not f.startswith("~")]
    return len(valid_files) > 0
def check_input_files(input_dir, script_name, selected_set=None):
    """Check input files with optional set filtering"""
    logger.info(f"Checking input files for {script_name} in {input_dir} (selected_set: {selected_set})")
 
    if not os.path.isdir(input_dir):
        logger.error(f"Input directory doesn't exist: {input_dir}")
        return False
     
    if script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm", "nd_carryover_processor"]:
        program = script_name.split("_")[-1].upper()
        if script_name == "nd_carryover_processor":
            program = "ND" # ND-specific processor
            if program:
                logger.info(f"Set program to {program} for carryover based on set {selected_set}")
            else:
                logger.error(f"Could not determine program for set {selected_set}")
                return False
        return check_exam_processor_files(input_dir, program, selected_set)
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
        valid_extensions = (".csv", ".xlsx", ".xls", ".pdf")
        input_files = [f for f in dir_contents if f.lower().endswith(valid_extensions) and not f.startswith("~")]
        return len(input_files) > 0
    except Exception:
        return False
def get_exam_processor_status(program, selected_set=None):
    """Get detailed status for exam processor including file counts"""
    logger.info(f"Getting status for {program}, set: {selected_set}")
 
    if program == "PROCESSOR" and selected_set:
        program = detect_program_from_set(selected_set)
        if program:
            logger.info(f"Overrode program to {program} based on set {selected_set}")
        else:
            logger.warning(f"Could not determine program from set {selected_set}")
 
    status_info = {
        'ready': False,
        'course_file': False,
        'raw_files_count': 0,
        'raw_files_list': [],
        'sets_ready': {}
    }
 
    course_dir = os.path.join(BASE_DIR, program, f"{program}-COURSES")
    if os.path.exists(course_dir):
        for file in os.listdir(course_dir):
            if file.lower().endswith('.xlsx') and any(keyword in file.lower() for keyword in ['course', 'credit']):
                status_info['course_file'] = True
                break
 
    program_dir = os.path.join(BASE_DIR, program)
    if not os.path.exists(program_dir):
        logger.error(f"Program directory not found: {program_dir}")
        return status_info
 
    valid_sets = BN_SETS if program == "BN" else (BM_SETS if program == "BM" else ND_SETS)
 
    sets_to_check = []
    if selected_set and selected_set != "all":
        if selected_set in valid_sets:
            sets_to_check = [selected_set]
    else:
        sets_to_check = valid_sets
 
    total_files = 0
    for set_name in sets_to_check:
        raw_results_path = os.path.join(program_dir, set_name, "RAW_RESULTS")
        set_files = []
     
        if os.path.exists(raw_results_path):
            for f in os.listdir(raw_results_path):
                file_path = os.path.join(raw_results_path, f)
                if (os.path.isfile(file_path) and
                    f.lower().endswith((".xlsx", ".xls")) and
                    not f.startswith("~$") and
                    not f.startswith(".")):
                    set_files.append(f)
     
        status_info['sets_ready'][set_name] = len(set_files) > 0
        status_info['raw_files_count'] += len(set_files)
        status_info['raw_files_list'].extend([f"{set_name}/{f}" for f in set_files])
        total_files += len(set_files)
 
    status_info['ready'] = status_info['course_file'] and total_files > 0
 
    logger.info(f"Status for {program}: course_file={status_info['course_file']}, raw_files={total_files}, ready={status_info['ready']}")
    return status_info
 
def count_processed_files(output_lines, script_name, selected_set=None):
    success_indicators = SUCCESS_INDICATORS.get(script_name, [])
    processed_files_set = set()
    logger.info(f"Raw output lines for {script_name}:")
    for line in output_lines:
        if line.strip():
            logger.info(f" OUTPUT: {line}")
    for line in output_lines:
        for indicator in success_indicators:
            match = re.search(indicator, line, re.IGNORECASE)
            if match:
                if script_name in ["exam_processor_bn", "exam_processor_bm", "exam_processor_nd", "nd_carryover_processor"]:
                    if "PROCESSING SEMESTER:" in line.upper() or "PROCESSING SET:" in line.upper():
                        set_name = match.group(1)
                        if set_name:
                            processed_files_set.add(f"Set: {set_name}")
                            logger.info(f"DETECTED SET: {set_name}")
                    elif "Mastersheet saved:" in line:
                        file_name = match.group(0).split("Mastersheet saved:")[-1].strip()
                        processed_files_set.add(f"Saved: {file_name}")
                    elif "Identified" in line and "carryover students" in line:
                        processed_files_set.add("Carryover students identified")
                    elif "Carryover records saved:" in line:
                        processed_files_set.add("Carryover records saved")
                    elif "Processing resit results for" in line:
                        processed_files_set.add("Resit processing started")
                    elif "Updated" in line and "scores for" in line and "students" in line:
                        processed_files_set.add("Resit scores updated")
                elif script_name == "utme":
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
                    elif "Cleaned CSV saved" in line:
                        file_name = match.group(1) if match.groups() else "cleaned_file"
                        processed_files_set.add(f"Cleaned: {file_name}")
                    elif "Master CSV saved" in line:
                        processed_files_set.add("Master file created")
                    elif "All processing completed successfully!" in line:
                        processed_files_set.add("Processing completed")
                else:
                    file_name = match.group(1) if match.groups() else line
                    processed_files_set.add(file_name)
    logger.info(f"Processed items for {script_name}: {processed_files_set}")
    return len(processed_files_set)
def get_success_message(script_name, processed_files, output_lines, selected_set=None):
    if processed_files == 0:
        return None
    if script_name == "clean":
        if any("All processing completed successfully!" in line for line in output_lines):
            return f"Successfully processed internal examination results! Generated master file and individual cleaned files."
        return f"Processed {processed_files} internal examination file(s)."
    elif script_name in ["exam_processor_nd", "exam_processor_bn", "exam_processor_bm"]:
        program = script_name.split("_")[-1].upper()
        program_name = {"ND": "ND", "BN": "Basic Nursing", "BM": "Basic Midwifery"}.get(program, program)
        upgrade_info = ""
        upgrade_count = ""
        carryover_info = ""
        resit_info = ""
        resit_updates = ""
     
        for line in output_lines:
            if "Applying upgrade rule:" in line:
                upgrade_match = re.search(r"Applying upgrade rule: (\d+)–49 → 50", line)
                if upgrade_match:
                    upgrade_info = f" Upgrade rule applied: {upgrade_match.group(1)}-49 → 50"
                    break
            elif "Upgraded" in line:
                upgrade_count_match = re.search(r"Upgraded (\d+) scores", line)
                if upgrade_count_match:
                    upgrade_count = f" Upgraded {upgrade_count_match.group(1)} scores"
                    break
            elif "Identified" in line and "carryover students" in line:
                carryover_match = re.search(r"Identified (\d+) carryover students", line)
                if carryover_match:
                    carryover_info = f" Identified {carryover_match.group(1)} carryover students"
                    break
            elif "Updated" in line and "scores for" in line and "students" in line:
                resit_match = re.search(r"Updated (\d+) scores for (\d+) students", line)
                if resit_match:
                    resit_updates = f" Resit: updated {resit_match.group(1)} scores for {resit_match.group(2)} students"
                    break
            elif "Processing resit results for" in line:
                resit_info = " Resit processing completed"
                break
     
        if resit_updates:
            return f"{program_name} Examination processing completed!{resit_updates}{upgrade_info}{upgrade_count}{carryover_info}"
        elif any(f"{program_name} Examination Results Processing completed successfully" in line for line in output_lines):
            return f"{program_name} Examination processing completed successfully! Processed {processed_files} set(s).{upgrade_info}{upgrade_count}{carryover_info}{resit_info}"
        elif any("Processing complete" in line for line in output_lines):
            return f"{program_name} Examination processing completed! Processed {processed_files} set(s).{upgrade_info}{upgrade_count}{carryover_info}{resit_info}"
        return f"Processed {processed_files} {program_name} examination set(s).{upgrade_info}{upgrade_count}{carryover_info}{resit_info}"
    elif script_name == "nd_carryover_processor":
        resit_updates = ""
        for line in output_lines:
            if "Updated" in line and "scores for" in line and "students" in line:
                resit_match = re.search(r"Updated (\d+) scores for (\d+) students", line)
                if resit_match:
                    resit_updates = f"Updated {resit_match.group(1)} scores for {resit_match.group(2)} students"
                    break
        if resit_updates:
            return f"ND carryover processing completed! {resit_updates}"
        elif any("ND CARRYOVER PROCESSING COMPLETED" in line for line in output_lines):
            return "ND carryover processing completed successfully"
        elif any("Processing resit results for" in line for line in output_lines):
            return "ND carryover processing completed but no scores updated"
        return "ND carryover processing completed"
    elif script_name == "utme":
        if any("Processing completed successfully" in line for line in output_lines):
            return f"PUTME processing completed successfully! Processed {processed_files} batch file(s)."
        return f"Processed {processed_files} PUTME batch file(s)."
    elif script_name == "caosce":
        if any("Processed" in line for line in output_lines):
            return f"CAOSCE processing completed! Processed {processed_files} file(s)."
        return f"Processed {processed_files} CAOSCE file(s)."
    elif script_name == "split":
        if any("Saved processed file:" in line for line in output_lines):
            return f"JAMB name splitting completed! Processed {processed_files} file(s)."
        return f"Processed {processed_files} JAMB file(s)."
    return f"Successfully processed {processed_files} file(s)."
def _get_script_path(script_name):
    script_path = os.path.join(SCRIPT_DIR, SCRIPT_MAP.get(script_name, ""))
    logger.info(f"Script path for {script_name}: {script_path}")
    if not os.path.exists(script_path):
        logger.error(f"Script not found: {script_path}")
        raise FileNotFoundError(f"Script {script_name} not found at {script_path}")
    return script_path
# ============================================================================
# ENHANCED: get_files_by_category - ONLY SHOW ZIP FILES
# ============================================================================
def get_files_by_category():
    """Get files organized by category - STRICTLY ONLY ZIP FILES"""
    from dataclasses import dataclass
    @dataclass
    class FileInfo:
        name: str
        relative_path: str
        folder: str
        subfolder: str
        size: int
        modified: int
        semester: str = ""
        set_name: str = ""
    files_by_category = {
        "nd_results": {},
        "bn_results": {},
        "bm_results": {},
        "putme_results": [],
        "caosce_results": [],
        "internal_results": [],
        "jamb_results": [],
    }
    # 🔒 ENFORCE ZIP-ONLY POLICY ACROSS ALL DIRECTORIES
    logger.info("🔒 Enforcing ZIP-only policy across all directories...")
 
    # Enforce for program directories
    for program in ["ND", "BN", "BM"]:
        program_dir = os.path.join(BASE_DIR, program)
        if not os.path.exists(program_dir):
            continue
     
        sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
        for set_name in sets:
            clean_dir = os.path.join(program_dir, set_name, "CLEAN_RESULTS")
            if os.path.exists(clean_dir):
                enforce_zip_only_policy(clean_dir)
    # Enforce for other result directories
    other_dirs = {
        "putme_results": os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT"),
        "caosce_results": os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT"),
        "internal_results": os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ"),
        "jamb_results": os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB"),
    }
 
    for category, dir_path in other_dirs.items():
        if os.path.exists(dir_path):
            enforce_zip_only_policy(dir_path)
    # Process program results - ONLY ZIP FILES
    for program in ["ND", "BN", "BM"]:
        program_dir = os.path.join(BASE_DIR, program)
     
        logger.info(f"📁 Scanning {program} directory: {program_dir}")
     
        if not os.path.exists(program_dir):
            logger.warning(f"⚠️ {program} directory not found")
            continue
     
        try:
            sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
         
            for set_name in sets:
                clean_dir = os.path.join(program_dir, set_name, "CLEAN_RESULTS")
             
                if not os.path.exists(clean_dir):
                    logger.debug(f"⚠️ CLEAN_RESULTS not found for {program}/{set_name}")
                    continue
             
                category = f"{program.lower()}_results"
                if set_name not in files_by_category[category]:
                    files_by_category[category][set_name] = []
             
                # 🔒 STRICT: Only include ZIP files
                try:
                    for file in os.listdir(clean_dir):
                        if not file.lower().endswith(".zip"):
                            continue
                     
                        if file.startswith("~$") or file.startswith("."):
                            continue
                     
                        file_path = os.path.join(clean_dir, file)
                     
                        try:
                            semester = extract_semester_from_filename(file)
                         
                            file_info = FileInfo(
                                name=file,
                                relative_path=os.path.relpath(file_path, BASE_DIR),
                                folder=os.path.basename(clean_dir),
                                subfolder="",
                                size=os.path.getsize(file_path),
                                modified=os.path.getmtime(file_path),
                                semester=semester,
                                set_name=set_name
                            )
                            files_by_category[category][set_name].append(file_info)
                            logger.debug(f"✅ Found ZIP: {file}")
                        except Exception as e:
                            logger.error(f"Error processing file {file}: {e}")
                            continue
                 
                    # Sort by modified time (newest first)
                    if files_by_category[category][set_name]:
                        files_by_category[category][set_name] = sorted(
                            files_by_category[category][set_name],
                            key=lambda x: x.modified,
                            reverse=True
                        )
                except Exception as e:
                    logger.error(f"Error listing files in {clean_dir}: {e}")
                 
        except Exception as e:
            logger.error(f"Error processing {program} directory: {e}")
    # Process other result types - ONLY ZIP FILES
    other_result_dirs = {
        "caosce_results": os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT"),
        "internal_results": os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ"),
        "jamb_results": os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB"),
        "putme_results": os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT")
    }
 
    for category, result_dir in other_result_dirs.items():
        logger.info(f"📁 Scanning {category} directory: {result_dir}")
     
        if not os.path.exists(result_dir):
            logger.warning(f"⚠️ Directory not found: {result_dir}")
            continue
         
        try:
            # 🔒 STRICT: Only include ZIP files
            for file in os.listdir(result_dir):
                if not file.lower().endswith('.zip'):
                    continue
             
                if file.startswith("~$") or file.startswith("."):
                    continue
             
                file_path = os.path.join(result_dir, file)
             
                try:
                    file_info = FileInfo(
                        name=file,
                        relative_path=os.path.relpath(file_path, BASE_DIR),
                        folder=os.path.basename(result_dir),
                        subfolder="",
                        size=os.path.getsize(file_path),
                        modified=os.path.getmtime(file_path)
                    )
                    files_by_category[category].append(file_info)
                    logger.debug(f"✅ Found {category} ZIP: {file}")
                except Exception as e:
                    logger.error(f"Error processing {category} file {file}: {e}")
                    continue
         
            # Sort by modified time (newest first)
            if files_by_category[category]:
                files_by_category[category] = sorted(
                    files_by_category[category],
                    key=lambda x: x.modified,
                    reverse=True
                )
             
        except Exception as e:
            logger.error(f"Error scanning {result_dir}: {e}")
    # Log summary - ONLY ZIP FILES
    logger.info("="*60)
    logger.info("📊 DOWNLOAD CENTER FILE SUMMARY (ZIP FILES ONLY):")
    logger.info("="*60)
    for category, files in files_by_category.items():
        if isinstance(files, dict):
            total = sum(len(f) for f in files.values())
            logger.info(f"{category}: {total} ZIP files across {len(files)} sets")
        else:
            logger.info(f"{category}: {len(files)} ZIP files")
    logger.info("="*60)
 
    return files_by_category
# ============================================================================
# NEW: Manual ZIP Creation Route
# ============================================================================
@app.route("/create_zip_manually/<category>/<set_name>")
@login_required
def create_zip_manually(category, set_name):
    """Manually create ZIP file for a category/set"""
    try:
        if category in ["nd_results", "bn_results", "bm_results"]:
            program = category.split('_')[0].upper()
            clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        else:
            dir_map = {
                "putme_results": os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT"),
                "caosce_results": os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT"),
                "internal_results": os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ"),
                "jamb_results": os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB"),
            }
            clean_dir = dir_map.get(category)
     
        if clean_dir and os.path.exists(clean_dir):
            if enforce_zip_only_policy(clean_dir):
                flash(f"✅ Successfully created ZIP file for {set_name or category}", "success")
            else:
                flash(f"❌ Failed to create ZIP file for {set_name or category}", "error")
        else:
            flash(f"❌ Directory not found for {set_name or category}", "error")
     
        return redirect(url_for("download_center"))
     
    except Exception as e:
        logger.error(f"Manual ZIP creation error: {e}")
        flash(f"Error creating ZIP: {str(e)}", "error")
        return redirect(url_for("download_center"))
# ============================================================================
# NEW: Route to Create Missing ZIP Files
# ============================================================================
@app.route("/create_missing_zips")
@login_required
def create_missing_zips_route():
    """Route to manually create missing ZIP files"""
    try:
        created_zips = create_missing_zips()
       
        if created_zips:
            flash(f"✅ Created {len(created_zips)} missing ZIP files: {', '.join(created_zips)}", "success")
        else:
            flash("ℹ️ No missing ZIP files needed to be created", "info")
           
        return redirect(url_for("download_center"))
       
    except Exception as e:
        logger.error(f"Error in create_missing_zips route: {e}")
        flash(f"Error creating missing ZIPs: {str(e)}", "error")
        return redirect(url_for("download_center"))
# ============================================================================
# NEW: Debug route for BM files
# ============================================================================
@app.route("/debug_bm_files")
@login_required
def debug_bm_files():
    """Debug route to check BM files"""
    debug_info = {
        "bm_directories": {},
        "all_files_found": []
    }
 
    for set_name in BM_SETS:
        set_dir = os.path.join(BASE_DIR, "BM", set_name)
        debug_info["bm_directories"][set_name] = {
            "exists": os.path.exists(set_dir),
            "clean_results_exists": False,
            "clean_results_files": []
        }
     
        if os.path.exists(set_dir):
            clean_dir = os.path.join(set_dir, "CLEAN_RESULTS")
            debug_info["bm_directories"][set_name]["clean_results_exists"] = os.path.exists(clean_dir)
         
            if os.path.exists(clean_dir):
                files = os.listdir(clean_dir)
                debug_info["bm_directories"][set_name]["clean_results_files"] = files
                debug_info["all_files_found"].extend([f"{set_name}/{f}" for f in files])
 
    return jsonify(debug_info)
# ============================================================================
# ENHANCED: get_sets_and_folders - ONLY ZIP FILES
# ============================================================================
def get_sets_and_folders():
    """Get sets and their result folders from CLEAN_RESULTS directories - ONLY ZIP FILES"""
    from dataclasses import dataclass
    @dataclass
    class FileInfo:
        name: str
        relative_path: str
        size: int
        modified: int
    @dataclass
    class FolderInfo:
        name: str
        files: list
    sets = {}
    logger.info(f"Scanning BASE_DIR for ZIP files only: {BASE_DIR}")
    # First ensure all results are properly zipped
    for program in ["ND", "BN", "BM"]:
        program_dir = os.path.join(BASE_DIR, program)
        if not os.path.exists(program_dir):
            continue
         
        valid_sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
        for set_name in valid_sets:
            clean_dir = os.path.join(program_dir, set_name, "CLEAN_RESULTS")
            if os.path.exists(clean_dir):
                ensure_zipped_results(clean_dir, f"exam_processor_{program.lower()}", set_name)
    for program in ["ND", "BN", "BM"]:
        program_dir = os.path.join(BASE_DIR, program)
        if not os.path.exists(program_dir):
            logger.warning(f"Program directory not found: {program_dir}")
            continue
         
        logger.info(f"Scanning program: {program_dir}")
        valid_sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
     
        for set_name in os.listdir(program_dir):
            set_path = os.path.join(program_dir, set_name)
            if not os.path.isdir(set_path) or set_name not in valid_sets:
                logger.warning(f"Skipping invalid set {set_name} for program {program}")
                continue
             
            logger.info(f"Scanning set: {set_path}")
            folders = []
            clean_results_path = os.path.join(set_path, "CLEAN_RESULTS")
         
            if os.path.exists(clean_results_path):
                logger.info(f"Found CLEAN_RESULTS: {clean_results_path}")
                try:
                    # Only show ZIP files in file browser
                    zip_files = []
                    for file in os.listdir(clean_results_path):
                        if file.lower().endswith(".zip"):
                            file_path = os.path.join(clean_results_path, file)
                            try:
                                relative_path = os.path.relpath(file_path, BASE_DIR)
                                zip_files.append(FileInfo(
                                    name=file,
                                    relative_path=relative_path,
                                    size=os.path.getsize(file_path),
                                    modified=os.path.getmtime(file_path)
                                ))
                                logger.info(f"Found ZIP file: {file} in {clean_results_path}")
                            except Exception as e:
                                logger.error(f"Error processing file {file}: {e}")
                 
                    if zip_files:
                        folders.append(FolderInfo(name="ZIP Results", files=zip_files))
                        logger.info(f"Added ZIP files folder with {len(zip_files)} files")
                except Exception as e:
                    logger.error(f"Error scanning {clean_results_path}: {e}")
                 
            if folders:
                set_key = f"{program}_{set_name}"
                sets[set_key] = folders
                logger.info(f"Added set {set_key} with {len(folders)} folders")
    # Other result types - only show ZIP files
    result_mappings = {
        "PUTME_RESULT": {
            "clean_dir": "CLEAN_PUTME_RESULT",
            "base_dir": "PUTME_RESULT"
        },
        "CAOSCE_RESULT": {
            "clean_dir": "CLEAN_CAOSCE_RESULT",
            "base_dir": "CAOSCE_RESULT"
        },
        "INTERNAL_RESULT": {
            "clean_dir": "CLEAN_OBJ",
            "base_dir": "OBJ_RESULT"
        },
        "JAMB_DB": {
            "clean_dir": "CLEAN_JAMB_DB",
            "base_dir": "JAMB_DB"
        }
    }
 
    for result_type, mapping in result_mappings.items():
        clean_dir_name = mapping["clean_dir"]
        base_dir_name = mapping["base_dir"]
     
        result_path = os.path.join(BASE_DIR, base_dir_name, clean_dir_name)
        logger.info(f"Scanning {result_type} at: {result_path}")
     
        if not os.path.exists(result_path):
            logger.warning(f"Directory not found: {result_path}")
            continue
         
        # Ensure results are zipped
        ensure_zipped_results(result_path, "other_processor")
         
        folders = []
        try:
            # Only show ZIP files
            zip_files = []
            for file in os.listdir(result_path):
                if file.lower().endswith(".zip"):
                    file_path = os.path.join(result_path, file)
                    try:
                        relative_path = os.path.relpath(file_path, BASE_DIR)
                        zip_files.append(FileInfo(
                            name=file,
                            relative_path=relative_path,
                            size=os.path.getsize(file_path),
                            modified=os.path.getmtime(file_path)
                        ))
                        logger.info(f"Found ZIP file: {file} in {result_path}")
                    except Exception as e:
                        logger.error(f"Error processing file {file}: {e}")
         
            if zip_files:
                folders.append(FolderInfo(name="ZIP Results", files=zip_files))
                logger.info(f"Added ZIP files folder with {len(zip_files)} files")
                 
        except Exception as e:
            logger.error(f"Error scanning {result_path}: {e}")
         
        if folders:
            sets[result_type] = folders
            logger.info(f"Added result type {result_type} with {len(folders)} folders")
        else:
            logger.info(f"No ZIP folders found for {result_type}")
    logger.info(f"Total sets found with ZIP files: {len(sets)}")
    for set_name, folders in sets.items():
        total_files = sum(len(folder.files) for folder in folders)
        logger.info(f"{set_name}: {total_files} ZIP files in {len(folders)} folders")
     
    return sets
# ============================================================================
# UPDATED: Helper Functions
# ============================================================================
def get_available_sets(program):
    """Get available academic sets for a program"""
    if program == "BN":
        return BN_SETS
    elif program == "BM":
        return BM_SETS
    elif program == "ND":
        return ND_SETS
    return []
# ============================================================================
# FIX C: UPDATED: get_carryover_records function - ND-SPECIFIC VERSION
# ============================================================================
def get_nd_carryover_records(set_name, semester_key=None):
    """Get carryover records for ND set - CHECK INSIDE RESULT FOLDERS FIRST."""
    try:
        if semester_key:
            semester_key = standardize_semester_key_nd(semester_key)
            logger.info(f"🔑 Standardized semester key: {semester_key}")
     
        # Direct path to CLEAN_RESULTS
        clean_dir = os.path.join(BASE_DIR, "ND", set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            logger.warning(f"❌ CLEAN_RESULTS not found: {clean_dir}")
            return []
     
        # Check centralized CARRYOVER_RECORDS first (NEW structure)
        carryover_dir = os.path.join(clean_dir, "CARRYOVER_RECORDS")
        if os.path.exists(carryover_dir):
            logger.info(f"✅ Found centralized CARRYOVER_RECORDS: {carryover_dir}")
            return load_carryover_json_files(carryover_dir, semester_key, "ND")
     
        # Fallback: Check inside latest result folder/ZIP (OLD structure)
        logger.info(f"⚠️ No centralized folder, checking results...")
     
        # Get all result items (folders and ZIPs)
        result_items = []
        for item in os.listdir(clean_dir):
            if item.startswith(f"{set_name}_RESULT-") and not item.startswith("CARRYOVER_"):
                result_items.append(item)
     
        if not result_items:
            logger.warning(f"❌ No result files found in: {clean_dir}")
            return []
     
        # Use latest result
        latest_item = sorted(result_items)[-1]
        latest_path = os.path.join(clean_dir, latest_item)
     
        logger.info(f"✅ Using latest result: {latest_item}")
     
        if latest_item.endswith('.zip'):
            return get_carryover_records_from_zip(latest_path, set_name, semester_key, "ND")
        else:
            # It's a directory
            carryover_dir = os.path.join(latest_path, "CARRYOVER_RECORDS")
            if os.path.exists(carryover_dir):
                return load_carryover_json_files(carryover_dir, semester_key, "ND")
            else:
                logger.warning(f"❌ No CARRYOVER_RECORDS in: {latest_path}")
                return []
             
    except Exception as e:
        logger.error(f"❌ Error getting carryover records: {e}")
        import traceback
        traceback.print_exc()
        return []
def get_nd_carryover_records_from_zip(zip_path, semester_key=None):
    """Extract carryover records from ZIP file for ND"""
    try:
        logger.info(f"📦 Extracting ND carryover records from ZIP: {zip_path}")
        carryover_files = []
     
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # List all files in ZIP for debugging
            all_files = zip_ref.namelist()
            logger.info(f"📁 Files in ZIP: {len(all_files)} files")
         
            # Look for carryover JSON files
            json_files = []
            for f in all_files:
                if f.endswith('.json') and ('CARRYOVER' in f.upper() or 'CO_STUDENT' in f.upper()):
                    json_files.append(f)
         
            logger.info(f"📁 Found {len(json_files)} potential carryover JSON files")
         
            if not json_files:
                logger.info(f"❌ No carryover JSON files found in ZIP")
                return []
             
            for json_file in json_files:
                file_semester = extract_semester_from_filename(json_file)
                file_semester_standardized = standardize_semester_key_nd(file_semester)
             
                if semester_key and file_semester_standardized != semester_key:
                    logger.info(f" ⏭️ Skipping {json_file} (doesn't match target {semester_key})")
                    continue
                 
                try:
                    with zip_ref.open(json_file) as f:
                        data = json.load(f)
                        carryover_files.append({
                            'filename': os.path.basename(json_file),
                            'semester': file_semester_standardized,
                            'data': data,
                            'count': len(data),
                            'file_path': f"{zip_path}/{json_file}"
                        })
                        logger.info(f"✅ Loaded carryover record: {json_file} ({len(data)} students)")
                except Exception as e:
                    logger.error(f"Error loading carryover file {json_file}: {e}")
     
        logger.info(f"✅ Loaded {len(carryover_files)} carryover records from ZIP")
        return carryover_files
     
    except Exception as e:
        logger.error(f"Error extracting carryover records from ZIP: {e}")
        return []
def load_nd_carryover_json_files(carryover_dir, semester_key=None):
    """Load carryover JSON files from directory for ND"""
    carryover_files = []
 
    # Standardize the target semester key
    if semester_key:
        semester_key = standardize_semester_key_nd(semester_key)
 
    for file in os.listdir(carryover_dir):
        if file.startswith("co_student_") and file.endswith(".json"):
            # Extract semester from filename and standardize it
            file_semester = extract_semester_from_filename(file)
            file_semester_standardized = standardize_semester_key_nd(file_semester)
         
            logger.info(f"📄 Found ND carryover file: {file}")
            logger.info(f" Original semester: {file_semester}")
            logger.info(f" Standardized: {file_semester_standardized}")
            logger.info(f" Target semester: {semester_key}")
         
            # If semester_key is specified, only load matching files
            if semester_key and file_semester_standardized != semester_key:
                logger.info(f" ⏭️ Skipping (doesn't match target)")
                continue
         
            file_path = os.path.join(carryover_dir, file)
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f)
                    carryover_files.append({
                        'filename': file,
                        'semester': file_semester_standardized, # Use standardized key
                        'data': data,
                        'count': len(data),
                        'file_path': file_path
                    })
                    logger.info(f" ✅ Loaded: {len(data)} records")
            except Exception as e:
                logger.error(f"Error loading {file}: {e}")
 
    logger.info(f"📊 Total ND carryover files loaded: {len(carryover_files)}")
    return carryover_files
# Original function kept for BN/BM compatibility
def get_carryover_records(program, set_name, semester_key=None):
    """Get carryover records - SIMPLIFIED PATH LOGIC."""
    try:
        if semester_key:
            semester_key = standardize_semester_key(semester_key, program)
            logger.info(f"🔑 Standardized semester key: {semester_key}")
     
        # Direct path to CLEAN_RESULTS
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            logger.warning(f"❌ CLEAN_RESULTS not found: {clean_dir}")
            return []
     
        # Check centralized CARRYOVER_RECORDS first (NEW structure)
        carryover_dir = os.path.join(clean_dir, "CARRYOVER_RECORDS")
        if os.path.exists(carryover_dir):
            logger.info(f"✅ Found centralized CARRYOVER_RECORDS: {carryover_dir}")
            return load_carryover_json_files(carryover_dir, semester_key, program)
     
        # Fallback: Check inside latest result folder/ZIP (OLD structure)
        logger.info(f"⚠️ No centralized folder, checking results...")
     
        # Get all result items (folders and ZIPs)
        result_items = []
        for item in os.listdir(clean_dir):
            if item.startswith(f"{set_name}_RESULT-") and not item.startswith("CARRYOVER_"):
                result_items.append(item)
     
        if not result_items:
            logger.warning(f"❌ No result files found in: {clean_dir}")
            return []
     
        # Use latest result
        latest_item = sorted(result_items)[-1]
        latest_path = os.path.join(clean_dir, latest_item)
     
        logger.info(f"✅ Using latest result: {latest_item}")
     
        if latest_item.endswith('.zip'):
            return get_carryover_records_from_zip(latest_path, set_name, semester_key, program)
        else:
            # It's a directory
            carryover_dir = os.path.join(latest_path, "CARRYOVER_RECORDS")
            if os.path.exists(carryover_dir):
                return load_carryover_json_files(carryover_dir, semester_key, program)
            else:
                logger.warning(f"❌ No CARRYOVER_RECORDS in: {latest_path}")
                return []
             
    except Exception as e:
        logger.error(f"❌ Error getting carryover records: {e}")
        import traceback
        traceback.print_exc()
        return []
def process_carryover_results(program, set_name, semester, file_path):
    """Process carryover results for a program"""
    try:
        # Your existing carryover processing logic here
        # This should integrate with your existing carryover processing system
     
        return {
            "success": True,
            "message": f"Processed {program} carryover results for {set_name}, {semester}"
        }
    except Exception as e:
        logger.error(f"Carryover processing error for {program}: {e}")
        return {
            "success": False,
            "message": str(e)
        }
# ============================================================================
# FIXED: get_carryover_records_from_zip function - UPDATED VERSION
# ============================================================================
def get_carryover_records_from_zip(zip_path, set_name, semester_key=None, program=None):
    """Extract carryover records from ZIP file - FIXED FOR BN"""
    try:
        logger.info(f"📦 Extracting carryover records from ZIP: {zip_path}")
        carryover_files = []
     
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # List all files in ZIP for debugging
            all_files = zip_ref.namelist()
            logger.info(f"📁 Files in ZIP: {len(all_files)} files")
         
            # Look for carryover JSON files - FIXED: Also check in CARRYOVER_* directories
            json_files = []
            for f in all_files:
                if f.endswith('.json') and ('CARRYOVER' in f.upper() or 'CO_STUDENT' in f.upper()):
                    json_files.append(f)
         
            logger.info(f"📁 Found {len(json_files)} potential carryover JSON files")
         
            if not json_files:
                logger.info(f"❌ No carryover JSON files found in ZIP")
                return []
             
            for json_file in json_files:
                file_semester = extract_semester_from_filename(json_file)
                file_semester_standardized = standardize_semester_key(file_semester, program)
             
                if semester_key and file_semester_standardized != semester_key:
                    logger.info(f" ⏭️ Skipping {json_file} (doesn't match target {semester_key})")
                    continue
                 
                try:
                    with zip_ref.open(json_file) as f:
                        data = json.load(f)
                        carryover_files.append({
                            'filename': os.path.basename(json_file),
                            'semester': file_semester_standardized,
                            'data': data,
                            'count': len(data),
                            'file_path': f"{zip_path}/{json_file}"
                        })
                        logger.info(f"✅ Loaded carryover record: {json_file} ({len(data)} students)")
                except Exception as e:
                    logger.error(f"Error loading carryover file {json_file}: {e}")
     
        logger.info(f"✅ Loaded {len(carryover_files)} carryover records from ZIP")
        return carryover_files
     
    except Exception as e:
        logger.error(f"Error extracting carryover records from ZIP: {e}")
        return []
# ============================================================================
# FIXED: load_carryover_json_files function - UPDATED VERSION
# ============================================================================
def load_carryover_json_files(carryover_dir, semester_key=None, program=None):
    """Load carryover JSON files from directory - FIXED."""
    carryover_files = []
 
    # Standardize the target semester key
    if semester_key:
        semester_key = standardize_semester_key(semester_key, program)
 
    for file in os.listdir(carryover_dir):
        if file.startswith("co_student_") and file.endswith(".json"):
            # Extract semester from filename and standardize it
            file_semester = extract_semester_from_filename(file)
            file_semester_standardized = standardize_semester_key(file_semester, program)
         
            logger.info(f"📄 Found carryover file: {file}")
            logger.info(f" Original semester: {file_semester}")
            logger.info(f" Standardized: {file_semester_standardized}")
            logger.info(f" Target semester: {semester_key}")
         
            # If semester_key is specified, only load matching files
            if semester_key and file_semester_standardized != semester_key:
                logger.info(f" ⏭️ Skipping (doesn't match target)")
                continue
         
            file_path = os.path.join(carryover_dir, file)
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f)
                    carryover_files.append({
                        'filename': file,
                        'semester': file_semester_standardized, # Use standardized key
                        'data': data,
                        'count': len(data),
                        'file_path': file_path
                    })
                    logger.info(f" ✅ Loaded: {len(data)} records")
            except Exception as e:
                logger.error(f"Error loading {file}: {e}")
 
    logger.info(f"📊 Total carryover files loaded: {len(carryover_files)}")
    return carryover_files
# ============================================================================
# FIXED: extract_semester_from_filename function - ENHANCED VERSION WITH GENERAL PATTERNS
# ============================================================================
def extract_semester_from_filename(filename):
    """Extract semester from filename using comprehensive pattern matching - FIXED."""
    filename_upper = filename.upper()
 
    # Remove possible prefix first to normalize
    filename_upper = re.sub(r'^(N|M|ND|BN|BM|SET|YEAR|SEMESTER)[-_]?', '', filename_upper)
 
    # Initial regex match (prefix optional)
    semester_pattern = r'((?:N|M|ND)?[-_]?(?:FIRST|SECOND|THIRD|1ST|2ND|3RD)[-_]?YEAR[-_]?(?:FIRST|SECOND|1ST|2ND)[-_]?SEMESTER)'
    match = re.search(semester_pattern, filename_upper, re.IGNORECASE)
 
    if match:
        extracted = match.group(1)
        # Remove archive prefix from extracted
        extracted = re.sub(r'^(N|M|ND|BN|BM)[-_]?', '', extracted)
        standardized = standardize_semester_key(extracted)
        logger.info(f"✅ Extracted and standardized: '{filename}' → '{standardized}' (regex match)")
        return standardized
 
    # Fallback to comprehensive pattern matching without prefix
    semester_patterns = {
        "FIRST-YEAR-FIRST-SEMESTER": [
            # General without prefix
            "FIRST.YEAR.FIRST.SEMESTER", "FIRST-YEAR-FIRST-SEMESTER", "FIRST_YEAR_FIRST_SEMESTER",
            "1ST.YEAR.1ST.SEMESTER", "1ST-YEAR-1ST-SEMESTER", "1ST_YEAR_1ST_SEMESTER",
            "YEAR1.SEMESTER1", "YEAR-1-SEMESTER-1", "YEAR_1_SEMESTER_1",
            "FIRST.SEMESTER.FIRST.YEAR", "1ST.SEMESTER.1ST.YEAR",
            # With N prefix
            "N.FIRST.YEAR.FIRST.SEMESTER", "N-FIRST-YEAR-FIRST-SEMESTER", "N_FIRST_YEAR_FIRST_SEMESTER",
            "N1ST.YEAR1ST.SEMESTER", "N1ST-YEAR-1ST-SEMESTER", "N1ST_YEAR_1ST_SEMESTER",
            "N YEAR1.SEMESTER1", "N-YEAR-1-SEMESTER-1", "N_YEAR_1_SEMESTER_1",
            "BN.FIRST.YEAR.FIRST.SEMESTER", "BN-FIRST-YEAR-FIRST-SEMESTER", "BN_FIRST_YEAR_FIRST_SEMESTER",
            # With M prefix
            "M.FIRST.YEAR.FIRST.SEMESTER", "M-FIRST-YEAR-FIRST-SEMESTER", "M_FIRST_YEAR_FIRST_SEMESTER",
            "M1ST.YEAR1ST.SEMESTER", "M1ST-YEAR-1ST-SEMESTER", "M1ST_YEAR_1ST_SEMESTER",
            "M YEAR1.SEMESTER1", "M-YEAR-1-SEMESTER-1", "M_YEAR_1_SEMESTER_1",
            "BM.FIRST.YEAR.FIRST.SEMESTER", "BM-FIRST-YEAR-FIRST-SEMESTER", "BM_FIRST_YEAR_FIRST_SEMESTER",
            # With ND prefix
            "ND.FIRST.YEAR.FIRST.SEMESTER", "ND-FIRST-YEAR-FIRST-SEMESTER", "ND_FIRST_YEAR_FIRST_SEMESTER",
            "ND1ST.YEAR1ST.SEMESTER", "ND1ST-YEAR-1ST-SEMESTER", "ND1ST_YEAR_1ST_SEMESTER",
            "ND YEAR1.SEMESTER1", "ND-YEAR-1-SEMESTER-1", "ND_YEAR_1_SEMESTER_1"
        ],
        "FIRST-YEAR-SECOND-SEMESTER": [
            # General without prefix
            "FIRST.YEAR.SECOND.SEMESTER", "FIRST-YEAR-SECOND-SEMESTER", "FIRST_YEAR_SECOND_SEMESTER",
            "1ST.YEAR.2ND.SEMESTER", "1ST-YEAR-2ND-SEMESTER", "1ST_YEAR_2ND_SEMESTER",
            "YEAR1.SEMESTER2", "YEAR-1-SEMESTER-2", "YEAR_1_SEMESTER_2",
            "SECOND.SEMESTER.FIRST.YEAR", "2ND.SEMESTER.1ST.YEAR",
            # With N prefix
            "N.FIRST.YEAR.SECOND.SEMESTER", "N-FIRST-YEAR-SECOND-SEMESTER", "N_FIRST_YEAR_SECOND_SEMESTER",
            "N1ST.YEAR2ND.SEMESTER", "N1ST-YEAR-2ND-SEMESTER", "N1ST_YEAR_2ND_SEMESTER",
            "N YEAR1.SEMESTER2", "N-YEAR-1-SEMESTER-2", "N_YEAR_1_SEMESTER_2",
            "BN.FIRST.YEAR.SECOND.SEMESTER", "BN-FIRST-YEAR-SECOND-SEMESTER", "BN_FIRST_YEAR_SECOND_SEMESTER",
            # With M prefix
            "M.FIRST.YEAR.SECOND.SEMESTER", "M-FIRST-YEAR-SECOND-SEMESTER", "M_FIRST_YEAR_SECOND_SEMESTER",
            "M1ST.YEAR2ND.SEMESTER", "M1ST-YEAR-2ND-SEMESTER", "M1ST_YEAR_2ND_SEMESTER",
            "M YEAR1.SEMESTER2", "M-YEAR-1-SEMESTER-2", "M_YEAR_1_SEMESTER_2",
            "BM.FIRST.YEAR.SECOND.SEMESTER", "BM-FIRST-YEAR-SECOND-SEMESTER", "BM_FIRST_YEAR_SECOND_SEMESTER",
            # With ND prefix
            "ND.FIRST.YEAR.SECOND.SEMESTER", "ND-FIRST-YEAR-SECOND-SEMESTER", "ND_FIRST_YEAR_SECOND_SEMESTER",
            "ND1ST.YEAR2ND.SEMESTER", "ND1ST-YEAR-2ND-SEMESTER", "ND1ST_YEAR_2ND_SEMESTER",
            "ND YEAR1.SEMESTER2", "ND-YEAR-1-SEMESTER-2", "ND_YEAR_1_SEMESTER_2"
        ],
        "SECOND-YEAR-FIRST-SEMESTER": [
            # General without prefix
            "SECOND.YEAR.FIRST.SEMESTER", "SECOND-YEAR-FIRST-SEMESTER", "SECOND_YEAR_FIRST_SEMESTER",
            "2ND.YEAR.1ST.SEMESTER", "2ND-YEAR-1ST-SEMESTER", "2ND_YEAR_1ST_SEMESTER",
            "YEAR2.SEMESTER1", "YEAR-2-SEMESTER-1", "YEAR_2_SEMESTER_1",
            "FIRST.SEMESTER.SECOND.YEAR", "1ST.SEMESTER.2ND.YEAR",
            # With N prefix
            "N.SECOND.YEAR.FIRST.SEMESTER", "N-SECOND-YEAR-FIRST-SEMESTER", "N_SECOND_YEAR_FIRST_SEMESTER",
            "N2ND.YEAR1ST.SEMESTER", "N2ND-YEAR-1ST-SEMESTER", "N2ND_YEAR_1ST_SEMESTER",
            "N YEAR2.SEMESTER1", "N-YEAR-2-SEMESTER-1", "N_YEAR_2_SEMESTER_1",
            "BN.SECOND.YEAR.FIRST.SEMESTER", "BN-SECOND-YEAR-FIRST-SEMESTER", "BN_SECOND_YEAR_FIRST_SEMESTER",
            # With M prefix
            "M.SECOND.YEAR.FIRST.SEMESTER", "M-SECOND-YEAR-FIRST-SEMESTER", "M_SECOND_YEAR_FIRST_SEMESTER",
            "M2ND.YEAR1ST.SEMESTER", "M2ND-YEAR-1ST-SEMESTER", "M2ND_YEAR_1ST_SEMESTER",
            "M YEAR2.SEMESTER1", "M-YEAR-2-SEMESTER-1", "M_YEAR_2_SEMESTER_1",
            "BM.SECOND.YEAR.FIRST.SEMESTER", "BM-SECOND-YEAR-FIRST-SEMESTER", "BM_SECOND_YEAR_FIRST_SEMESTER",
            # With ND prefix
            "ND.SECOND.YEAR.FIRST.SEMESTER", "ND-SECOND-YEAR-FIRST-SEMESTER", "ND_SECOND_YEAR_FIRST_SEMESTER",
            "ND2ND.YEAR1ST.SEMESTER", "ND2ND-YEAR-1ST-SEMESTER", "ND2ND_YEAR_1ST_SEMESTER",
            "ND YEAR2.SEMESTER1", "ND-YEAR-2-SEMESTER-1", "ND_YEAR_2_SEMESTER_1"
        ],
        # Add similar blocks for all other semesters: SECOND-YEAR-SECOND-SEMESTER, THIRD-YEAR-FIRST-SEMESTER, THIRD-YEAR-SECOND-SEMESTER
        "SECOND-YEAR-SECOND-SEMESTER": [
            # General
            "SECOND.YEAR.SECOND.SEMESTER", "SECOND-YEAR-SECOND-SEMESTER", "SECOND_YEAR_SECOND_SEMESTER",
            "2ND.YEAR.2ND.SEMESTER", "2ND-YEAR-2ND-SEMESTER", "2ND_YEAR_2ND_SEMESTER",
            "YEAR2.SEMESTER2", "YEAR-2-SEMESTER-2", "YEAR_2_SEMESTER_2",
            "SECOND.SEMESTER.SECOND.YEAR", "2ND.SEMESTER.2ND.YEAR",
            # N
            "N.SECOND.YEAR.SECOND.SEMESTER", "N-SECOND-YEAR-SECOND-SEMESTER", "N_SECOND_YEAR_SECOND_SEMESTER",
            "N2ND.YEAR2ND.SEMESTER", "N2ND-YEAR-2ND-SEMESTER", "N2ND_YEAR_2ND_SEMESTER",
            "N YEAR2.SEMESTER2", "N-YEAR-2-SEMESTER-2", "N_YEAR_2_SEMESTER_2",
            "BN.SECOND.YEAR.SECOND.SEMESTER", "BN-SECOND-YEAR-SECOND-SEMESTER", "BN_SECOND_YEAR_SECOND_SEMESTER",
            # M
            "M.SECOND.YEAR.SECOND.SEMESTER", "M-SECOND-YEAR-SECOND-SEMESTER", "M_SECOND_YEAR_SECOND_SEMESTER",
            "M2ND.YEAR2ND.SEMESTER", "M2ND-YEAR-2ND-SEMESTER", "M2ND_YEAR_2ND_SEMESTER",
            "M YEAR2.SEMESTER2", "M-YEAR-2-SEMESTER-2", "M_YEAR_2_SEMESTER_2",
            "BM.SECOND.YEAR.SECOND.SEMESTER", "BM-SECOND-YEAR-SECOND-SEMESTER", "BM_SECOND_YEAR_SECOND_SEMESTER",
            # ND
            "ND.SECOND.YEAR.SECOND.SEMESTER", "ND-SECOND-YEAR-SECOND-SEMESTER", "ND_SECOND_YEAR_SECOND_SEMESTER",
            "ND2ND.YEAR2ND.SEMESTER", "ND2ND-YEAR-2ND-SEMESTER", "ND2ND_YEAR_2ND_SEMESTER",
            "ND YEAR2.SEMESTER2", "ND-YEAR-2-SEMESTER-2", "ND_YEAR_2_SEMESTER_2"
        ],
        "THIRD-YEAR-FIRST-SEMESTER": [
            # General
            "THIRD.YEAR.FIRST.SEMESTER", "THIRD-YEAR-FIRST-SEMESTER", "THIRD_YEAR_FIRST_SEMESTER",
            "3RD.YEAR.1ST.SEMESTER", "3RD-YEAR-1ST-SEMESTER", "3RD_YEAR_1ST_SEMESTER",
            "YEAR3.SEMESTER1", "YEAR-3-SEMESTER-1", "YEAR_3_SEMESTER_1",
            # N
            "N.THIRD.YEAR.FIRST.SEMESTER", "N-THIRD-YEAR-FIRST-SEMESTER", "N_THIRD_YEAR_FIRST_SEMESTER",
            "N3RD.YEAR1ST.SEMESTER", "N3RD-YEAR-1ST-SEMESTER", "N3RD_YEAR_1ST_SEMESTER",
            "N YEAR3.SEMESTER1", "N-YEAR-3-SEMESTER-1", "N_YEAR_3_SEMESTER_1",
            "BN.THIRD.YEAR.FIRST.SEMESTER", "BN-THIRD-YEAR-FIRST-SEMESTER", "BN_THIRD_YEAR_FIRST_SEMESTER",
            # M
            "M.THIRD.YEAR.FIRST.SEMESTER", "M-THIRD-YEAR-FIRST-SEMESTER", "M_THIRD_YEAR_FIRST_SEMESTER",
            "M3RD.YEAR1ST.SEMESTER", "M3RD-YEAR-1ST-SEMESTER", "M3RD_YEAR_1ST_SEMESTER",
            "M YEAR3.SEMESTER1", "M-YEAR-3-SEMESTER-1", "M_YEAR_3_SEMESTER_1",
            "BM.THIRD.YEAR.FIRST.SEMESTER", "BM-THIRD-YEAR-FIRST-SEMESTER", "BM_THIRD_YEAR_FIRST_SEMESTER",
            # ND (though ND may not have third year, include for completeness)
            "ND.THIRD.YEAR.FIRST.SEMESTER", "ND-THIRD-YEAR-FIRST-SEMESTER", "ND_THIRD_YEAR_FIRST_SEMESTER",
            "ND3RD.YEAR1ST.SEMESTER", "ND3RD-YEAR-1ST-SEMESTER", "ND3RD_YEAR_1ST_SEMESTER",
            "ND YEAR3.SEMESTER1", "ND-YEAR-3-SEMESTER-1", "ND_YEAR_3_SEMESTER_1"
        ],
        "THIRD-YEAR-SECOND-SEMESTER": [
            # General
            "THIRD.YEAR.SECOND.SEMESTER", "THIRD-YEAR-SECOND-SEMESTER", "THIRD_YEAR_SECOND_SEMESTER",
            "3RD.YEAR.2ND.SEMESTER", "3RD-YEAR-2ND-SEMESTER", "3RD_YEAR_2ND_SEMESTER",
            "YEAR3.SEMESTER2", "YEAR-3-SEMESTER-2", "YEAR_3_SEMESTER_2",
            # N
            "N.THIRD.YEAR.SECOND.SEMESTER", "N-THIRD-YEAR-SECOND-SEMESTER", "N_THIRD_YEAR_SECOND_SEMESTER",
            "N3RD.YEAR2ND.SEMESTER", "N3RD-YEAR-2ND-SEMESTER", "N3RD_YEAR_2ND_SEMESTER",
            "N YEAR3.SEMESTER2", "N-YEAR-3-SEMESTER-2", "N_YEAR_3_SEMESTER_2",
            "BN.THIRD.YEAR.SECOND.SEMESTER", "BN-THIRD-YEAR-SECOND-SEMESTER", "BN_THIRD_YEAR_SECOND_SEMESTER",
            # M
            "M.THIRD.YEAR.SECOND.SEMESTER", "M-THIRD-YEAR-SECOND-SEMESTER", "M_THIRD_YEAR_SECOND_SEMESTER",
            "M3RD.YEAR2ND.SEMESTER", "M3RD-YEAR-2ND-SEMESTER", "M3RD_YEAR_2ND_SEMESTER",
            "M YEAR3.SEMESTER2", "M-YEAR-3-SEMESTER-2", "M_YEAR_3_SEMESTER_2",
            "BM.THIRD.YEAR.SECOND.SEMESTER", "BM-THIRD-YEAR-SECOND-SEMESTER", "BM_THIRD_YEAR_SECOND_SEMESTER",
            # ND
            "ND.THIRD.YEAR.SECOND.SEMESTER", "ND-THIRD-YEAR-SECOND-SEMESTER", "ND_THIRD_YEAR_SECOND_SEMESTER",
            "ND3RD.YEAR2ND.SEMESTER", "ND3RD-YEAR-2ND-SEMESTER", "ND3RD_YEAR_2ND_SEMESTER",
            "ND YEAR3.SEMESTER2", "ND-YEAR-3-SEMESTER-2", "ND_YEAR_3_SEMESTER_2"
        ]
    }
 
    for semester_key, patterns in semester_patterns.items():
        for pattern in patterns:
            flexible_pattern = pattern.replace('.', '[ ._ -]?')
            if re.search(flexible_pattern, filename_upper, re.IGNORECASE):
                logger.info(f"✅ Matched semester '{semester_key}' for filename: {filename}")
                return semester_key
 
    logger.warning(f"❌ Could not determine semester for filename: {filename}")
    return "UNKNOWN_SEMESTER"
# ============================================================================
# FIXED: get_carryover_summary function - UPDATED VERSION
# ============================================================================
def get_carryover_summary(program, set_name):
    """Get summary of carryover students - FIXED VERSION with ZIP support"""
    # CRITICAL FIX: Check if CLEAN_RESULTS exists before trying to scan
    clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
    if not os.path.exists(clean_dir):
        logger.info(f"⚠️ CLEAN_RESULTS not found for {program}/{set_name}, skipping")
        return {
            'total_students': 0,
            'total_courses': 0,
            'by_semester': {},
            'recent_semester': None,
            'recent_count': 0
        }
 
    carryover_records = get_carryover_records(program, set_name)
    summary = {
        'total_students': 0,
        'total_courses': 0,
        'by_semester': {},
        'recent_semester': None,
        'recent_count': 0
    }
 
    logger.info(f"📊 Carryover summary for {program}/{set_name}: {len(carryover_records)} records found")
 
    for record in carryover_records:
        semester = record['semester']
        if semester not in summary['by_semester']:
            summary['by_semester'][semester] = 0
     
        summary['by_semester'][semester] += record['count']
        summary['total_students'] += record['count']
        summary['total_courses'] += sum(len(student['failed_courses']) for student in record['data'])
 
    if summary['by_semester']:
        summary['recent_semester'] = max(summary['by_semester'].keys(),
                                       key=lambda x: summary['by_semester'][x])
        summary['recent_count'] = summary['by_semester'][summary['recent_semester']]
 
    logger.info(f"📊 Final summary: {summary}")
    return summary
def rename_carryover_files(carryover_output_dir, semester_key, resit_timestamp):
    """Rename files in carryover output directory to include CARRYOVER prefix"""
    renamed_files = []
 
    for root, dirs, files in os.walk(carryover_output_dir):
        for file in files:
            if file.lower().endswith(('.xlsx', '.csv', '.pdf')):
                old_path = os.path.join(root, file)
             
                if file.startswith('CARRYOVER_'):
                    renamed_files.append(old_path)
                    continue
             
                file_extension = os.path.splitext(file)[1]
                file_base = os.path.splitext(file)[0]
             
                new_filename = f"CARRYOVER_{semester_key}_{resit_timestamp}_{file_base}{file_extension}"
                new_path = os.path.join(root, new_filename)
             
                try:
                    os.rename(old_path, new_path)
                    renamed_files.append(new_path)
                    logger.info(f"Renamed carryover file: {file} → {new_filename}")
                except Exception as e:
                    logger.error(f"Failed to rename {file}: {e}")
                    renamed_files.append(old_path)
 
    return renamed_files
def debug_resit_processing_details(program_code, set_name, semester_key, resit_file_path, clean_dir):
    """Debug function to provide detailed information about resit processing"""
    debug_info = {
        'program': program_code,
        'set': set_name,
        'semester': semester_key,
        'resit_file': resit_file_path,
        'resit_file_exists': os.path.exists(resit_file_path),
        'clean_dir': clean_dir,
        'clean_dir_exists': os.path.exists(clean_dir),
        'timestamp_folders': [],
        'carryover_records': []
    }
 
    if os.path.exists(clean_dir):
        debug_info['timestamp_folders'] = [f for f in os.listdir(clean_dir)
                                         if f.startswith(f"{set_name}_RESULT-") and os.path.isdir(os.path.join(clean_dir, f))]
 
    latest_folder = sorted(debug_info['timestamp_folders'])[-1] if debug_info['timestamp_folders'] else None
    if latest_folder:
        output_dir = os.path.join(clean_dir, latest_folder)
        carryover_dir = os.path.join(output_dir, "CARRYOVER_RECORDS")
        if os.path.exists(carryover_dir):
            debug_info['carryover_records'] = [f for f in os.listdir(carryover_dir)
                                             if f.startswith("co_student_") and semester_key in f and f.endswith('.json')]
 
    logger.info("RESIT PROCESSING DEBUG INFO:")
    for key, value in debug_info.items():
        logger.info(f" {key}: {value}")
 
    return debug_info
# ============================================================================
# NEW: Function to verify set-specific processing
# ============================================================================
def verify_set_specific_processing(program, selected_set, clean_dir):
    """Verify that processing only affected the selected set"""
    try:
        if selected_set == "all":
            return True # All sets were intended to be processed
         
        # Check if any files were created for other sets
        all_sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
        other_sets = [s for s in all_sets if s != selected_set]
     
        affected_other_sets = []
        for other_set in other_sets:
            other_clean_dir = os.path.join(BASE_DIR, program, other_set, "CLEAN_RESULTS")
            if os.path.exists(other_clean_dir):
                # Check for new result directories in other sets
                new_dirs = [d for d in os.listdir(other_clean_dir)
                           if d.startswith(f"{other_set}_RESULT-") and os.path.isdir(os.path.join(other_clean_dir, d))]
                if new_dirs:
                    affected_other_sets.append(other_set)
     
        if affected_other_sets:
            logger.warning(f"⚠️ Processing for {selected_set} affected other sets: {affected_other_sets}")
            return False
        else:
            logger.info(f"✅ Processing correctly limited to {selected_set}")
            return True
         
    except Exception as e:
        logger.error(f"❌ Error verifying set-specific processing: {e}")
        return True # Don't block processing due to verification error
# ============================================================================
# ENHANCED: Script Processing with STRICT ZIP Enforcement
# ============================================================================
def process_script_with_strict_zip_enforcement(script_name, program, selected_set, env, script_path):
    """Process script with STRICT ZIP enforcement - NO SCATTERED FILES ALLOWED"""
    try:
        # Run the script
        result = subprocess.run(
            [sys.executable, script_path],
            env=env,
            text=True,
            capture_output=True,
            timeout=600,
        )
     
        output_lines = result.stdout.splitlines()
        error_lines = result.stderr.splitlines()
        # Log output
        logger.info("=== SCRIPT STDOUT ===")
        for line in output_lines:
            logger.info(line)
         
        if error_lines:
            logger.info("=== SCRIPT STDERR ===")
            for line in error_lines:
                logger.error(line)
        if result.returncode != 0:
            error_msg = "Script failed. Check logs for details."
            if error_lines:
                error_msg += f" Error: {error_lines[-1]}"
            return {"success": False, "error": error_msg, "output": output_lines}
        # 🔒 STRICT ZIP ENFORCEMENT for ALL scripts
        clean_dir = get_clean_directory(script_name, program, selected_set)
     
        # Enforce ZIP-only policy
        if os.path.exists(clean_dir):
            logger.info(f"🔒 Enforcing STRICT ZIP-only policy for {script_name}")
            enforce_zip_only_policy(clean_dir)
     
        return {"success": True, "output": output_lines}
     
    except subprocess.TimeoutExpired:
        error_msg = f"Script timed out after 10 minutes: {script_name}"
        logger.error(error_msg)
        return {"success": False, "error": error_msg}
    except Exception as e:
        error_msg = f"Script execution error: {str(e)}"
        logger.error(error_msg)
        return {"success": False, "error": error_msg}
# ============================================================================
# NEW: Cleanup function for empty artifacts
# ============================================================================
def clean_up_empty_artifacts(program):
    """Clean up empty ZIP files and directories after processing"""
    try:
        sets = get_available_sets(program)
        for set_name in sets:
            clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
            if not os.path.exists(clean_dir):
                continue
            
            # Remove small/empty ZIP files
            removed_zips = 0
            for f in os.listdir(clean_dir):
                if f.endswith('.zip'):
                    path = os.path.join(clean_dir, f)
                    if os.path.getsize(path) < 200:
                        try:
                            os.remove(path)
                            logger.info(f"🗑️ Removed small/empty ZIP: {f}")
                            removed_zips += 1
                        except Exception as e:
                            logger.error(f"Error removing small ZIP {f}: {e}")
            
            # Remove empty directories
            removed_dirs = 0
            for d in os.listdir(clean_dir):
                path = os.path.join(clean_dir, d)
                if os.path.isdir(path):
                    if not os.listdir(path):
                        try:
                            shutil.rmtree(path)
                            logger.info(f"🗑️ Removed empty directory: {d}")
                            removed_dirs += 1
                        except Exception as e:
                            logger.error(f"Error removing empty dir {d}: {e}")
            
            if removed_zips or removed_dirs:
                logger.info(f"🧹 Cleaned up {program}/{set_name}: {removed_zips} ZIPs, {removed_dirs} dirs removed")
        
    except Exception as e:
        logger.error(f"Error in clean_up_empty_artifacts for {program}: {e}")
# ============================================================================
# NEW: Debug route for ZIP status
# ============================================================================
@app.route("/debug_zip_status")
@login_required
def debug_zip_status():
    """Debug route to check ZIP file status across all directories"""
    debug_info = {
        "base_dir": BASE_DIR,
        "scan_summary": {},
        "problems_found": []
    }
 
    # Scan all program directories
    for program in ["ND", "BN", "BM"]:
        program_dir = os.path.join(BASE_DIR, program)
        debug_info["scan_summary"][program] = {}
     
        if not os.path.exists(program_dir):
            debug_info["scan_summary"][program]["error"] = "Directory not found"
            continue
         
        sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
     
        for set_name in sets:
            clean_dir = os.path.join(program_dir, set_name, "CLEAN_RESULTS")
            debug_info["scan_summary"][program][set_name] = {
                "clean_dir": clean_dir,
                "exists": os.path.exists(clean_dir),
                "zip_files": [],
                "scattered_dirs": [],
                "scattered_files": []
            }
         
            if os.path.exists(clean_dir):
                # Check for ZIP files
                zip_files = [f for f in os.listdir(clean_dir) if f.endswith('.zip')]
                debug_info["scan_summary"][program][set_name]["zip_files"] = zip_files
             
                # Check for scattered directories
                scattered_dirs = [d for d in os.listdir(clean_dir)
                                if os.path.isdir(os.path.join(clean_dir, d)) and
                                ("RESULT" in d or "RESIT" in d)]
                debug_info["scan_summary"][program][set_name]["scattered_dirs"] = scattered_dirs
             
                # Check for scattered files
                scattered_files = [f for f in os.listdir(clean_dir)
                                 if os.path.isfile(os.path.join(clean_dir, f)) and
                                 not f.endswith('.zip') and not f.startswith('~')]
                debug_info["scan_summary"][program][set_name]["scattered_files"] = scattered_files
             
                # Record problems
                if scattered_dirs or scattered_files:
                    problem_msg = f"{program}/{set_name}: {len(scattered_dirs)} dirs, {len(scattered_files)} files need cleanup"
                    debug_info["problems_found"].append(problem_msg)
 
    return jsonify(debug_info)
@app.route("/fix_scattered_files")
@login_required
def fix_scattered_files():
    """Route to manually fix scattered files by creating ZIPs"""
    try:
        fixed_count = 0
     
        for program in ["ND", "BN", "BM"]:
            program_dir = os.path.join(BASE_DIR, program)
            if not os.path.exists(program_dir):
                continue
             
            sets = ND_SETS if program == "ND" else (BN_SETS if program == "BN" else BM_SETS)
         
            for set_name in sets:
                clean_dir = os.path.join(program_dir, set_name, "CLEAN_RESULTS")
                if os.path.exists(clean_dir):
                    if ensure_zipped_results(clean_dir, f"exam_processor_{program.lower()}", set_name):
                        fixed_count += 1
     
        flash(f"Fixed scattered files in {fixed_count} directories", "success")
        return redirect(url_for("download_center"))
     
    except Exception as e:
        logger.error(f"Error fixing scattered files: {e}")
        flash(f"Error fixing scattered files: {str(e)}", "error")
        return redirect(url_for("download_center"))
# ============================================================================
# UPDATED: Dashboard Route
# ============================================================================
@app.route("/dashboard")
@login_required
def dashboard():
    """Main dashboard with program-specific links"""
    try:
        # Get carryover summaries for all programs
        carryover_summaries = {}
     
        # ND Carryover Summary
        nd_carryover_data = {}
        for set_name in ND_SETS:
            records = get_carryover_records("ND", set_name)
            if records:
                nd_carryover_data[set_name] = {
                    'total_students': sum(record['count'] for record in records),
                    'total_courses': len(records)
                }
        if nd_carryover_data:
            carryover_summaries['ND'] = nd_carryover_data
     
        # BN Carryover Summary
        bn_carryover_data = {}
        for set_name in BN_SETS:
            records = get_carryover_records("BN", set_name)
            if records:
                bn_carryover_data[set_name] = {
                    'total_students': sum(record['count'] for record in records),
                    'total_courses': len(records)
                }
        if bn_carryover_data:
            carryover_summaries['BN'] = bn_carryover_data
     
        # BM Carryover Summary
        bm_carryover_data = {}
        for set_name in BM_SETS:
            records = get_carryover_records("BM", set_name)
            if records:
                bm_carryover_data[set_name] = {
                    'total_students': sum(record['count'] for record in records),
                    'total_courses': len(records)
                }
        if bm_carryover_data:
            carryover_summaries['BM'] = bm_carryover_data
     
        return render_template(
            "dashboard.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            carryover_summaries=carryover_summaries
        )
    except Exception as e:
        logger.error(f"Dashboard error: {e}")
        flash(f"Error loading dashboard: {str(e)}", "error")
        return render_template("dashboard.html", carryover_summaries={})
# Routes
@app.route("/", methods=["GET"])
def index():
    return redirect(url_for("login"))
@app.route("/login", methods=["GET", "POST"])
def login():
    try:
        app.logger.info(f"Login route accessed - Method: {request.method}")
        logger.info(f"Attempting to load template: {os.path.join(TEMPLATE_DIR, 'login.html')}")
        if request.method == "POST":
            password = request.form.get("password")
            app.logger.info(f"Login attempt - password provided: {bool(password)}")
            if password == PASSWORD:
                session["logged_in"] = True
                flash("Successfully logged in!", "success")
                app.logger.info("Login successful")
                return redirect(url_for("dashboard"))
            else:
                flash("Invalid password. Please try again.", "error")
                app.logger.warning("Invalid password attempt")
        return render_template(
            "login.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None
        )
    except TemplateNotFound as e:
        app.logger.error(f"Template not found: {e} at {TEMPLATE_DIR}")
        logger.error(f"Template error: {e} at {TEMPLATE_DIR}")
        return f"<h1>Template Error</h1><p>Template not found: login.html at {TEMPLATE_DIR}</p><p>Available templates: {os.listdir(TEMPLATE_DIR) if os.path.exists(TEMPLATE_DIR) else 'None'}</p>", 500
    except Exception as e:
        app.logger.error(f"Login error: {e}")
        logger.error(f"Login error: {e}")
        return f"<h1>Server Error</h1><p>{str(e)}</p>", 500
@app.route("/debug_paths")
@login_required
def debug_paths():
    try:
        paths = {
            "Project Root": PROJECT_ROOT,
            "Script Directory": SCRIPT_DIR,
            "Base Directory": BASE_DIR,
            "Template Directory": TEMPLATE_DIR,
            "Static Directory": STATIC_DIR,
        }
        path_status = {key: os.path.exists(path) for key, path in paths.items()}
        templates = os.listdir(TEMPLATE_DIR) if os.path.exists(TEMPLATE_DIR) else []
        script_paths = {name: os.path.join(SCRIPT_DIR, path) for name, path in SCRIPT_MAP.items()}
        script_status = {name: os.path.exists(path) for name, path in script_paths.items()}
        env_vars = {
            "FLASK_SECRET": os.getenv("FLASK_SECRET", "Not set"),
            "STUDENT_CLEANER_PASSWORD": os.getenv("STUDENT_CLEANER_PASSWORD", "Not set"),
            "COLLEGE_NAME": COLLEGE,
            "DEPARTMENT": DEPARTMENT,
            "BASE_DIR": BASE_DIR,
        }
        return render_template(
            "debug_paths.html",
            environment="Railway Production" if not is_local_environment() else "Local Development",
            paths=paths,
            path_status=path_status,
            templates=templates,
            script_paths=script_paths,
            script_status=script_status,
            env_vars=env_vars,
            college=COLLEGE,
            department=DEPARTMENT,
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            nd_sets=ND_SETS,
            bn_sets=BN_SETS,
            bm_sets=BM_SETS,
            programs=PROGRAMS
        )
    except TemplateNotFound as e:
        logger.error(f"Debug paths template not found: {e}")
        flash("Debug paths template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"Debug paths error: {e}")
        flash(f"Error loading debug page: {str(e)}", "error")
        return redirect(url_for("dashboard"))
@app.route("/debug_internal_files")
@login_required
def debug_internal_files():
    """Debug route to check internal exam files detection"""
    input_dir = os.path.join(BASE_DIR, "OBJ_RESULT", "RAW_OBJ")
    files_info = {
        'input_dir': input_dir,
        'exists': os.path.exists(input_dir),
        'files_found': [],
        'check_internal_exam_files_result': check_internal_exam_files(input_dir)
    }
 
    if os.path.exists(input_dir):
        # Method 1: glob method (same as processing script)
        csv_files = glob.glob(os.path.join(input_dir, "*.csv"))
        xls_files = glob.glob(os.path.join(input_dir, "*.xls")) + glob.glob(os.path.join(input_dir, "*.xlsx"))
        all_files = [f for f in (csv_files + xls_files) if not os.path.basename(f).startswith("~$")]
     
        # Method 2: Pattern-based method
        pattern_files = [
            f for f in os.listdir(input_dir)
            if f.lower().endswith((".xlsx", ".xls", ".csv"))
            and not f.startswith("~")
            and (
                f.startswith("Set") or
                "ND" in f.upper() and "SET" in f.upper() or
                "OBJ" in f.upper() or
                "RESULT" in f.upper()
            )
        ]
     
        files_info['glob_method_files'] = [os.path.basename(f) for f in all_files]
        files_info['pattern_method_files'] = pattern_files
        files_info['all_files_in_dir'] = os.listdir(input_dir)
     
        # Final combined list
        final_files = list(set(all_files + [os.path.join(input_dir, f) for f in pattern_files]))
        files_info['final_files'] = [os.path.basename(f) for f in final_files]
        files_info['files_found'] = files_info['final_files']
 
    return jsonify(files_info)
@app.route("/debug_dir_contents")
@login_required
def debug_dir_contents():
    try:
        dir_contents = {}
        for dir_path in [BASE_DIR]:
            contents = {}
            if os.path.exists(dir_path):
                for root, dirs, files in os.walk(dir_path):
                    relative_path = os.path.relpath(root, dir_path)
                    contents[relative_path or "."] = {
                        "dirs": dirs,
                        "files": [f for f in files if f.lower().endswith((".xlsx", ".csv", ".pdf"))]
                    }
            dir_contents[dir_path] = contents
        return jsonify(dir_contents)
    except Exception as e:
        app.logger.error(f"Debug dir contents error: {e}")
        return jsonify({"error": str(e)}), 500
@app.route("/upload_center", methods=["GET", "POST"])
@login_required
def upload_center():
    try:
        if request.method == "GET":
            return render_template(
                "upload_center.html",
                college=COLLEGE,
                department=DEPARTMENT,
                environment="Railway Production" if not is_local_environment() else "Local Development",
                logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
                nd_sets=ND_SETS,
                bn_sets=BN_SETS,
                bm_sets=BM_SETS,
                programs=PROGRAMS
            )
        return redirect(url_for("dashboard"))
    except TemplateNotFound as e:
        logger.error(f"Upload center template not found: {e}")
        flash("Upload center template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"Upload center error: {e}")
        flash(f"Error loading upload center: {str(e)}", "error")
        return redirect(url_for("dashboard"))
@app.route("/handle_upload", methods=["POST"])
@login_required
def handle_upload():
    """Handle file uploads from the upload center"""
    try:
        program = request.form.get("program")
        files = request.files.getlist("files")
     
        if not program:
            flash("Please select a program.", "error")
            return redirect(url_for("upload_center"))
         
        if not files or all(file.filename == '' for file in files):
            flash("Please select at least one file to upload.", "error")
            return redirect(url_for("upload_center"))
     
        program_map = {
            "nd": ("exam_processor_nd", "ND"),
            "bn": ("exam_processor_bn", "BN"),
            "bm": ("exam_processor_bm", "BM")
        }
     
        if program not in program_map:
            flash("Invalid program selected.", "error")
            return redirect(url_for("upload_center"))
         
        script_name, program_code = program_map[program]
     
        set_name = None
        if program == "nd":
            set_name = request.form.get("nd_set")
        elif program == "bn":
            set_name = request.form.get("bn_set")
        elif program == "bm":
            set_name = request.form.get("bm_set")
         
        if not set_name:
            flash(f"Please select a {program.upper()} set.", "error")
            return redirect(url_for("upload_center"))
     
        raw_dir = get_raw_directory(script_name, program_code, set_name)
        logger.info(f"Uploading to directory: {raw_dir}")
     
        os.makedirs(raw_dir, exist_ok=True)
     
        saved_files = []
        skipped_files = []
     
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(raw_dir, filename)
             
                file.save(file_path)
                saved_files.append(filename)
                logger.info(f"Saved file: {file_path}")
             
                if filename.lower().endswith(".zip"):
                    try:
                        with zipfile.ZipFile(file_path, "r") as zip_ref:
                            zip_ref.extractall(raw_dir)
                        logger.info(f"Extracted ZIP file: {filename}")
                        saved_files.append(f"{filename} (extracted)")
                    except zipfile.BadZipFile:
                        logger.error(f"Invalid ZIP file: {filename}")
                        skipped_files.append(f"{filename} (invalid ZIP)")
                    except Exception as e:
                        logger.error(f"Error extracting ZIP {filename}: {e}")
                        skipped_files.append(f"{filename} (extraction error)")
            else:
                skipped_files.append(file.filename if file.filename else "unknown file")
     
        if saved_files:
            success_msg = f"Successfully uploaded {len(saved_files)} file(s) to {program_code}/{set_name}/RAW_RESULTS"
            if skipped_files:
                success_msg += f" (Skipped {len(skipped_files)} invalid files)"
            flash(success_msg, "success")
            logger.info(f"Upload completed: {success_msg}")
        else:
            flash("No valid files were uploaded. Please check file formats.", "error")
     
        if skipped_files:
            flash(f"Skipped files: {', '.join(skipped_files)}", "warning")
         
        return redirect(url_for("upload_center"))
     
    except Exception as e:
        app.logger.error(f"Upload error: {e}")
        flash(f"Upload failed: {str(e)}", "error")
        return redirect(url_for("upload_center"))
# ============================================================================
# UPDATED: handle_resit_upload function - CHANGED: Simplified for ND only
# ============================================================================
@app.route("/handle_resit_upload", methods=["POST"])
@login_required
def handle_resit_upload():
    """Handle ND resit file uploads - CHANGED: Simplified for ND only"""
    try:
        logger.info("ND CARRYOVER UPLOAD: Route called")
     
        # Get form data - simplified for ND only
        set_name = request.form.get("nd_set")
        selected_semesters = request.form.getlist("selected_semesters")
        resit_file = request.files.get("resit_file")
     
        logger.info(f"Received - Set: {set_name}, Semesters: {selected_semesters}, File: {resit_file.filename if resit_file else 'None'}")
     
        # Validation
        if not resit_file or resit_file.filename == '':
            flash("Please select a file", "error")
            return redirect(url_for("upload_center"))
     
        if not set_name or set_name not in ND_SETS:
            flash(f"Please select a valid ND set from {ND_SETS}", "error")
            return redirect(url_for("upload_center"))
     
        if not selected_semesters:
            flash("Please select at least one semester", "error")
            return redirect(url_for("upload_center"))
     
        # Save file to ND carryover directory
        raw_dir = os.path.join(BASE_DIR, "ND", set_name, "RAW_RESULTS", "CARRYOVER")
        os.makedirs(raw_dir, exist_ok=True)
     
        from werkzeug.utils import secure_filename
        filename = secure_filename(resit_file.filename)
        file_path = os.path.join(raw_dir, filename)
        resit_file.save(file_path)
     
        logger.info(f"File saved: {file_path}")
     
        # Verify
        if os.path.exists(file_path):
            file_size = os.path.getsize(file_path)
            logger.info(f"✅ ND resit file saved: {file_path} ({file_size} bytes)")
         
            semester_display = ", ".join(selected_semesters)
            flash(f"Successfully uploaded ND carryover file to {set_name}/CARRYOVER for semesters: {semester_display}", "success")
        else:
            logger.error(f"❌ File was not saved: {file_path}")
            flash("Failed to save resit file", "error")
     
        return redirect(url_for("upload_center"))
     
    except Exception as e:
        logger.error(f"ERROR in ND resit upload: {str(e)}")
        flash(f"Upload failed: {str(e)}", "error")
        return redirect(url_for("upload_center"))
@app.route("/download_center")
@login_required
def download_center():
    try:
        files_by_category = get_files_by_category()
        for category, files in files_by_category.items():
            if isinstance(files, dict):
                total_files = sum(len(f) for f in files.values())
                app.logger.info(f"Download center - {category}: {total_files} ZIP files across {len(files)} sets")
            else:
                app.logger.info(f"Download center - {category}: {len(files)} ZIP files")
        return render_template(
            "download_center.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            files_by_category=files_by_category,
            nd_sets=ND_SETS,
            bn_sets=BN_SETS,
            bm_sets=BM_SETS,
            programs=PROGRAMS
        )
    except TemplateNotFound as e:
        logger.error(f"Download center template not found: {e}")
        flash("Download center template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"Download center error: {e}")
        flash(f"Error loading download center: {str(e)}", "error")
        return redirect(url_for("dashboard"))
@app.route("/file_browser")
@login_required
def file_browser():
    try:
        sets = get_sets_and_folders()
     
        nd_sets = []
        bn_sets = []
        bm_sets = []
     
        for key in sets.keys():
            if key.startswith('ND_'):
                nd_sets.append(key.replace('ND_', ''))
            elif key.startswith('BN_'):
                bn_sets.append(key.replace('BN_', ''))
            elif key.startswith('BM_'):
                bm_sets.append(key.replace('BM_', ''))
     
        app.logger.info(f"File browser - ND sets: {nd_sets}, BN sets: {bn_sets}, BM sets: {bm_sets}")
        app.logger.info(f"File browser - Total sets: {len(sets)}")
     
        for set_key, folders in sets.items():
            total_files = sum(len(folder.files) for folder in folders)
            app.logger.info(f"Set '{set_key}': {total_files} ZIP files across {len(folders)} folders")
     
        return render_template(
            "file_browser.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            sets=sets,
            nd_sets=nd_sets,
            bn_sets=bn_sets,
            bm_sets=bm_sets,
            programs=PROGRAMS,
            BASE_DIR=BASE_DIR,
            processed_dir=BASE_DIR
        )
    except TemplateNotFound as e:
        logger.error(f"File browser template not found: {e}")
        flash("File browser template not found.", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        app.logger.error(f"File browser error: {e}")
        flash(f"Error loading file browser: {str(e)}", "error")
        return redirect(url_for("dashboard"))
# ============================================================================
# FIXED: Carryover route - UPDATED VERSION
# ============================================================================
@app.route("/carryover")
@login_required
def carryover():
    """Redirect to ND carryover for backward compatibility"""
    flash("Redirected to ND Carryover Management", "info")
    return redirect(url_for("nd_carryover"))
# ============================================================================
# FIXED: BN Carryover Management Route - ADDED bn_sets
# ============================================================================
@app.route("/bn_carryover")
@login_required
def bn_carryover():
    """Basic Nursing carryover management dashboard"""
    try:
        bn_carryover_data = {}
     
        for set_name in BN_SETS:
            clean_dir = os.path.join(BASE_DIR, "BN", set_name, "CLEAN_RESULTS")
            if not os.path.exists(clean_dir):
                continue
         
            records = get_carryover_records("BN", set_name)
            if records:
                bn_carryover_data[set_name] = {
                    'records': records,
                    'total_students': sum(record['count'] for record in records),
                    'total_semesters': len(records)
                }
     
        return render_template(
            "bn_carryover_management.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            bn_carryover_data=bn_carryover_data,
            bn_sets=BN_SETS # ✅ FIX: Added this line
        )
    except Exception as e:
        logger.error(f"BN carryover management error: {e}")
        flash(f"Error loading BN carryover management: {str(e)}", "error")
        return redirect(url_for("dashboard"))
# ============================================================================
# FIXED: BM Carryover Management Route - ADDED bm_sets
# ============================================================================
@app.route("/bm_carryover")
@login_required
def bm_carryover():
    """Basic Midwifery carryover management dashboard"""
    try:
        bm_carryover_data = {}
     
        for set_name in BM_SETS:
            clean_dir = os.path.join(BASE_DIR, "BM", set_name, "CLEAN_RESULTS")
            if not os.path.exists(clean_dir):
                continue
         
            records = get_carryover_records("BM", set_name)
            if records:
                bm_carryover_data[set_name] = {
                    'records': records,
                    'total_students': sum(record['count'] for record in records),
                    'total_semesters': len(records)
                }
     
        return render_template(
            "bm_carryover_management.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            bm_carryover_data=bm_carryover_data,
            bm_sets=BM_SETS # ✅ FIX: Added this line
        )
    except Exception as e:
        logger.error(f"BM carryover management error: {e}")
        flash(f"Error loading BM carryover management: {str(e)}", "error")
        return redirect(url_for("dashboard"))
# ============================================================================
# NEW ROUTE: ND Carryover Management
# ============================================================================
@app.route("/nd_carryover")
@login_required
def nd_carryover():
    """National Diploma carryover management dashboard"""
    try:
        nd_carryover_data = {}
     
        for set_name in ND_SETS:
            clean_dir = os.path.join(BASE_DIR, "ND", set_name, "CLEAN_RESULTS")
            if not os.path.exists(clean_dir):
                continue
         
            records = get_carryover_records("ND", set_name)
            if records:
                nd_carryover_data[set_name] = {
                    'records': records,
                    'total_students': sum(record['count'] for record in records),
                    'total_semesters': len(records)
                }
     
        return render_template(
            "nd_carryover_management.html",
            college=COLLEGE,
            department=DEPARTMENT,
            environment="Railway Production" if not is_local_environment() else "Local Development",
            logo_url=url_for("static", filename="logo.png") if os.path.exists(os.path.join(STATIC_DIR, "logo.png")) else None,
            nd_carryover_data=nd_carryover_data,
            nd_sets=ND_SETS
        )
    except Exception as e:
        logger.error(f"ND carryover management error: {e}")
        flash(f"Error loading ND carryover management: {str(e)}", "error")
        return redirect(url_for("dashboard"))
@app.route("/debug_carryover_files/<program>/<set_name>")
@login_required
def debug_carryover_files_detail(program, set_name):
    """Debug route to check carryover file semester matching - ENHANCED VERSION."""
    try:
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            return jsonify({'error': f"Clean directory not found: {clean_dir}"})
     
        debug_info = {
            'clean_dir': clean_dir,
            'items': [],
            'regular_results': [],
            'carryover_results': [],
            'carryover_records': [],
            'available_files': os.listdir(clean_dir)
        }
     
        # List all items with detailed classification
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
            item_info = {
                'name': item,
                'type': 'dir' if os.path.isdir(item_path) else 'file',
                'is_regular_result': (item.startswith(f"{set_name}_RESULT-") and
                                    "RESULT" in item.upper() and
                                    not "CARRYOVER" in item.upper()),
                'is_carryover_result': 'CARRYOVER' in item.upper(),
                'is_zip': item.endswith('.zip')
            }
            debug_info['items'].append(item_info)
         
            if item_info['is_regular_result']:
                debug_info['regular_results'].append(item)
            elif item_info['is_carryover_result']:
                debug_info['carryover_results'].append(item)
     
        # Check carryover records in latest REGULAR result (not carryover result)
        if debug_info['regular_results']:
            latest_regular = sorted(debug_info['regular_results'])[-1]
            debug_info['latest_regular'] = latest_regular
            latest_path = os.path.join(clean_dir, latest_regular)
         
            if latest_regular.endswith('.zip'):
                # Extract from ZIP
                try:
                    with zipfile.ZipFile(latest_path, 'r') as zip_ref:
                        # Look for carryover JSON files
                        json_files = [f for f in zip_ref.namelist()
                                    if f.startswith("CARRYOVER_RECORDS/") and f.endswith('.json')]
                     
                        for json_file in json_files:
                            extracted = extract_semester_from_filename(json_file)
                            standardized = standardize_semester_key(extracted, program)
                            debug_info['carryover_records'].append({
                                'filename': json_file,
                                'extracted_semester': extracted,
                                'standardized_semester': standardized,
                                'source': f'ZIP: {latest_regular}'
                            })
                except Exception as e:
                    debug_info['zip_error'] = str(e)
            else:
                # Check directory
                carryover_dir = os.path.join(latest_path, "CARRYOVER_RECORDS")
                if os.path.exists(carryover_dir):
                    for file in os.listdir(carryover_dir):
                        if file.endswith('.json'):
                            extracted = extract_semester_from_filename(file)
                            standardized = standardize_semester_key(extracted, program)
                            debug_info['carryover_records'].append({
                                'filename': file,
                                'extracted_semester': extracted,
                                'standardized_semester': standardized,
                                'source': f'DIR: {latest_regular}'
                            })
     
        return jsonify(debug_info)
     
    except Exception as e:
        return jsonify({'error': str(e)})
# ============================================================================
# FIX D: UPDATED: process_resit function - COMPLETELY REWRITTEN FOR ND CARRYOVER
# ============================================================================
@app.route("/process_resit", methods=["POST"])
@login_required
def process_resit():
    """Process resit results - FIXED with proper env vars."""
    try:
        logger.info("🎯 RESIT PROCESSING: Starting")
     
        # Get form data
        set_name = request.form.get("resit_set", "").strip()
        semester_key = request.form.get("resit_semester", "").strip()
        resit_file = request.files.get("resit_file")
     
        logger.info(f"📥 Received - Set: {set_name}, Semester: {semester_key}")
     
        # Validation
        if not all([set_name, semester_key, resit_file]):
            missing = []
            if not set_name: missing.append("set")
            if not semester_key: missing.append("semester")
            if not resit_file: missing.append("file")
            flash(f"Missing fields: {', '.join(missing)}", "error")
            return redirect(url_for("nd_carryover"))
     
        # Determine program from set name
        if set_name in ND_SETS:
            program = "ND"
            processor_script = "nd_carryover_processor.py"
        elif set_name in BN_SETS:
            program = "BN"
            processor_script = "exam_processor_bn.py"
        elif set_name in BM_SETS:
            program = "BM"
            processor_script = "exam_processor_bm.py"
        else:
            flash(f"Invalid set: {set_name}", "error")
            return redirect(url_for("nd_carryover"))
     
        # Save resit file
        resit_dir = os.path.join(BASE_DIR, program, set_name, "RAW_RESULTS", "CARRYOVER")
        os.makedirs(resit_dir, exist_ok=True)
     
        filename = secure_filename(resit_file.filename)
        resit_file_path = os.path.join(resit_dir, filename)
        resit_file.save(resit_file_path)
     
        if not os.path.exists(resit_file_path):
            flash("Failed to save resit file", "error")
            return redirect(url_for("nd_carryover"))
     
        logger.info(f"✅ Resit file saved: {resit_file_path}")
     
        # ============================================================================
        # STEP 1: Find latest regular ZIP for BASE_RESULT_PATH
        # ============================================================================
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        regular_zips = [f for f in os.listdir(clean_dir)
                       if f.startswith(f"{set_name}_RESULT-")
                       and f.endswith('.zip')
                       and 'CARRYOVER' not in f.upper()
                       and 'UPDATED' not in f.upper()]
        if regular_zips:
            latest_zip = sorted(regular_zips)[-1]
            base_result_path = os.path.join(clean_dir, latest_zip)
            os.environ['BASE_RESULT_PATH'] = base_result_path
            logger.info(f"Set BASE_RESULT_PATH: {base_result_path}")
        else:
            flash("No regular result ZIP found for update. Please process regular results first.", "error")
            return redirect(url_for("nd_carryover"))
        # ============================================================================
        # STEP 2: Set OUTPUT_DIR
        # ============================================================================
        os.environ['OUTPUT_DIR'] = clean_dir
        logger.info(f"Set OUTPUT_DIR: {clean_dir}")
        # Set other environment variables
        os.environ['SELECTED_SET'] = set_name
        os.environ['SELECTED_SEMESTERS'] = semester_key
        os.environ['RESIT_FILE_PATH'] = resit_file_path
        os.environ['WEB_MODE'] = 'true'
     
        # Get script path
        script_path = os.path.join(SCRIPT_DIR, processor_script)
     
        logger.info(f"🚀 Running processor: {script_path}")
        for key in ["BASE_DIR", "SELECTED_SET", "SELECTED_SEMESTERS", "RESIT_FILE_PATH"]:
            logger.info(f" {key}: {os.environ.get(key)}")
     
        # Run processor
        result = subprocess.run(
            [sys.executable, script_path],
            env=os.environ.copy(),
            text=True,
            capture_output=True,
            timeout=600,
        )
     
        # Parse results
        output_lines = result.stdout.splitlines()
        error_lines = result.stderr.splitlines()
     
        # Log output
        logger.info("="*60)
        logger.info("PROCESSOR OUTPUT:")
        logger.info("="*60)
        for line in output_lines:
            logger.info(line)
     
        if error_lines:
            logger.info("="*60)
            logger.info("PROCESSOR ERRORS:")
            logger.info("="*60)
            for line in error_lines:
                logger.error(line)
     
        # Check success
        success_indicators = [
            "CARRYOVER PROCESSING COMPLETED" in " ".join(output_lines),
            "Updated" in " ".join(output_lines) and "scores" in " ".join(output_lines),
            result.returncode == 0
        ]
     
        if any(success_indicators):
            # Extract update message
            update_msg = None
            for line in output_lines:
                if "Updated" in line and "scores for" in line and "students" in line:
                    update_msg = line.strip()
                    break
         
            msg = f"✅ Resit processing completed! {update_msg}" if update_msg else "✅ Resit processing completed"
            flash(msg, "success")
            logger.info(f"✅ SUCCESS: {msg}")
        else:
            error_msg = "❌ Resit processing failed"
         
            # Find specific error
            if error_lines:
                for line in reversed(error_lines):
                    if line.strip() and not line.startswith(("File", "Traceback", " ")):
                        error_msg += f": {line.strip()}"
                        break
         
            flash(error_msg, "error")
            logger.error(error_msg)
     
        # Redirect based on program
        if program == "ND":
            return redirect(url_for("nd_carryover"))
        elif program == "BN":
            return redirect(url_for("bn_carryover"))
        else:
            return redirect(url_for("bm_carryover"))
     
    except subprocess.TimeoutExpired:
        logger.error("RESIT PROCESSING TIMEOUT")
        flash("Resit processing timed out", "error")
        return redirect(url_for("nd_carryover"))
    except Exception as e:
        logger.error(f"RESIT PROCESSING ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f"Resit processing failed: {str(e)}", "error")
        return redirect(url_for("nd_carryover"))
# ============================================================================
# UPDATED: Process ND Carryover Route with Enhanced Logging
# ============================================================================
@app.route('/process_nd_carryover', methods=['POST'])
@login_required
def process_nd_carryover():
    """Process ND carryover resit results with automatic mastersheet update"""
    try:
        logger.info("ND CARRYOVER: Route called")
      
        # Get form data
        selected_set = request.form.get('selected_set')
        selected_semester = request.form.get('selected_semester')
        resit_file = request.files.get('resit_file')
      
        logger.info(f"ND CARRYOVER: Set={selected_set}, Semester={selected_semester}")
      
        # Validation
        if not all([selected_set, selected_semester, resit_file]):
            flash('All fields are required', 'error')
            return redirect(url_for('nd_carryover'))
      
        # Save uploaded file
        upload_dir = os.path.join(BASE_DIR, "ND", selected_set, "RAW_RESULTS", "CARRYOVER")
        os.makedirs(upload_dir, exist_ok=True)
      
        filename = f"nd_resit_{selected_semester}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = os.path.join(upload_dir, filename)
        resit_file.save(file_path)
      
        logger.info(f"✅ Resit file saved: {file_path}")
      
        # ============================================================================
        # STEP 1: Find latest regular ZIP for BASE_RESULT_PATH
        # ============================================================================
        clean_dir = os.path.join(BASE_DIR, "ND", selected_set, "CLEAN_RESULTS")
        regular_zips = [f for f in os.listdir(clean_dir)
                       if f.startswith(f"{selected_set}_RESULT-")
                       and f.endswith('.zip')
                       and 'CARRYOVER' not in f.upper()
                       and 'UPDATED' not in f.upper()]
        if regular_zips:
            latest_zip = sorted(regular_zips)[-1]
            base_result_path = os.path.join(clean_dir, latest_zip)
            os.environ['BASE_RESULT_PATH'] = base_result_path
            logger.info(f"Set BASE_RESULT_PATH: {base_result_path}")
        else:
            flash("No regular result ZIP found for update. Please process regular results first.", "error")
            return redirect(url_for("nd_carryover"))
        # ============================================================================
        # STEP 2: Set OUTPUT_DIR
        # ============================================================================
        os.environ['OUTPUT_DIR'] = clean_dir
        logger.info(f"Set OUTPUT_DIR: {clean_dir}")
        # Set other environment variables
        os.environ['SELECTED_SET'] = selected_set
        os.environ['SELECTED_SEMESTERS'] = selected_semester
        os.environ['RESIT_FILE_PATH'] = file_path
        os.environ['WEB_MODE'] = 'true'
      
        # Run processor
        script_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(script_dir)
        processor_path = os.path.join(SCRIPT_DIR, 'nd_carryover_processor.py')
      
        if not os.path.exists(processor_path):
            flash(f'Processor not found: {processor_path}', 'error')
            return redirect(url_for('nd_carryover'))
      
        result = subprocess.run(
            [sys.executable, processor_path],
            capture_output=True,
            text=True,
            timeout=600
        )
      
        # Clean up uploaded file
        try:
            os.remove(file_path)
        except:
            pass
      
        # ENHANCED: Log subprocess output
        logger.info("=== ND CARRYOVER PROCESSOR OUTPUT ===")
        for line in result.stdout.splitlines():
            logger.info(line)
        logger.info("=== ND CARRYOVER PROCESSOR ERRORS ===")
        for line in result.stderr.splitlines():
            logger.error(line)
      
        if result.returncode == 0:
            # Check if UPDATED_ ZIP was actually created
            updated_zips = [f for f in os.listdir(clean_dir)
                          if f.startswith("UPDATED_") and f.endswith('.zip')]
          
            if updated_zips:
                flash('✅ Carryover processing completed! UPDATED ZIP was created.', 'success')
                logger.info(f"✅ UPDATED ZIP created: {updated_zips[0]}")
            else:
                flash('⚠️ Carryover processing completed but UPDATED ZIP was not created. Check logs.', 'warning')
                logger.warning("UPDATED ZIP was not created despite successful processing")
        else:
            error_msg = result.stderr.splitlines()[-1] if result.stderr else "Unknown error"
            flash(f'❌ Carryover processing failed: {error_msg}', 'error')
            logger.error(f"Subprocess failed with code {result.returncode}")
      
        return redirect(url_for('nd_carryover'))
      
    except Exception as e:
        logger.error(f"ND carryover error: {e}")
        flash(f'Error: {str(e)}', 'error')
        return redirect(url_for('nd_carryover'))
@app.route("/debug_resit_files/<program>/<set_name>")
@login_required
def debug_resit_files(program, set_name):
    """Debug route to check generated resit files"""
    try:
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            return jsonify({'error': f"Clean directory not found: {clean_dir}"})
     
        resit_folders = [f for f in os.listdir(clean_dir)
                        if f.startswith("RESIT_") and os.path.isdir(os.path.join(clean_dir, f))]
     
        debug_info = {
            'clean_dir': clean_dir,
            'resit_folders': [],
            'resit_zips': []
        }
     
        for folder in resit_folders:
            folder_path = os.path.join(clean_dir, folder)
            files = []
            for root, dirs, filenames in os.walk(folder_path):
                for file in filenames:
                    if file.lower().endswith(('.xlsx', '.csv', '.pdf')):
                        files.append({
                            'name': file,
                            'path': os.path.relpath(os.path.join(root, file), clean_dir),
                            'size': os.path.getsize(os.path.join(root, file))
                        })
            debug_info['resit_folders'].append({
                'name': folder,
                'files': files
            })
     
        zip_files = [f for f in os.listdir(clean_dir)
                    if f.startswith("RESIT_") and f.endswith('.zip')]
     
        for zip_file in zip_files:
            zip_path = os.path.join(clean_dir, zip_file)
            debug_info['resit_zips'].append({
                'name': zip_file,
                'size': os.path.getsize(zip_path),
                'path': zip_path
            })
     
        return jsonify(debug_info)
     
    except Exception as e:
        return jsonify({'error': str(e)})
# ============================================================================
# FIXED: delete function - CORRECTED VERSION
# ============================================================================
@app.route("/delete/<path:filename>", methods=["POST"])
@login_required
def delete(filename):
    try:
        critical_dirs = [SCRIPT_DIR, TEMPLATE_DIR, STATIC_DIR, PROJECT_ROOT]
     
        file_path = None
        for root, _, files in os.walk(BASE_DIR):
            if os.path.basename(filename) in files:
                candidate = os.path.join(root, os.path.basename(filename))
                if os.path.exists(candidate):
                    file_path = candidate
                    break
     
        if not file_path:
            for root, dirs, _ in os.walk(BASE_DIR):
                if os.path.basename(filename) in dirs:
                    candidate = os.path.join(root, os.path.basename(filename))
                    if os.path.exists(candidate):
                        file_path = candidate
                        break
     
        if not file_path:
            flash(f"Path '{filename}' not found.", "error")
            logger.warning(f"Deletion failed: Path not found - {filename}")
            return redirect(request.referrer or url_for("file_browser"))
     
        abs_file_path = os.path.abspath(file_path)
        for critical_dir in critical_dirs:
            abs_critical_dir = os.path.abspath(critical_dir)
            if (abs_file_path == abs_critical_dir or
                os.path.dirname(abs_file_path) == abs_critical_dir):
                flash(f"Cannot delete critical system path: {filename}", "error")
                logger.warning(f"Deletion blocked: Attempted to delete critical system path - {filename}")
                return redirect(request.referrer or url_for("file_browser"))
     
        if request.form.get("confirm") != "true":
            flash(f"Deletion of '{filename}' requires confirmation.", "warning")
            logger.info(f"Deletion of {filename} requires confirmation")
            return redirect(request.referrer or url_for("file_browser"))
     
        if os.path.isfile(file_path):
            os.remove(file_path)
            flash(f"File '{filename}' deleted successfully.", "success")
            logger.info(f"Deleted file: {file_path}")
        elif os.path.isdir(file_path):
            shutil.rmtree(file_path, ignore_errors=True)
            flash(f"Folder '{filename}' deleted successfully.", "success")
            logger.info(f"Deleted folder: {file_path}")
        else:
            flash(f"Path '{filename}' does not exist.", "error")
            logger.warning(f"Deletion failed: Path does not exist - {file_path}")
     
        return redirect(request.referrer or url_for("file_browser"))
    except Exception as e:
        app.logger.error(f"Delete error: {e}")
        logger.error(f"Delete error for {filename}: {e}")
        flash(f"Failed to delete '{filename}': {str(e)}", "error")
        return redirect(request.referrer or url_for("file_browser"))
@app.route("/delete_file/<path:filename>", methods=["POST"])
@login_required
def delete_file(filename):
    return delete(filename)
@app.route("/logout")
@login_required
def logout():
    session.pop("logged_in", None)
    flash("You have been logged out.", "success")
    return redirect(url_for("login"))
def get_form_parameters():
    """Get parameters from environment variables set by the web form."""
    # EXTENSIVE DEBUGGING OF ENVIRONMENT
    print("🕵️ DEBUG: ENVIRONMENT VARIABLE DUMP:")
    all_env_vars = dict(os.environ)
    for key, value in all_env_vars.items():
        if any(kw in key for kw in ['SELECTED', 'PROCESSING', 'SEMESTER', 'THRESHOLD', 'UPGRADE']):
            print(f" {key}: {value}")
 
    print("🎯 DEBUG - FORM PARAMETERS:")
    print(" SELECTED_SET: {}".format(os.getenv('SELECTED_SET')))
    print(" PROCESSING_MODE: {}".format(os.getenv('PROCESSING_MODE')))
    print(" SELECTED_SEMESTERS: {}".format(os.getenv('SELECTED_SEMESTERS')))
    print(" UPGRADE_THRESHOLD: {}".format(os.getenv('UPGRADE_THRESHOLD')))
 
    selected_set = os.getenv('SELECTED_SET', 'all')
    processing_mode = os.getenv('PROCESSING_MODE', 'auto')
    selected_semesters_str = os.getenv('SELECTED_SEMESTERS', '')
    pass_threshold = float(os.getenv('PASS_THRESHOLD', '50.0'))
    generate_pdf = os.getenv('GENERATE_PDF', 'True').lower() == 'true'
    track_withdrawn = os.getenv('TRACK_WITHDRAWN', 'True').lower() == 'true'
 
    # Convert semester string to list
    selected_semesters = []
    if selected_semesters_str:
        selected_semesters = selected_semesters_str.split(',')
 
    print("🎯 FINAL PARAMETERS:")
    print(" Selected Set: {}".format(selected_set))
    print(" Processing Mode: {}".format(processing_mode))
    print(" Selected Semesters: {}".format(selected_semesters))
    print(" Pass Threshold: {}".format(pass_threshold))
 
    return {
        'selected_set': selected_set,
        'processing_mode': processing_mode,
        'selected_semesters': selected_semesters,
        'pass_threshold': pass_threshold,
        'generate_pdf': generate_pdf,
        'track_withdrawn': track_withdrawn
    }
# ============================================================================
# FIX 2: UPDATED run_script function to handle both GET and POST properly
# ============================================================================
@app.route("/run/<script_name>", methods=["GET", "POST"])
@login_required
def run_script(script_name):
    """Handle both GET (redirect to form) and POST (process form) requests"""
 
    if request.method == "GET":
        # Redirect to appropriate processor page
        if script_name == 'exam_processor_nd':
            return redirect(url_for('nd_regular_exam_processor'))
        elif script_name == 'exam_processor_bn':
            return redirect(url_for('bn_regular_exam_processor'))
        elif script_name == 'exam_processor_bm':
            return redirect(url_for('bm_regular_exam_processor'))
        elif script_name in ['utme', 'caosce', 'clean', 'split']:
            # These scripts run directly without a form
            flash(f"Processing {script_name}...", "info")
        else:
            flash("Invalid script", "error")
            return redirect(url_for("dashboard"))
 
    # Handle POST - process form submission
    try:
        logger.info(f"Processing {script_name} with POST data")
     
        # Handle exam processors
        if script_name in ['exam_processor_nd', 'exam_processor_bn', 'exam_processor_bm']:
            # Extract common parameters
            if script_name == 'exam_processor_bn':
                selected_set = request.form.get('selected_set')
                program = 'BN'
            elif script_name == 'exam_processor_bm':
                selected_set = request.form.get('midwifery_set') or request.form.get('selected_set')
                program = 'BM'
            elif script_name == 'exam_processor_nd':
                selected_set = request.form.get('selected_set')
                program = 'ND'
         
            processing_mode = request.form.get('processing_mode', 'auto')
            selected_semesters = request.form.getlist('selected_semesters')
            pass_threshold = request.form.get('pass_threshold', '50.0')
            upgrade_threshold = request.form.get('upgrade_threshold', '0')
            generate_pdf = request.form.get('generate_pdf') == 'on'
            track_withdrawn = request.form.get('track_withdrawn') == 'on'
         
            # Validation
            if not selected_set:
                flash("Please select an academic set", "error")
                return redirect(request.referrer or url_for("dashboard"))
         
            # Setup environment variables for script
            env = os.environ.copy()
            env["BASE_DIR"] = BASE_DIR
            env["SELECTED_SET"] = selected_set
            env["PROCESSING_MODE"] = processing_mode
            env["PASS_THRESHOLD"] = str(pass_threshold)
            env["UPGRADE_THRESHOLD"] = str(upgrade_threshold)
            env["GENERATE_PDF"] = str(generate_pdf)
            env["TRACK_WITHDRAWN"] = str(track_withdrawn)
         
            if processing_mode == 'manual' and selected_semesters:
                env["SELECTED_SEMESTERS"] = ','.join(selected_semesters)
     
        # Handle other scripts (PUTME, CAOSCE, Internal, JAMB)
        else:
            env = os.environ.copy()
            env["BASE_DIR"] = BASE_DIR
            program = None
            selected_set = None
         
            # Check if files exist
            input_dir = get_input_directory(script_name)
            if not check_input_files(input_dir, script_name):
                flash(f"No input files found for {script_name}", "error")
                return redirect(url_for("dashboard"))
     
        # Get script path
        script_path = _get_script_path(script_name)
     
        # Run script
        if script_name in ['exam_processor_nd', 'exam_processor_bn', 'exam_processor_bm']:
            result = process_script_with_strict_zip_enforcement(
                script_name, program, selected_set, env, script_path
            )
            # Call cleanup after processing
            if result.get("success"):
                clean_up_empty_artifacts(program)
        else:
            # Run other scripts directly
            result = subprocess.run(
                [sys.executable, script_path],
                env=env,
                text=True,
                capture_output=True,
                timeout=600,
            )
         
            if result.returncode == 0:
                output_lines = result.stdout.splitlines()
                processed_files = count_processed_files(output_lines, script_name)
                success_msg = get_success_message(script_name, processed_files, output_lines)
              
                # 🔒 ENFORCE ZIP-ONLY POLICY for all scripts
                clean_dir = get_clean_directory(script_name)
                if os.path.exists(clean_dir):
                    # First try to create ZIPs for scripts that don't auto-create them
                    if script_name in ["caosce", "clean", "split", "utme"]:
                        logger.info(f"🔄 Auto-creating ZIP for {script_name}")
                        script_dirs = {
                            "caosce": os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT"),
                            "clean": os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ"),
                            "split": os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB"),
                            "utme": os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT")
                        }
                        target_dir = script_dirs.get(script_name)
                        if target_dir and os.path.exists(target_dir):
                            # Use the enforce_zip_only_policy with create_zip_if_missing=True
                            enforce_zip_only_policy(target_dir)
               
                # Enforce ZIP-only policy for clean directory
                clean_dir = get_clean_directory(script_name)
                if os.path.exists(clean_dir):
                    enforce_zip_only_policy(clean_dir)
                    logger.info(f"🔒 Enforced ZIP-only policy for {script_name} in {clean_dir}")
               
                flash(success_msg or "Processing completed successfully!", "success")
            else:
                error_msg = result.stderr.splitlines()[-1] if result.stderr else "Unknown error"
                flash(f"Processing failed: {error_msg}", "error")
     
        # Handle exam processor results
        if script_name in ['exam_processor_nd', 'exam_processor_bn', 'exam_processor_bm']:
            if result.get("success"):
                flash(f"{program} examination processing completed successfully!", "success")
            else:
                flash(f"Processing failed: {result.get('error', 'Unknown error')}", "error")
     
        return redirect(url_for("dashboard"))
     
    except subprocess.TimeoutExpired:
        flash("Processing timed out after 10 minutes", "error")
        return redirect(url_for("dashboard"))
    except Exception as e:
        logger.error(f"Error processing {script_name}: {e}")
        import traceback
        traceback.print_exc()
        flash(f"Processing error: {str(e)}", "error")
        return redirect(url_for("dashboard"))
@app.route("/download/<path:filename>")
@login_required
def download(filename):
    abs_path = os.path.join(BASE_DIR, filename)
    if os.path.exists(abs_path) and os.path.isfile(abs_path):
        directory = os.path.dirname(abs_path)
        filename = os.path.basename(abs_path)
        return send_from_directory(directory, filename, as_attachment=True)
    else:
        flash("File not found", "error")
        return redirect(request.referrer)
# ============================================================================
# FIX 2: Added missing route for download_zip
# ============================================================================
@app.route("/download_zip/<set_name>")
@login_required
def download_zip(set_name):
    """Download all files for a set as a ZIP"""
    try:
        # Determine program from set name
        program = detect_program_from_set(set_name)
        if not program:
            flash(f"Could not determine program for set {set_name}", "error")
            return redirect(url_for("download_center"))
     
        clean_dir = os.path.join(BASE_DIR, program, set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            flash(f"No results found for {set_name}", "error")
            return redirect(url_for("download_center"))
     
        # Find existing ZIP file
        zip_files = [f for f in os.listdir(clean_dir)
                    if f.startswith(f"{set_name}_RESULT-") and f.endswith('.zip')]
     
        if zip_files:
            latest_zip = sorted(zip_files)[-1]
            zip_path = os.path.join(clean_dir, latest_zip)
            return send_file(zip_path, as_attachment=True, download_name=latest_zip)
        else:
            flash(f"No ZIP file found for {set_name}", "error")
            return redirect(url_for("download_center"))
         
    except Exception as e:
        logger.error(f"Download ZIP error: {e}")
        flash(f"Error downloading ZIP: {str(e)}", "error")
        return redirect(url_for("download_center"))
# ============================================================================
# NEW: PUTME Processing Route with Proper Form Handling
# ============================================================================
@app.route("/putme_processor", methods=["GET", "POST"])
@login_required
def putme_processor():
    """PUTME Results processor form and handler"""
    try:
        if request.method == "GET":
            return render_template(
                "putme_processor.html",
                college=COLLEGE,
                department=DEPARTMENT,
                environment="Railway Production" if not is_local_environment() else "Local Development"
            )
     
        # POST - Handle form submission
        convert_column = request.form.get("convert_column", "n")
        convert_value = request.form.get("convert_value", "")
     
        # Validate input directory
        input_dir = os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_PUTME_RESULT")
        candidate_dir = os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_CANDIDATE_BATCHES")
     
        if not check_putme_files(input_dir):
            flash("No PUTME files found in RAW_PUTME_RESULT directory", "error")
            return redirect(url_for("putme_processor"))
     
        # Setup environment
        env = os.environ.copy()
        env["BASE_DIR"] = BASE_DIR
     
        # Build command arguments
        script_path = os.path.join(SCRIPT_DIR, "utme_result.py")
        cmd = [sys.executable, script_path]
     
        # Add arguments
        cmd.extend(["--input-dir", input_dir])
        cmd.extend(["--candidate-dir", candidate_dir])
        cmd.extend(["--output-dir", os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT")])
     
        # Handle conversion
        if convert_column == "y" and convert_value:
            try:
                max_score = int(convert_value)
                cmd.extend(["--non-interactive"])
                cmd.extend(["--converted-score-max", str(max_score)])
                logger.info(f"Adding score conversion: Score/{max_score}")
            except ValueError:
                flash(f"Invalid conversion value: {convert_value}", "warning")
        else:
            cmd.extend(["--non-interactive"])
     
        logger.info(f"Running PUTME command: {' '.join(cmd)}")
     
        # Run script
        result = subprocess.run(
            cmd,
            env=env,
            text=True,
            capture_output=True,
            timeout=600,
        )
     
        output_lines = result.stdout.splitlines()
        error_lines = result.stderr.splitlines()
     
        # Log output
        logger.info("=== PUTME PROCESSING OUTPUT ===")
        for line in output_lines:
            logger.info(line)
     
        if error_lines:
            logger.error("=== PUTME PROCESSING ERRORS ===")
            for line in error_lines:
                logger.error(line)
     
        # Check success
        if result.returncode == 0:
            processed_files = count_processed_files(output_lines, "utme")
            success_msg = get_success_message("utme", processed_files, output_lines)
         
            # Enforce ZIP-only policy
            clean_dir = os.path.join(BASE_DIR, "PUTME_RESULT", "CLEAN_PUTME_RESULT")
            if os.path.exists(clean_dir):
                enforce_zip_only_policy(clean_dir)
         
            flash(success_msg or "PUTME processing completed successfully!", "success")
        else:
            error_msg = error_lines[-1] if error_lines else "Unknown error"
            flash(f"PUTME processing failed: {error_msg}", "error")
     
        return redirect(url_for("putme_processor"))
     
    except subprocess.TimeoutExpired:
        flash("PUTME processing timed out after 10 minutes", "error")
        return redirect(url_for("putme_processor"))
    except Exception as e:
        logger.error(f"PUTME processing error: {e}")
        import traceback
        traceback.print_exc()
        flash(f"Error: {str(e)}", "error")
        return redirect(url_for("putme_processor"))
# ============================================================================
# NEW: Debug route for PUTME files
# ============================================================================
@app.route("/debug_putme_files")
@login_required
def debug_putme_files():
    """Debug route to check PUTME directory structure and files"""
    input_dir = os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_PUTME_RESULT")
    candidate_dir = os.path.join(BASE_DIR, "PUTME_RESULT", "RAW_CANDIDATE_BATCHES")
 
    return jsonify({
        'input_dir': input_dir,
        'input_exists': os.path.exists(input_dir),
        'input_files': os.listdir(input_dir) if os.path.exists(input_dir) else [],
        'candidate_dir': candidate_dir,
        'candidate_exists': os.path.exists(candidate_dir),
        'candidate_files': os.listdir(candidate_dir) if os.path.exists(candidate_dir) else []
    })
# ============================================================================
# NEW: Routes for other script processors (CAOSCE, Internal, JAMB)
# ============================================================================
@app.route("/caosce_processor", methods=["GET", "POST"])
@login_required
def caosce_processor():
    """CAOSCE Results processor form and handler"""
    try:
        if request.method == "GET":
            return render_template(
                "caosce_processor.html",
                college=COLLEGE,
                department=DEPARTMENT,
                environment="Railway Production" if not is_local_environment() else "Local Development"
            )
     
        # POST - Handle form submission
        # Validate input directory
        input_dir = os.path.join(BASE_DIR, "CAOSCE_RESULT", "RAW_CAOSCE_RESULT")
     
        if not check_caosce_files(input_dir):
            flash("No CAOSCE files found in RAW_CAOSCE_RESULT directory", "error")
            return redirect(url_for("caosce_processor"))
     
        # Setup environment
        env = os.environ.copy()
        env["BASE_DIR"] = BASE_DIR
     
        # Get script path
        script_path = os.path.join(SCRIPT_DIR, "caosce_result.py")
     
        # Run script
        result = subprocess.run(
            [sys.executable, script_path],
            env=env,
            text=True,
            capture_output=True,
            timeout=600,
        )
     
        output_lines = result.stdout.splitlines()
        error_lines = result.stderr.splitlines()
     
        # Log output
        logger.info("=== CAOSCE PROCESSING OUTPUT ===")
        for line in output_lines:
            logger.info(line)
     
        if error_lines:
            logger.error("=== CAOSCE PROCESSING ERRORS ===")
            for line in error_lines:
                logger.error(line)
     
        # Check success
        if result.returncode == 0:
            processed_files = count_processed_files(output_lines, "caosce")
            success_msg = get_success_message("caosce", processed_files, output_lines)
         
            # Enforce ZIP-only policy
            clean_dir = os.path.join(BASE_DIR, "CAOSCE_RESULT", "CLEAN_CAOSCE_RESULT")
            if os.path.exists(clean_dir):
                enforce_zip_only_policy(clean_dir)
         
            flash(success_msg or "CAOSCE processing completed successfully!", "success")
        else:
            error_msg = error_lines[-1] if error_lines else "Unknown error"
            flash(f"CAOSCE processing failed: {error_msg}", "error")
     
        return redirect(url_for("caosce_processor"))
     
    except subprocess.TimeoutExpired:
        flash("CAOSCE processing timed out after 10 minutes", "error")
        return redirect(url_for("caosce_processor"))
    except Exception as e:
        logger.error(f"CAOSCE processing error: {e}")
        import traceback
        traceback.print_exc()
        flash(f"Error: {str(e)}", "error")
        return redirect(url_for("caosce_processor"))
@app.route("/internal_processor", methods=["GET", "POST"])
@login_required
def internal_processor():
    """Internal Results processor form and handler"""
    try:
        if request.method == "GET":
            return render_template(
                "internal_processor.html",
                college=COLLEGE,
                department=DEPARTMENT,
                environment="Railway Production" if not is_local_environment() else "Local Development"
            )
     
        # POST - Handle form submission
        # Validate input directory
        input_dir = os.path.join(BASE_DIR, "OBJ_RESULT", "RAW_OBJ")
     
        if not check_internal_exam_files(input_dir):
            flash("No internal exam files found in RAW_OBJ directory", "error")
            return redirect(url_for("internal_processor"))
     
        # Setup environment
        env = os.environ.copy()
        env["BASE_DIR"] = BASE_DIR
     
        # Get script path
        script_path = os.path.join(SCRIPT_DIR, "obj_results.py")
     
        # Run script
        result = subprocess.run(
            [sys.executable, script_path],
            env=env,
            text=True,
            capture_output=True,
            timeout=600,
        )
     
        output_lines = result.stdout.splitlines()
        error_lines = result.stderr.splitlines()
     
        # Log output
        logger.info("=== INTERNAL EXAM PROCESSING OUTPUT ===")
        for line in output_lines:
            logger.info(line)
     
        if error_lines:
            logger.error("=== INTERNAL EXAM PROCESSING ERRORS ===")
            for line in error_lines:
                logger.error(line)
     
        # Check success
        if result.returncode == 0:
            processed_files = count_processed_files(output_lines, "clean")
            success_msg = get_success_message("clean", processed_files, output_lines)
         
            # Enforce ZIP-only policy
            clean_dir = os.path.join(BASE_DIR, "OBJ_RESULT", "CLEAN_OBJ")
            if os.path.exists(clean_dir):
                enforce_zip_only_policy(clean_dir)
         
            flash(success_msg or "Internal exam processing completed successfully!", "success")
        else:
            error_msg = error_lines[-1] if error_lines else "Unknown error"
            flash(f"Internal exam processing failed: {error_msg}", "error")
     
        return redirect(url_for("internal_processor"))
     
    except subprocess.TimeoutExpired:
        flash("Internal exam processing timed out after 10 minutes", "error")
        return redirect(url_for("internal_processor"))
    except Exception as e:
        logger.error(f"Internal exam processing error: {e}")
        import traceback
        traceback.print_exc()
        flash(f"Error: {str(e)}", "error")
        return redirect(url_for("internal_processor"))
@app.route("/jamb_processor", methods=["GET", "POST"])
@login_required
def jamb_processor():
    """JAMB Results processor form and handler"""
    try:
        if request.method == "GET":
            return render_template(
                "jamb_processor.html",
                college=COLLEGE,
                department=DEPARTMENT,
                environment="Railway Production" if not is_local_environment() else "Local Development"
            )
     
        # POST - Handle form submission
        # Validate input directory
        input_dir = os.path.join(BASE_DIR, "JAMB_DB", "RAW_JAMB_DB")
     
        if not check_split_files(input_dir):
            flash("No JAMB files found in RAW_JAMB_DB directory", "error")
            return redirect(url_for("jamb_processor"))
     
        # Setup environment
        env = os.environ.copy()
        env["BASE_DIR"] = BASE_DIR
     
        # Get script path
        script_path = os.path.join(SCRIPT_DIR, "split_names.py")
     
        # Run script
        result = subprocess.run(
            [sys.executable, script_path],
            env=env,
            text=True,
            capture_output=True,
            timeout=600,
        )
     
        output_lines = result.stdout.splitlines()
        error_lines = result.stderr.splitlines()
     
        # Log output
        logger.info("=== JAMB PROCESSING OUTPUT ===")
        for line in output_lines:
            logger.info(line)
     
        if error_lines:
            logger.error("=== JAMB PROCESSING ERRORS ===")
            for line in error_lines:
                logger.error(line)
     
        # Check success
        if result.returncode == 0:
            processed_files = count_processed_files(output_lines, "split")
            success_msg = get_success_message("split", processed_files, output_lines)
         
            # Enforce ZIP-only policy
            clean_dir = os.path.join(BASE_DIR, "JAMB_DB", "CLEAN_JAMB_DB")
            if os.path.exists(clean_dir):
                enforce_zip_only_policy(clean_dir)
         
            flash(success_msg or "JAMB processing completed successfully!", "success")
        else:
            error_msg = error_lines[-1] if error_lines else "Unknown error"
            flash(f"JAMB processing failed: {error_msg}", "error")
     
        return redirect(url_for("jamb_processor"))
     
    except subprocess.TimeoutExpired:
        flash("JAMB processing timed out after 10 minutes", "error")
        return redirect(url_for("jamb_processor"))
    except Exception as e:
        logger.error(f"JAMB processing error: {e}")
        import traceback
        traceback.print_exc()
        flash(f"Error: {str(e)}", "error")
        return redirect(url_for("jamb_processor"))
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    mode = "local" if is_local_environment() else "cloud"
    logger.info(f"Starting Flask app in {mode.upper()} mode on port {port}...")
    app.run(host="0.0.0.0", port=port, debug=True)