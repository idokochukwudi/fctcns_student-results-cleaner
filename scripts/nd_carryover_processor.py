#!/usr/bin/env python3
"""
nd_carryover_processor.py - ND-ONLY CARRYOVER PROCESSOR
Fixed Issues:
1. Removed all BN/BM logic and references
2. Fixed typo on line 262: .ast(str) -> .astype(str) 
3. Enhanced course code normalization for better matching
4. Added comprehensive debug logging
5. Improved course title/unit lookup with fallback strategies
6. Fixed main() function with proper environment variable handling
"""

import os
import sys
import re
import pandas as pd
import numpy as np
from datetime import datetime
import glob
import json
import traceback
import shutil
import zipfile
import tempfile
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def standardize_semester_key(semester_key):
    """Standardize semester key to canonical format - ND ONLY."""
    if not semester_key:
        return None
    
    key_upper = semester_key.upper()
    
    # ND-ONLY canonical mappings
    canonical_mappings = {
        # ND First Year First Semester variants
        ("FIRST", "YEAR", "FIRST", "SEMESTER"): "ND-FIRST-YEAR-FIRST-SEMESTER",
        ("1ST", "YEAR", "1ST", "SEMESTER"): "ND-FIRST-YEAR-FIRST-SEMESTER",
        ("YEAR", "1", "SEMESTER", "1"): "ND-FIRST-YEAR-FIRST-SEMESTER",
        
        # ND First Year Second Semester variants
        ("FIRST", "YEAR", "SECOND", "SEMESTER"): "ND-FIRST-YEAR-SECOND-SEMESTER",
        ("1ST", "YEAR", "2ND", "SEMESTER"): "ND-FIRST-YEAR-SECOND-SEMESTER",
        ("YEAR", "1", "SEMESTER", "2"): "ND-FIRST-YEAR-SECOND-SEMESTER",
        
        # ND Second Year First Semester variants
        ("SECOND", "YEAR", "FIRST", "SEMESTER"): "ND-SECOND-YEAR-FIRST-SEMESTER",
        ("2ND", "YEAR", "1ST", "SEMESTER"): "ND-SECOND-YEAR-FIRST-SEMESTER",
        ("YEAR", "2", "SEMESTER", "1"): "ND-SECOND-YEAR-FIRST-SEMESTER",
        
        # ND Second Year Second Semester variants
        ("SECOND", "YEAR", "SECOND", "SEMESTER"): "ND-SECOND-YEAR-SECOND-SEMESTER",
        ("2ND", "YEAR", "2ND", "SEMESTER"): "ND-SECOND-YEAR-SECOND-SEMESTER",
        ("YEAR", "2", "SEMESTER", "2"): "ND-SECOND-YEAR-SECOND-SEMESTER",
    }
    
    # ND-only patterns
    patterns = [
        r'(FIRST|1ST|YEAR.?1).*?(FIRST|1ST|SEMESTER.?1)',
        r'(FIRST|1ST|YEAR.?1).*?(SECOND|2ND|SEMESTER.?2)',
        r'(SECOND|2ND|YEAR.?2).*?(FIRST|1ST|SEMESTER.?1)',
        r'(SECOND|2ND|YEAR.?2).*?(SECOND|2ND|SEMESTER.?2)',
    ]
    
    for pattern_idx, pattern in enumerate(patterns):
        if re.search(pattern, key_upper):
            if pattern_idx == 0:
                return "ND-FIRST-YEAR-FIRST-SEMESTER"
            elif pattern_idx == 1:
                return "ND-FIRST-YEAR-SECOND-SEMESTER"
            elif pattern_idx == 2:
                return "ND-SECOND-YEAR-FIRST-SEMESTER"
            elif pattern_idx == 3:
                return "ND-SECOND-YEAR-SECOND-SEMESTER"
    
    # If no match, return original
    print(f"Could not standardize semester key: {semester_key}")
    return semester_key

def get_previous_semester(semester_key):
    """Get the previous semester key for ND carryover ONLY."""
    standardized = standardize_semester_key(semester_key)
    
    # ND semesters ONLY
    if standardized == "ND-FIRST-YEAR-SECOND-SEMESTER":
        return "ND-FIRST-YEAR-FIRST-SEMESTER"
    elif standardized == "ND-SECOND-YEAR-FIRST-SEMESTER":
        return "ND-FIRST-YEAR-SECOND-SEMESTER"
    elif standardized == "ND-SECOND-YEAR-SECOND-SEMESTER":
        return "ND-SECOND-YEAR-FIRST-SEMESTER"
    else:
        return None  # No previous for first semester

# ----------------------------
# Configuration
# ----------------------------
def get_base_directory():
    """Get base directory - ENHANCED VERSION."""
    # First try environment variable
    if os.getenv('BASE_DIR'):
        base_dir = os.getenv('BASE_DIR')
        if os.path.exists(base_dir):
            return base_dir
    
    # Try the user's home directory with student_result_cleaner
    home_dir = os.path.expanduser('~')
    default_dir = os.path.join(home_dir, 'student_result_cleaner')
    
    # Check if EXAMS_INTERNAL exists in the default directory
    if os.path.exists(os.path.join(default_dir, "EXAMS_INTERNAL")):
        return default_dir
    
    # If not, check if we're already in a directory that contains EXAMS_INTERNAL
    current_script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.exists(os.path.join(current_script_dir, "EXAMS_INTERNAL")):
        return current_script_dir
    
    # Check parent directory
    parent_dir = os.path.dirname(current_script_dir)
    if os.path.exists(os.path.join(parent_dir, "EXAMS_INTERNAL")):
        return parent_dir
    
    # Final fallback
    return default_dir

BASE_DIR = get_base_directory()
TIMESTAMP_FMT = "%d-%m-%Y_%H%M%S"
DEFAULT_PASS_THRESHOLD = 50.0
DEFAULT_LOGO_PATH = os.path.join(os.path.dirname(__file__), "logo.png")

def sanitize_filename(filename):
    """Remove or replace characters that are not safe for filenames."""
    return re.sub(r'[^\w\-_.]', '_', filename)

def find_exam_number_column(df):
    """Find the exam number column in a DataFrame."""
    possible_names = ['EXAM NUMBER', 'REG. No', 'REG NO', 'REGISTRATION NUMBER', 'MAT NO', 'STUDENT ID']
    for col in df.columns:
        col_upper = str(col).upper()
        for possible_name in possible_names:
            if possible_name in col_upper:
                return col
    return None

def load_course_data():
    """Load ND course data ONLY."""
    return load_nd_course_data()

def load_nd_course_data():
    """Load ND course data from course-code-creditUnit.xlsx."""
    possible_course_files = [
        os.path.join(BASE_DIR, "EXAMS_INTERNAL", "ND", "ND-COURSES", "course-code-creditUnit.xlsx"),
        os.path.join(BASE_DIR, "ND", "ND-COURSES", "course-code-creditUnit.xlsx"),
        os.path.join(BASE_DIR, "EXAMS_INTERNAL", "ND-COURSES", "course-code-creditUnit.xlsx"),
        os.path.join(BASE_DIR, "course-code-creditUnit.xlsx"),
    ]
    
    course_file = None
    for possible_file in possible_course_files:
        if os.path.exists(possible_file):
            course_file = possible_file
            print(f"✅ Found ND course file: {course_file}")
            break
    
    if not course_file:
        print(f"❌ Main ND course file not found in standard locations")
        alternative_files = find_alternative_course_files()
        if alternative_files:
            course_file = alternative_files[0]
            print(f"🔄 Using alternative ND course file: {course_file}")
        else:
            print("❌ No ND course files found anywhere!")
            return {}, {}, {}, {}
    
    print(f"📚 Loading ND course data from: {course_file}")
    return _load_course_data_from_file(course_file)

def _load_course_data_from_file(course_file):
    """Generic function to load course data from Excel file - FIXED VERSION."""
    try:
        xl = pd.ExcelFile(course_file)
        semester_course_titles = {}
        semester_credit_units = {}
        course_code_to_title = {}
        course_code_to_unit = {}
        
        print(f"📖 Available sheets: {xl.sheet_names}")
        
        for sheet in xl.sheet_names:
            sheet_standard = standardize_semester_key(sheet)
            print(f"📖 Reading sheet: {sheet} (standardized: {sheet_standard})")
            try:
                df = pd.read_excel(course_file, sheet_name=sheet, engine='openpyxl', header=0)
                
                # Convert columns to string and clean
                df.columns = [str(c).strip().upper() for c in df.columns]
                
                # Look for course code, title, and credit unit columns with flexible matching
                code_col = None
                title_col = None
                unit_col = None
                
                for col in df.columns:
                    col_clean = str(col).upper()
                    if any(keyword in col_clean for keyword in ['COURSE CODE', 'CODE', 'COURSECODE']):
                        code_col = col
                    elif any(keyword in col_clean for keyword in ['COURSE TITLE', 'TITLE', 'COURSENAME']):
                        title_col = col
                    elif any(keyword in col_clean for keyword in ['CU', 'CREDIT', 'UNIT', 'CREDIT UNIT']):
                        unit_col = col
                
                print(f"🔍 Detected columns - Code: {code_col}, Title: {title_col}, Unit: {unit_col}")
                
                if not all([code_col, title_col, unit_col]):
                    print(f"⚠️ Sheet '{sheet}' missing required columns - found: code={code_col}, title={title_col}, unit={unit_col}")
                    # Try to use first three columns as fallback
                    if len(df.columns) >= 3:
                        code_col, title_col, unit_col = df.columns[0], df.columns[1], df.columns[2]
                        print(f"🔄 Using fallback columns: {code_col}, {title_col}, {unit_col}")
                    else:
                        print(f"❌ Sheet '{sheet}' doesn't have enough columns - skipped")
                        continue
                
                # Clean the data
                df_clean = df.dropna(subset=[code_col]).copy()
                if df_clean.empty:
                    print(f"⚠️ Sheet '{sheet}' has no data after cleaning - skipped")
                    continue
                
                # Convert credit units to numeric, handling errors
                df_clean[unit_col] = pd.to_numeric(df_clean[unit_col], errors='coerce')
                df_clean = df_clean.dropna(subset=[unit_col])
                
                # FIXED: Remove rows with "TOTAL" in course code (typo fixed: .ast -> .astype)
                df_clean = df_clean[~df_clean[code_col].astype(str).str.contains('TOTAL', case=False, na=False)]
                
                if df_clean.empty:
                    print(f"⚠️ Sheet '{sheet}' has no valid rows after cleaning - skipped")
                    continue
                
                codes = df_clean[code_col].astype(str).str.strip().tolist()
                titles = df_clean[title_col].astype(str).str.strip().tolist()
                units = df_clean[unit_col].astype(float).tolist()

                print(f"📋 Found {len(codes)} courses in {sheet}:")
                for i, (code, title, unit) in enumerate(zip(codes[:5], titles[:5], units[:5])):
                    print(f"   - '{code}': '{title}' (CU: {unit})")

                # Create mapping dictionaries with ENHANCED normalization strategies
                sheet_titles = {}
                sheet_units = {}
                
                for code, title, unit in zip(codes, titles, units):
                    if not code or code.upper() in ['NAN', 'NONE', '']:
                        continue
                    
                    # ENHANCED: Create comprehensive normalization variants for robust matching
                    variants = [
                        # Basic variants
                        code.upper().strip(),
                        code.strip(),
                        code.upper(),
                        code.lower(),
                        code.title(),
                        
                        # Space removal variants
                        code.upper().replace(' ', ''),
                        code.replace(' ', ''),
                        re.sub(r'\s+', '', code.upper()),
                        re.sub(r'\s+', '', code),
                        
                        # Special character removal (keep only alphanumeric)
                        re.sub(r'[^a-zA-Z0-9]', '', code.upper()),
                        re.sub(r'[^a-zA-Z0-9]', '', code),
                        
                        # Dash and underscore variants
                        code.upper().replace('-', ''),
                        code.upper().replace('_', ''),
                        code.replace('-', '').replace('_', ''),
                        code.upper().replace('-', '').replace('_', '').replace(' ', ''),
                        
                        # WITH common prefixes (for matching with prefix)
                        f"NUR{code.upper()}",
                        f"NUR{code.upper().replace(' ', '')}",
                        f"NUR{re.sub(r'[^a-zA-Z0-9]', '', code.upper())}",
                        f"NSC{code.upper()}",
                        f"NSC{code.upper().replace(' ', '')}",
                        f"NSC{re.sub(r'[^a-zA-Z0-9]', '', code.upper())}",
                        
                        # WITHOUT common prefixes (for matching without prefix)
                        code.upper().replace('NUR', '').strip(),
                        code.upper().replace('NSC', '').strip(),
                        re.sub(r'^(NUR|NSC)', '', code.upper()).strip(),
                        re.sub(r'^(NUR|NSC)', '', code.upper()).replace(' ', '').strip(),
                        
                        # Number-focused variants (for codes like "101", "201")
                        re.sub(r'[^0-9]', '', code),
                        
                        # Common variations with dots
                        code.upper().replace('.', ''),
                        code.replace('.', ''),
                    ]
                    
                    # Remove duplicates while preserving order
                    variants = list(dict.fromkeys([v for v in variants if v and v not in ['NAN', 'NONE', '']]))
                    
                    # Add all variants to mappings
                    for variant in variants:
                        sheet_titles[variant] = title
                        sheet_units[variant] = unit
                        course_code_to_title[variant] = title
                        course_code_to_unit[variant] = unit
                
                semester_course_titles[sheet_standard] = sheet_titles
                semester_credit_units[sheet_standard] = sheet_units
                
            except Exception as e:
                print(f"❌ Error processing sheet '{sheet}': {e}")
                traceback.print_exc()
                continue
        
        print(f"✅ Loaded course data for sheets: {list(semester_course_titles.keys())}")
        print(f"📊 Total course mappings: {len(course_code_to_title)}")
        
        # Debug: Show some course mappings
        print("🔍 Sample course mappings:")
        sample_items = list(course_code_to_title.items())[:15]
        for code, title in sample_items:
            unit = course_code_to_unit.get(code, 0)
            print(f"   '{code}' -> '{title}' (CU: {unit})")
            
        return semester_course_titles, semester_credit_units, course_code_to_title, course_code_to_unit
        
    except Exception as e:
        print(f"❌ Error loading course data: {e}")
        traceback.print_exc()
        return {}, {}, {}, {}

def find_alternative_course_files():
    """Look for alternative course files for ND."""
    base_dirs = [
        os.path.join(BASE_DIR, "EXAMS_INTERNAL", "ND", "ND-COURSES"),
        os.path.join(BASE_DIR, "ND", "ND-COURSES"),
        os.path.join(BASE_DIR, "COURSES"),
        os.path.join(BASE_DIR, "EXAMS_INTERNAL"),
    ]
    
    course_files = []
    for base_dir in base_dirs:
        if os.path.exists(base_dir):
            for file in os.listdir(base_dir):
                if 'course' in file.lower() and file.endswith(('.xlsx', '.xls')):
                    full_path = os.path.join(base_dir, file)
                    course_files.append(full_path)
                    print(f"📁 Found ND course file: {full_path}")
    
    return course_files

def debug_course_matching(resit_file_path, course_code_to_title, course_code_to_unit):
    """Debug function to check why course codes aren't matching - ENHANCED."""
    print(f"\n🔍 DEBUGGING ND COURSE MATCHING")
    print("=" * 50)
    
    # Read resit file to see what course codes we have
    resit_df = pd.read_excel(resit_file_path, header=0)
    resit_exam_col = find_exam_number_column(resit_df)
    
    # Get all course codes from resit file
    resit_courses = []
    for col in resit_df.columns:
        if col != resit_exam_col and col != 'NAME' and not 'Unnamed' in str(col):
            resit_courses.append(col)
    
    print(f"📋 ND Course codes from resit file: {resit_courses}")
    print(f"📊 Total courses in ND resit file: {len(resit_courses)}")
    
    # Check each resit course against course file
    for course in resit_courses:
        print(f"\n🔍 Checking ND course: '{course}'")
        original_code = str(course).strip()
        
        # Generate ENHANCED variants for matching
        variants = [
            original_code.upper().strip(),
            original_code.strip(),
            original_code.upper(),
            original_code,
            original_code.lower(),
            original_code.title(),
            original_code.upper().replace(' ', ''),
            original_code.replace(' ', ''),
            re.sub(r'\s+', '', original_code.upper()),
            re.sub(r'\s+', '', original_code),
            re.sub(r'[^a-zA-Z0-9]', '', original_code.upper()),
            re.sub(r'[^a-zA-Z0-9]', '', original_code),
            original_code.upper().replace('-', ''),
            original_code.upper().replace('_', ''),
            original_code.replace('-', '').replace('_', ''),
            original_code.upper().replace('-', '').replace('_', '').replace(' ', ''),
            f"NUR{original_code.upper()}",
            f"NUR{original_code.upper().replace(' ', '')}",
            f"NSC{original_code.upper()}",
            f"NSC{original_code.upper().replace(' ', '')}",
            original_code.upper().replace('NUR', '').strip(),
            original_code.upper().replace('NSC', '').strip(),
            re.sub(r'^(NUR|NSC)', '', original_code.upper()).strip(),
            re.sub(r'[^0-9]', '', original_code),
            original_code.upper().replace('.', ''),
            original_code.replace('.', ''),
        ]
        
        # Remove duplicates
        variants = list(dict.fromkeys([v for v in variants if v and v != 'NAN']))
        
        print(f"   Generated {len(variants)} variants to try")
        
        found = False
        for variant in variants:
            if variant in course_code_to_title:
                title = course_code_to_title[variant]
                unit = course_code_to_unit.get(variant, 0)
                print(f"   ✅ FOUND: '{variant}' -> '{title}' (CU: {unit})")
                found = True
                break
        
        if not found:
            print(f"   ❌ NOT FOUND: No match for '{course}'")
            # Show some similar keys from course file
            similar_keys = []
            # Check for partial matches
            for key in list(course_code_to_title.keys())[:50]:  # Check first 50 keys
                if any(part in key.upper() for part in original_code.upper().split() if len(part) > 2):
                    similar_keys.append(key)
            
            if similar_keys:
                print(f"   💡 Similar keys found: {similar_keys[:5]}")
            else:
                print(f"   💡 Sample available keys: {list(course_code_to_title.keys())[:10]}")

def find_course_title(course_code, course_titles_dict, course_code_to_title):
    """Robust function to find course title with comprehensive matching strategies - ENHANCED."""
    if not course_code or str(course_code).upper() in ['NAN', 'NONE', '']:
        return str(course_code) if course_code else "Unknown Course"
    
    original_code = str(course_code).strip()
    
    # Generate ENHANCED comprehensive matching variants
    variants = [
        # Basic normalizations
        original_code.upper().strip(),
        original_code.strip(),
        original_code.upper(),
        original_code,
        original_code.lower(),
        original_code.title(),
        
        # Space handling variations
        original_code.upper().replace(' ', ''),
        original_code.replace(' ', ''),
        re.sub(r'\s+', '', original_code.upper()),
        re.sub(r'\s+', '', original_code),
        
        # Special character handling
        re.sub(r'[^a-zA-Z0-9]', '', original_code.upper()),
        re.sub(r'[^a-zA-Z0-9]', '', original_code),
        
        # Common formatting issues
        original_code.upper().replace('-', ''),
        original_code.upper().replace('_', ''),
        original_code.replace('-', '').replace('_', ''),
        original_code.upper().replace('-', '').replace('_', '').replace(' ', ''),
        
        # WITH common prefixes
        f"NUR{original_code.upper()}",
        f"NUR{original_code.upper().replace(' ', '')}",
        f"NUR{re.sub(r'[^a-zA-Z0-9]', '', original_code.upper())}",
        f"NSC{original_code.upper()}",
        f"NSC{original_code.upper().replace(' ', '')}",
        f"NSC{re.sub(r'[^a-zA-Z0-9]', '', original_code.upper())}",
        
        # WITHOUT common prefixes
        original_code.upper().replace('NUR', '').strip(),
        original_code.upper().replace('NSC', '').strip(),
        re.sub(r'^(NUR|NSC)', '', original_code.upper()).strip(),
        re.sub(r'^(NUR|NSC)', '', original_code.upper()).replace(' ', '').strip(),
        
        # Number-focused variants
        re.sub(r'[^0-9]', '', original_code),
        
        # Dot removal
        original_code.upper().replace('.', ''),
        original_code.replace('.', ''),
    ]
    
    # Remove duplicates
    variants = list(dict.fromkeys([v for v in variants if v and v != 'NAN']))
    
    # Try each strategy in order
    for variant in variants:
        # Try course_titles_dict first (semester-specific)
        if variant in course_titles_dict:
            title = course_titles_dict[variant]
            print(f"✅ Found title for '{original_code}' using variant '{variant}': '{title}'")
            return title
        
        # Try global course_code_to_title
        if variant in course_code_to_title:
            title = course_code_to_title[variant]
            print(f"✅ Found title for '{original_code}' using global variant '{variant}': '{title}'")
            return title
    
    # If no match found, log and return descriptive original code
    print(f"⚠️ Could not find course title for: '{original_code}'")
    print(f"   Tried {len(variants)} variants without success")
    return f"{original_code} (Title Not Found)"

def find_credit_unit(course_code, credit_units_dict, course_code_to_unit):
    """Robust function to find credit unit with comprehensive matching strategies - ENHANCED."""
    if not course_code or str(course_code).upper() in ['NAN', 'NONE', '']:
        return 0
    
    original_code = str(course_code).strip()
    
    # Generate the same ENHANCED variants as title matching
    variants = [
        original_code.upper().strip(),
        original_code.strip(),
        original_code.upper(),
        original_code,
        original_code.lower(),
        original_code.title(),
        original_code.upper().replace(' ', ''),
        original_code.replace(' ', ''),
        re.sub(r'\s+', '', original_code.upper()),
        re.sub(r'\s+', '', original_code),
        re.sub(r'[^a-zA-Z0-9]', '', original_code.upper()),
        re.sub(r'[^a-zA-Z0-9]', '', original_code),
        original_code.upper().replace('-', ''),
        original_code.upper().replace('_', ''),
        original_code.replace('-', '').replace('_', ''),
        original_code.upper().replace('-', '').replace('_', '').replace(' ', ''),
        f"NUR{original_code.upper()}",
        f"NUR{original_code.upper().replace(' ', '')}",
        f"NUR{re.sub(r'[^a-zA-Z0-9]', '', original_code.upper())}",
        f"NSC{original_code.upper()}",
        f"NSC{original_code.upper().replace(' ', '')}",
        f"NSC{re.sub(r'[^a-zA-Z0-9]', '', original_code.upper())}",
        original_code.upper().replace('NUR', '').strip(),
        original_code.upper().replace('NSC', '').strip(),
        re.sub(r'^(NUR|NSC)', '', original_code.upper()).strip(),
        re.sub(r'^(NUR|NSC)', '', original_code.upper()).replace(' ', '').strip(),
        re.sub(r'[^0-9]', '', original_code),
        original_code.upper().replace('.', ''),
        original_code.replace('.', ''),
    ]
    
    # Remove duplicates
    variants = list(dict.fromkeys([v for v in variants if v and v != 'NAN']))
    
    # Try each strategy
    for variant in variants:
        if variant in credit_units_dict:
            unit = credit_units_dict[variant]
            return unit
        
        if variant in course_code_to_unit:
            unit = course_code_to_unit[variant]
            return unit
    
    print(f"⚠️ Could not find credit unit for: '{original_code}', defaulting to 2")
    return 2  # Default credit unit

def get_semester_display_info(semester_key):
    """Get display information for ND semester key ONLY."""
    semester_lower = semester_key.lower()
    
    # ND ONLY
    if 'first-year-first-semester' in semester_lower:
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI", "Semester 1"
    elif 'first-year-second-semester' in semester_lower:
        return 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI", "Semester 2"
    elif 'second-year-first-semester' in semester_lower:
        return 2, 1, "YEAR TWO", "FIRST SEMESTER", "NDII", "Semester 3"
    elif 'second-year-second-semester' in semester_lower:
        return 2, 2, "YEAR TWO", "SECOND SEMESTER", "NDII", "Semester 4"
    else:
        return 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI", "Semester 1"

def get_grade_point(score):
    """Determine grade point based on score - NIGERIAN 5.0 SCALE."""
    try:
        score = float(score)
        if score >= 70: return 5.0
        elif score >= 60: return 4.0
        elif score >= 50: return 3.0
        elif score >= 45: return 2.0
        elif score >= 40: return 1.0
        else: return 0.0
    except (ValueError, TypeError):
        return 0.0

def extract_mastersheet_from_zip(zip_path, semester_key):
    """Extract mastersheet from ZIP file and return temporary file path."""
    try:
        print(f"📦 Looking for mastersheet in ZIP: {zip_path}")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            all_files = zip_ref.namelist()
            print(f"📁 Files in ZIP: {all_files}")
            
            mastersheet_files = [f for f in all_files if 'mastersheet' in f.lower() and f.endswith('.xlsx')]
            
            if not mastersheet_files:
                print(f"❌ No mastersheet found in ZIP")
                return None, None
            
            mastersheet_name = mastersheet_files[0]
            print(f"✅ Found mastersheet: {mastersheet_name}")
            
            temp_dir = tempfile.mkdtemp()
            temp_mastersheet_path = os.path.join(temp_dir, f"mastersheet_{semester_key}.xlsx")
            
            with open(temp_mastersheet_path, 'wb') as f:
                f.write(zip_ref.read(mastersheet_name))
            
            print(f"✅ Extracted mastersheet to: {temp_mastersheet_path}")
            return temp_mastersheet_path, temp_dir
            
    except Exception as e:
        print(f"❌ Error extracting mastersheet from ZIP: {e}")
        traceback.print_exc()
        return None, None

def find_latest_zip_file(clean_dir):
    """Find the latest ZIP file in clean results directory."""
    print(f"🔍 Looking for ND ZIP files in: {clean_dir}")
    
    if not os.path.exists(clean_dir):
        print(f"❌ ND clean directory doesn't exist: {clean_dir}")
        return None
    
    all_files = os.listdir(clean_dir)
    
    zip_files = []
    for f in all_files:
        if f.lower().endswith('.zip'):
            if 'carryover' in f.lower():
                print(f"⚠️ Skipping ND carryover ZIP: {f}")
                continue
            
            if any(pattern in f for pattern in ['_RESULT-', 'RESULT_', 'RESULT-']):
                zip_files.append(f)
                print(f"✅ Found ND regular results ZIP: {f}")
            else:
                print(f"ℹ️ Found other ND ZIP (not a result file): {f}")
    
    if not zip_files:
        print(f"❌ No ND regular results ZIP files found (excluding carryover files)")
        fallback_zips = [f for f in all_files if f.lower().endswith('.zip') and 'carryover' not in f.lower()]
        if fallback_zips:
            print(f"⚠️ Using fallback ND ZIP files: {fallback_zips}")
            zip_files = fallback_zips
        else:
            print(f"❌ No ND ZIP files found at all in {clean_dir}")
            return None
    
    print(f"✅ Final ND ZIP files to consider: {zip_files}")
    
    zip_files_with_path = [os.path.join(clean_dir, f) for f in zip_files]
    latest_zip = sorted(zip_files_with_path, key=os.path.getmtime, reverse=True)[0]
    
    print(f"🎯 Using latest ND ZIP: {latest_zip}")
    return latest_zip

def find_latest_result_folder(clean_dir, set_name):
    """Find the latest result folder in clean results directory."""
    print(f"🔍 Looking for ND result folders in: {clean_dir}")
    
    if not os.path.exists(clean_dir):
        print(f"❌ ND clean directory doesn't exist: {clean_dir}")
        return None
    
    all_items = os.listdir(clean_dir)
    
    result_folders = [f for f in all_items if os.path.isdir(os.path.join(clean_dir, f)) and f.startswith(f"{set_name}_RESULT-")]
    
    if not result_folders:
        print(f"❌ No ND result folders found")
        return None
    
    print(f"✅ Found ND result folders: {result_folders}")
    
    folders_with_path = [os.path.join(clean_dir, f) for f in result_folders]
    latest_folder = sorted(folders_with_path, key=os.path.getmtime, reverse=True)[0]
    
    print(f"🎯 Using latest ND result folder: {latest_folder}")
    return latest_folder

def find_latest_mastersheet_source(clean_dir, set_name):
    """Find the latest source for mastersheet: prefer ZIP, fallback to folder."""
    print(f"🔍 Looking for ND mastersheet source in: {clean_dir}")
    
    if not os.path.exists(clean_dir):
        print(f"❌ ND clean directory doesn't exist: {clean_dir}")
        return None, None
    
    zip_path = find_latest_zip_file(clean_dir)
    if zip_path:
        print(f"✅ Using ND ZIP source: {zip_path}")
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_files = zip_ref.namelist()
                mastersheet_files = [f for f in zip_files if 'mastersheet' in f.lower() and f.endswith('.xlsx')]
                if mastersheet_files:
                    print(f"✅ ND ZIP contains mastersheet files: {mastersheet_files}")
                    return zip_path, 'zip'
                else:
                    print(f"⚠️ ND ZIP found but no mastersheet inside: {zip_path}")
        except Exception as e:
            print(f"⚠️ Error checking ND ZIP contents: {e}")
    
    folder_path = find_latest_result_folder(clean_dir, set_name)
    if folder_path:
        print(f"✅ Using ND folder source: {folder_path}")
        return folder_path, 'folder'
    
    print(f"❌ No valid ND ZIP files or result folders found in {clean_dir}")
    return None, None

def get_mastersheet_path(source_path, source_type, semester_key):
    """Get mastersheet path based on source type (zip or folder)."""
    temp_dir = None
    if source_type == 'zip':
        temp_mastersheet_path, temp_dir = extract_mastersheet_from_zip(source_path, semester_key)
        if not temp_mastersheet_path:
            print("❌ Failed to extract mastersheet from ZIP")
            return None, None
    elif source_type == 'folder':
        all_files = os.listdir(source_path)
        mastersheet_files = [f for f in all_files if 'mastersheet' in f.lower() and f.endswith('.xlsx')]
        
        if not mastersheet_files:
            print(f"❌ No mastersheet found in folder {source_path}")
            return None, None
        
        mastersheet_name = mastersheet_files[0]
        temp_mastersheet_path = os.path.join(source_path, mastersheet_name)
        print(f"✅ Found mastersheet in folder: {temp_mastersheet_path}")
    else:
        return None, None
    
    return temp_mastersheet_path, temp_dir

def get_matching_sheet(xl, target_key):
    """Find matching sheet name with variants."""
    target_upper = target_key.upper().replace('-', ' ').replace('_', ' ').replace('.', ' ')
    target_upper = ' '.join(target_upper.split())
    
    possible_keys = [
        target_key,
        target_key.upper(),
        target_key.lower(),
        target_key.title(),
        target_key.replace('-', ' ').upper(),
        target_key.replace('-', ' ').lower(),
        target_key.replace('-', ' ').title(),
        target_key.replace('-', '_').upper(),
        target_key.replace('-', '_').lower(),
        target_key.replace('-', '_').title(),
        target_key.replace('First', '1st'),
        target_key.replace('Second', '2nd'),
        target_key.replace('Third', '3rd'),
        target_key.replace('YEAR', 'YR'),
        target_key.replace('SEMESTER', 'SEM'),
        target_upper,
        target_upper.replace('FIRST', '1ST'),
        target_upper.replace('SECOND', '2ND'),
        target_upper.replace('THIRD', '3RD'),
        target_upper.replace('YEAR', 'YR'),
        target_upper.replace('SEMESTER', 'SEM'),
    ]
    
    possible_keys = list(set([k for k in possible_keys if k]))
    
    print(f"🔍 Trying sheet variants for '{target_key}': {possible_keys}")
    
    for sheet in xl.sheet_names:
        sheet_normalized = sheet.upper().replace('-', ' ').replace('_', ' ').replace('.', ' ')
        sheet_normalized = ' '.join(sheet_normalized.split())
        
        if any(p == sheet or p in sheet or p == sheet_normalized or p in sheet_normalized for p in possible_keys):
            print(f"✅ Found matching sheet: '{sheet}' for '{target_key}'")
            return sheet
    
    print(f"❌ No matching sheet found for '{target_key}'")
    print(f"📖 Available sheets: {xl.sheet_names}")
    return None

def load_previous_gpas(mastersheet_path, current_semester_key):
    """Load previous GPA data from mastersheet for ND CGPA calculation."""
    all_student_data = {}
    current_standard = standardize_semester_key(current_semester_key)
    
    # ND semesters ONLY
    all_semesters = {
        "ND-FIRST-YEAR-FIRST-SEMESTER": [],
        "ND-FIRST-YEAR-SECOND-SEMESTER": ["ND-FIRST-YEAR-FIRST-SEMESTER"],
        "ND-SECOND-YEAR-FIRST-SEMESTER": ["ND-FIRST-YEAR-FIRST-SEMESTER", "ND-FIRST-YEAR-SECOND-SEMESTER"],
        "ND-SECOND-YEAR-SECOND-SEMESTER": ["ND-FIRST-YEAR-FIRST-SEMESTER", "ND-FIRST-YEAR-SECOND-SEMESTER", "ND-SECOND-YEAR-FIRST-SEMESTER"]
    }
    
    semesters_to_load = all_semesters.get(current_standard, [])
    print(f"📊 Loading previous ND GPAs for {current_standard}: {semesters_to_load}")

    if not os.path.exists(mastersheet_path):
        print(f"❌ ND Mastersheet not found: {mastersheet_path}")
        return {}

    try:
        xl = pd.ExcelFile(mastersheet_path)
        print(f"📖 Available sheets in ND mastersheet: {xl.sheet_names}")
    except Exception as e:
        print(f"❌ Error opening ND mastersheet: {e}")
        return {}

    for semester in semesters_to_load:
        try:
            sheet_name = get_matching_sheet(xl, semester)
            if not sheet_name:
                print(f"⚠️ Skipping ND semester {semester} - no matching sheet found")
                continue
            
            print(f"📖 Reading ND sheet '{sheet_name}' for semester {semester}")
            df = pd.read_excel(mastersheet_path, sheet_name=sheet_name, header=5)
            
            if df.empty or len(df.columns) < 3:
                df = pd.read_excel(mastersheet_path, sheet_name=sheet_name, header=0)
                print(f"🔄 Using header row 0 for ND sheet '{sheet_name}'")
            
            exam_col = find_exam_number_column(df)
            gpa_col = None
            credit_col = None
            
            for col in df.columns:
                col_str = str(col).upper()
                if 'GPA' in col_str and 'CGPA' not in col_str:
                    gpa_col = col
                if 'CU PASSED' in col_str or 'CREDIT' in col_str or 'UNIT' in col_str:
                    credit_col = col
            
            print(f"🔍 Columns found - Exam: {exam_col}, GPA: {gpa_col}, Credits: {credit_col}")
            
            if exam_col and gpa_col:
                for idx, row in df.iterrows():
                    try:
                        exam_no = str(row[exam_col]).strip()
                        if pd.isna(exam_no) or exam_no in ['', 'NAN', 'NONE']:
                            continue
                            
                        gpa_value = row[gpa_col]
                        if pd.isna(gpa_value):
                            continue
                            
                        credits = 30
                        if credit_col and credit_col in row and pd.notna(row[credit_col]):
                            try:
                                credits = int(float(row[credit_col]))
                            except (ValueError, TypeError):
                                credits = 30
                        
                        if exam_no not in all_student_data:
                            all_student_data[exam_no] = {'gpas': [], 'credits': [], 'semesters': []}
                        
                        all_student_data[exam_no]['gpas'].append(float(gpa_value))
                        all_student_data[exam_no]['credits'].append(credits)
                        all_student_data[exam_no]['semesters'].append(semester)
                        
                        if idx < 3:
                            print(f"📊 Loaded ND GPA for {exam_no}: {gpa_value} with {credits} credits from {semester}")
                            
                    except (ValueError, TypeError) as e:
                        print(f"⚠️ Error processing row {idx} for ND {semester}: {e}")
                        continue
            else:
                print(f"⚠️ Missing required columns in ND {sheet_name}: exam_col={exam_col}, gpa_col={gpa_col}")
                
        except Exception as e:
            print(f"⚠️ Could not load data from ND {semester}: {e}")
            traceback.print_exc()
    
    print(f"📊 Loaded cumulative ND data for {len(all_student_data)} students")
    return all_student_data

def calculate_cgpa(student_data, current_gpa, current_credits):
    """Calculate Cumulative GPA for ND."""
    if not student_data or not student_data.get('gpas'):
        print(f"⚠️ No previous ND GPA data, using current GPA: {current_gpa}")
        return current_gpa

    total_grade_points = 0.0
    total_credits = 0

    print(f"🔢 Calculating ND CGPA from {len(student_data['gpas'])} previous semesters")
    
    for prev_gpa, prev_credits in zip(student_data['gpas'], student_data['credits']):
        total_grade_points += prev_gpa * prev_credits
        total_credits += prev_credits
        print(f"   - GPA: {prev_gpa}, Credits: {prev_credits}, Running Total: {total_grade_points}/{total_credits}")

    total_grade_points += current_gpa * current_credits
    total_credits += current_credits

    print(f"📊 Final ND calculation: {total_grade_points} / {total_credits}")

    if total_credits > 0:
        cgpa = round(total_grade_points / total_credits, 2)
        print(f"✅ Calculated ND CGPA: {cgpa}")
        return cgpa
    else:
        print(f"⚠️ No ND credits, returning current GPA: {current_gpa}")
        return current_gpa

def get_previous_semesters_for_display(current_semester_key):
    """Get list of previous semesters for ND GPA display in mastersheet."""
    current_standard = standardize_semester_key(current_semester_key)
    
    # ND ONLY
    semester_mapping = {
        "ND-FIRST-YEAR-FIRST-SEMESTER": [],
        "ND-FIRST-YEAR-SECOND-SEMESTER": ["Semester 1"],
        "ND-SECOND-YEAR-FIRST-SEMESTER": ["Semester 1", "Semester 2"], 
        "ND-SECOND-YEAR-SECOND-SEMESTER": ["Semester 1", "Semester 2", "Semester 3"]
    }
    
    return semester_mapping.get(current_standard, [])

def extract_semester_from_filename(filename):
    """Extract semester from filename - ND ONLY."""
    filename_upper = filename.upper()
    
    semester_pattern = r'(ND[-_]?(?:FIRST|SECOND|1ST|2ND)[-_]?YEAR[-_]?(?:FIRST|SECOND|1ST|2ND)[-_]?SEMESTER)'
    match = re.search(semester_pattern, filename_upper)
    
    if match:
        extracted = match.group(1)
        standardized = standardize_semester_key(extracted)
        print(f"✅ Extracted and standardized: '{filename}' → '{standardized}'")
        return standardized
    
    semester_patterns = {
        "ND-FIRST-YEAR-FIRST-SEMESTER": [
            "FIRST.YEAR.FIRST.SEMESTER", "FIRST-YEAR-FIRST-SEMESTER",
        ],
        "ND-FIRST-YEAR-SECOND-SEMESTER": [
            "FIRST.YEAR.SECOND.SEMESTER", "FIRST-YEAR-SECOND-SEMESTER",
        ],
        "ND-SECOND-YEAR-FIRST-SEMESTER": [
            "SECOND.YEAR.FIRST.SEMESTER", "SECOND-YEAR-FIRST-SEMESTER",
        ],
        "ND-SECOND-YEAR-SECOND-SEMESTER": [
            "SECOND.YEAR.SECOND.SEMESTER", "SECOND-YEAR-SECOND-SEMESTER",
        ]
    }
    
    for semester_key, patterns in semester_patterns.items():
        for pattern in patterns:
            flexible_pattern = pattern.replace('.', '[._\\- ]?')
            if re.search(flexible_pattern, filename_upper, re.IGNORECASE):
                print(f"✅ Matched semester '{semester_key}' for filename: {filename}")
                return semester_key
    
    print(f"❌ Could not determine semester for filename: {filename}")
    return "UNKNOWN_SEMESTER"

def load_carryover_json_files(carryover_dir, semester_key=None):
    """Load carryover JSON files from directory."""
    carryover_files = []
    
    if semester_key:
        semester_key = standardize_semester_key(semester_key)
    
    previous_semester = get_previous_semester(semester_key)
    print(f"🔑 Target ND semester: {semester_key}")
    print(f"🔑 Previous ND semester for carryover: {previous_semester}")
    
    for file in os.listdir(carryover_dir):
        if file.startswith("co_student_") and file.endswith(".json"):
            file_semester = extract_semester_from_filename(file)
            file_semester_standardized = standardize_semester_key(file_semester)
            
            print(f"📄 Found ND carryover file: {file}")
            print(f"   Original semester: {file_semester}")
            print(f"   Standardized: {file_semester_standardized}")
            print(f"   Target previous: {previous_semester}")
            
            if previous_semester and file_semester_standardized != previous_semester:
                print(f"   ⏭️ Skipping (doesn't match previous ND semester)")
                continue
            
            file_path = os.path.join(carryover_dir, file)
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f)
                    carryover_files.append({
                        'filename': file,
                        'semester': file_semester_standardized,
                        'data': data,
                        'count': len(data),
                        'file_path': file_path
                    })
                    print(f"   ✅ Loaded: {len(data)} ND records")
            except Exception as e:
                print(f"Error loading ND {file}: {e}")
    
    print(f"📊 Total ND carryover files loaded: {len(carryover_files)}")
    return carryover_files

def get_carryover_records_from_zip(zip_path, set_name, semester_key):
    """Get carryover records from ZIP file."""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            temp_dir = tempfile.mkdtemp()
            carryover_dir = os.path.join(temp_dir, "CARRYOVER_RECORDS")
            os.makedirs(carryover_dir, exist_ok=True)
            for member in zip_ref.namelist():
                if member.startswith("CARRYOVER_RECORDS/"):
                    zip_ref.extract(member, temp_dir)
            records = load_carryover_json_files(carryover_dir, semester_key)
            shutil.rmtree(temp_dir)
            print(f"✅ Loaded {len(records)} ND carryover records from ZIP")
            return records
    except Exception as e:
        print(f"❌ Error loading from ND ZIP: {e}")
        return []

def get_carryover_records(set_name, semester_key=None):
    """Get carryover records for ND."""
    try:
        if semester_key:
            semester_key = standardize_semester_key(semester_key)
            print(f"🔑 Using standardized ND semester key: {semester_key}")
        
        clean_dir = os.path.join(BASE_DIR, "ND", set_name, "CLEAN_RESULTS")
        if not os.path.exists(clean_dir):
            print(f"❌ ND clean directory not found: {clean_dir}")
            return []
        
        timestamp_items = []
        
        for item in os.listdir(clean_dir):
            item_path = os.path.join(clean_dir, item)
            
            if item.startswith(f"{set_name}_RESULT-") and "CARRYOVER" not in item.upper():
                if os.path.isdir(item_path) or item.endswith('.zip'):
                    timestamp_items.append(item)
                    print(f"Found ND regular result: {item}")
        
        if not timestamp_items:
            print(f"❌ No ND regular result files found in: {clean_dir}")
            return []
        
        latest_item = sorted(timestamp_items)[-1]
        latest_path = os.path.join(clean_dir, latest_item)
        print(f"✅ Using latest ND result: {latest_item}")
        
        if latest_item.endswith('.zip'):
            return get_carryover_records_from_zip(latest_path, set_name, semester_key)
        else:
            carryover_dir = os.path.join(latest_path, "CARRYOVER_RECORDS")
            if not os.path.exists(carryover_dir):
                print(f"❌ No ND CARRYOVER_RECORDS folder in: {latest_path}")
                return []
            return load_carryover_json_files(carryover_dir, semester_key)
            
    except Exception as e:
        print(f"Error getting ND carryover records: {e}")
        return []

def process_carryover_results(resit_file_path, source_path, source_type, semester_key, set_name, pass_threshold, output_dir):
    """Process ND carryover results ONLY."""
    print(f"\n🔄 PROCESSING ND CARRYOVER RESULTS FOR {semester_key}")
    print("=" * 60)
    
    semester_course_titles, semester_credit_units, course_code_to_title, course_code_to_unit = load_course_data()
    
    debug_course_matching(resit_file_path, course_code_to_title, course_code_to_unit)
    
    year, sem_num, level, sem_display, set_code, sem_name = get_semester_display_info(semester_key)
    
    possible_sheet_keys = [
        f"{set_code} {sem_display}",
        f"{set_code.replace('NDII', 'ND II').replace('NDI', 'ND I')} {sem_display}",
        semester_key,
        semester_key.replace('-', ' ').upper(),
        f"{set_code} {sem_name}",
        f"{level} {sem_display}",
    ]
    
    course_titles_dict = {}
    credit_units_dict = {}
    
    for sheet_key in possible_sheet_keys:
        sheet_standard = standardize_semester_key(sheet_key)
        if sheet_standard in semester_course_titles:
            course_titles_dict = semester_course_titles[sheet_standard]
            credit_units_dict = semester_credit_units[sheet_standard]
            print(f"✅ Using ND sheet key: '{sheet_key}' with {len(course_titles_dict)} courses")
            break
        else:
            print(f"❌ ND sheet key not found: '{sheet_key}'")
    
    if not course_titles_dict:
        print(f"⚠️ No ND semester-specific course data found, using global course mappings")
        course_titles_dict = course_code_to_title
        credit_units_dict = course_code_to_unit
    
    print(f"📊 Final ND course mappings: {len(course_titles_dict)} titles, {len(credit_units_dict)} units")
    
    timestamp = datetime.now().strftime(TIMESTAMP_FMT)
    carryover_output_dir = os.path.join(output_dir, f"CARRYOVER_{set_name}_{semester_key}_{timestamp}")
    os.makedirs(carryover_output_dir, exist_ok=True)
    
    if not os.path.exists(resit_file_path):
        print(f"❌ ND resit file not found: {resit_file_path}")
        return False
    
    temp_mastersheet_path = None
    temp_dir = None
    
    try:
        temp_mastersheet_path, temp_dir = get_mastersheet_path(source_path, source_type, semester_key)
        
        if not temp_mastersheet_path:
            print(f"❌ Failed to get ND mastersheet")
            return False
        
        print(f"📖 Reading ND files...")
        resit_df = pd.read_excel(resit_file_path, header=0)
        
        # DEBUG: Print resit file info
        print(f"📊 Resit file rows: {len(resit_df)}")
        print(f"📊 Resit file columns: {resit_df.columns.tolist()}")
        resit_exam_col = find_exam_number_column(resit_df)
        print(f"📊 Resit exam column: '{resit_exam_col}'")
        if resit_exam_col:
            print(f"📊 Sample resit exam numbers: {resit_df[resit_exam_col].head().tolist()}")
        
        xl = pd.ExcelFile(temp_mastersheet_path)
        sheet_name = get_matching_sheet(xl, semester_key)
        if not sheet_name:
            print(f"❌ No matching ND sheet found for {semester_key}")
            return False
        
        print(f"📖 Using ND sheet '{sheet_name}' for current semester {semester_key}")
        
        try:
            mastersheet_df = pd.read_excel(temp_mastersheet_path, sheet_name=sheet_name, header=5)
        except:
            try:
                mastersheet_df = pd.read_excel(temp_mastersheet_path, sheet_name=sheet_name, header=0)
                print(f"⚠️ Using header row 0 for ND mastersheet")
            except Exception as e:
                print(f"❌ Error reading ND mastersheet: {e}")
                return False
        
        # DEBUG: Print mastersheet info
        print(f"📊 Mastersheet rows: {len(mastersheet_df)}")
        mastersheet_exam_col = find_exam_number_column(mastersheet_df) or 'EXAM NUMBER'
        print(f"📊 Mastersheet exam column: '{mastersheet_exam_col}'")
        if mastersheet_exam_col in mastersheet_df.columns:
            print(f"📊 Sample mastersheet exam numbers: {mastersheet_df[mastersheet_exam_col].head().tolist()}")
        
        print(f"✅ ND files loaded - Resit: {len(resit_df)} rows, Mastersheet: {len(mastersheet_df)} students")
        
        resit_exam_col = find_exam_number_column(resit_df)
        mastersheet_exam_col = find_exam_number_column(mastersheet_df) or 'EXAM NUMBER'
        
        if not resit_exam_col:
            print(f"❌ Cannot find exam number column in ND resit file")
            return False
        
        print(f"📝 ND Exam columns - Resit: '{resit_exam_col}', Mastersheet: '{mastersheet_exam_col}'")
        
        cgpa_data = load_previous_gpas(temp_mastersheet_path, semester_key)
        
        carryover_data = []
        updated_students = set()
        
        print(f"\n🎯 PROCESSING ND RESIT SCORES...")
        
        for idx, resit_row in resit_df.iterrows():
            exam_no = str(resit_row[resit_exam_col]).strip().upper()
            if not exam_no or exam_no in ['NAN', 'NONE', '']:
                continue
            
            student_mask = mastersheet_df[mastersheet_exam_col].astype(str).str.strip().str.upper() == exam_no
            
            # DEBUG: Print student processing info
            print(f"\n🔍 Processing student: {exam_no}")
            print(f"   Found in mastersheet: {student_mask.any()}")
            
            if not student_mask.any():
                print(f"⚠️ ND Student {exam_no} not found in mastersheet - skipping")
                continue
            
            student_data = mastersheet_df[student_mask].iloc[0]
            student_name = student_data.get('NAME', 'Unknown')
            
            if student_mask.any():
                print(f"   Student name: {student_name}")
                print(f"   Resit courses to check: {[col for col in resit_df.columns if col not in [resit_exam_col, 'NAME']]}")
            
            current_credits = 0
            for col in mastersheet_df.columns:
                if 'CU PASSED' in str(col).upper():
                    current_credits = student_data.get(col, 0)
                    break
            
            student_record = {
                'EXAM NUMBER': exam_no,
                'NAME': student_name,
                'RESIT_COURSES': {},
                'CURRENT_GPA': student_data.get('GPA', 0),
                'CURRENT_CREDITS': current_credits
            }
            
            if exam_no in cgpa_data:
                student_record['CURRENT_CGPA'] = calculate_cgpa(
                    cgpa_data[exam_no], 
                    student_record['CURRENT_GPA'], 
                    current_credits
                )
            else:
                student_record['CURRENT_CGPA'] = student_record['CURRENT_GPA']
            
            if exam_no in cgpa_data:
                student_gpa_data = cgpa_data[exam_no]
                for i, prev_semester in enumerate(student_gpa_data['semesters']):
                    sem_display_name = get_semester_display_info(prev_semester)[5]
                    student_record[f'GPA_{sem_display_name}'] = student_gpa_data['gpas'][i]
                    print(f"📊 Stored ND GPA for {exam_no}: {sem_display_name} = {student_gpa_data['gpas'][i]}")
            
            for col in resit_df.columns:
                if col == resit_exam_col or col == 'NAME' or 'Unnamed' in str(col):
                    continue
                    
                resit_score = resit_row.get(col)
                if pd.isna(resit_score) or resit_score == '':
                    continue
                
                try:
                    resit_score_val = float(resit_score)
                except (ValueError, TypeError):
                    continue
                
                if col in mastersheet_df.columns:
                    original_score = student_data.get(col)
                    if pd.isna(original_score):
                        continue
                    
                    try:
                        original_score_val = float(original_score) if not pd.isna(original_score) else 0.0
                    except (ValueError, TypeError):
                        original_score_val = 0.0
                    
                    if original_score_val < pass_threshold:
                        course_title = find_course_title(col, course_titles_dict, course_code_to_title)
                        credit_unit = find_credit_unit(col, credit_units_dict, course_code_to_unit)
                        
                        student_record['RESIT_COURSES'][col] = {
                            'original_score': original_score_val,
                            'resit_score': resit_score_val,
                            'updated': resit_score_val >= pass_threshold,
                            'course_title': course_title,
                            'credit_unit': credit_unit
                        }
            
            if student_record['RESIT_COURSES']:
                carryover_data.append(student_record)
                updated_students.add(exam_no)
                print(f"✅ ND {exam_no}: {len(student_record['RESIT_COURSES'])} resit courses, CGPA: {student_record['CURRENT_CGPA']}")
        
        # DEBUG: Print final stats
        print(f"\n📊 FINAL STATS:")
        print(f"   Total carryover students processed: {len(carryover_data)}")
        print(f"   Students with updates: {len(updated_students)}")
        
        if carryover_data:
            print(f"\n📊 GENERATING ND CARRYOVER MASTERSHEET...")
            carryover_mastersheet_path = generate_carryover_mastersheet(
                carryover_data, carryover_output_dir, semester_key, set_name, timestamp, 
                cgpa_data, course_titles_dict, credit_units_dict, course_code_to_title, course_code_to_unit
            )
            
            print(f"\n📄 GENERATING ND INDIVIDUAL STUDENT REPORTS...")
            generate_individual_reports(
                carryover_data, carryover_output_dir, semester_key, set_name, timestamp, cgpa_data
            )
            
            zip_path = os.path.join(output_dir, f"CARRYOVER_{set_name}_{semester_key}_{timestamp}.zip")
            if create_carryover_zip(carryover_output_dir, zip_path):
                print(f"✅ Final ND carryover ZIP created: {zip_path}")
            
            print(f"\n🎉 ND CARRYOVER PROCESSING COMPLETED!")
            print(f"📁 Output directory: {carryover_output_dir}")
            print(f"📦 ZIP file: {zip_path}")
            print(f"👨‍🎓 ND Students processed: {len(carryover_data)}")
            
            return True
        else:
            print(f"❌ No ND carryover data found to process")
            return False
            
    except Exception as e:
        print(f"❌ Error processing ND carryover results: {e}")
        traceback.print_exc()
        return False
    finally:
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            print(f"✅ Cleaned up ND temporary files")

def generate_remarks(resit_courses):
    """Generate remarks for ND resit performance."""
    passed_count = sum(1 for course_data in resit_courses.values() 
                      if course_data['resit_score'] >= DEFAULT_PASS_THRESHOLD)
    total_count = len(resit_courses)
    
    if passed_count == total_count:
        return "All courses passed in resit"
    elif passed_count > 0:
        return f"{passed_count}/{total_count} courses passed in resit"
    else:
        return "No improvement in resit"

def generate_carryover_mastersheet(carryover_data, output_dir, semester_key, set_name, timestamp, cgpa_data, course_titles, course_units, course_code_to_title, course_code_to_unit):
    """Generate ND CARRYOVER_mastersheet."""
    
    wb = Workbook()
    ws = wb.active
    
    ws.title = "ND_CARRYOVER_RESULTS"
    program_name = "NATIONAL DIPLOMA"
    program_abbr = "ND"
    
    if os.path.exists(DEFAULT_LOGO_PATH):
        try:
            from openpyxl.drawing.image import Image
            img = Image(DEFAULT_LOGO_PATH)
            img.width = 80
            img.height = 80
            ws.add_image(img, 'A1')
        except Exception as e:
            print(f"⚠️ Could not add logo: {e}")
    
    current_year = 2025
    next_year = 2026
    year, sem_num, level, sem_display, set_code, current_semester_name = get_semester_display_info(semester_key)
    
    all_courses = set()
    for student in carryover_data:
        all_courses.update(student['RESIT_COURSES'].keys())
    
    previous_semesters = get_previous_semesters_for_display(semester_key)
    
    headers = ['S/N', 'EXAM NUMBER', 'NAME']
    
    for prev_sem in previous_semesters:
        headers.append(f'GPA {prev_sem}')
    
    course_headers = []
    for course in sorted(all_courses):
        course_headers.extend([f'{course}', f'{course}_RESIT'])
    
    headers.extend(course_headers)
    headers.extend([f'GPA {current_semester_name}', 'CGPA', 'REMARKS'])
    
    total_columns = len(headers)
    last_column = get_column_letter(total_columns)
    
    ws.merge_cells(f'A3:{last_column}3')
    title_cell = ws['A3']
    title_cell.value = "FCT COLLEGE OF NURSING SCIENCES, GWAGWALADA-ABUJA"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    ws.merge_cells(f'A4:{last_column}4')
    subtitle_cell = ws['A4']
    
    subtitle_cell.value = f"RESIT - {current_year}/{next_year} SESSION {program_name} {level} {sem_display} EXAMINATIONS RESULT — {datetime.now().strftime('%B %d, %Y')}"
    
    subtitle_cell.font = Font(bold=True, size=12)
    subtitle_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    print(f"🔍 ND Courses found in resit data: {sorted(all_courses)}")
    print(f"📊 ND GPA columns for {semester_key}: Previous={previous_semesters}, Current={current_semester_name}")
    
    headers = ['S/N', 'EXAM NUMBER', 'NAME']
    
    for prev_sem in previous_semesters:
        headers.append(f'GPA {prev_sem}')
    
    course_headers = []
    course_title_mapping = {}
    course_unit_mapping = {}
    
    for course in sorted(all_courses):
        course_title = find_course_title(course, course_titles, course_code_to_title)
        course_title_mapping[course] = course_title
        
        credit_unit = find_credit_unit(course, course_units, course_code_to_unit)
        course_unit_mapping[course] = credit_unit
        
        if len(course_title) > 30:
            course_title = course_title[:27] + "..."
        course_headers.extend([f'{course}', f'{course}_RESIT'])
    
    headers.extend(course_headers)
    headers.extend([f'GPA {current_semester_name}', 'CGPA', 'REMARKS'])
    
    title_row = [''] * 3
    
    for prev_sem in previous_semesters:
        title_row.extend([''])
    
    for course in sorted(all_courses):
        course_title = course_title_mapping[course]
        if len(course_title) > 30:
            course_title = course_title[:27] + "..."
        title_row.extend([course_title, course_title])
    
    title_row.extend(['', '', ''])
    
    ws.append(title_row)
    
    credit_row = [''] * 3
    
    for prev_sem in previous_semesters:
        credit_row.extend([''])
    
    for course in sorted(all_courses):
        credit_unit = course_unit_mapping[course]
        credit_row.extend([f'CU: {credit_unit}', f'CU: {credit_unit}'])
    
    credit_row.extend(['', '', ''])
    
    ws.append(credit_row)
    
    code_row = ['S/N', 'EXAM NUMBER', 'NAME']
    
    for prev_sem in previous_semesters:
        code_row.append(f'GPA {prev_sem}')
    
    for course in sorted(all_courses):
        code_row.extend([f'{course}', f'{course}_RESIT'])
    
    code_row.extend([f'GPA {current_semester_name}', 'CGPA', 'REMARKS'])
    
    ws.append(code_row)
    
    course_colors = [
        "E6F3FF", "FFF0E6", "E6FFE6", "FFF6E6", "F0E6FF",
        "E6FFFF", "FFE6F2", "F5F5DC", "E6F7FF", "FFF5E6",
    ]
    
    start_col = 4
    if previous_semesters:
        start_col += len(previous_semesters)
    
    color_index = 0
    for course in sorted(all_courses):
        for row in [5, 6, 7]:
            for offset in [0, 1]:
                cell = ws.cell(row=row, column=start_col + offset)
                cell.fill = PatternFill(start_color=course_colors[color_index % len(course_colors)], 
                                      end_color=course_colors[color_index % len(course_colors)], 
                                      fill_type="solid")
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
        
        for offset in [0, 1]:
            cell = ws.cell(row=5, column=start_col + offset)
            cell.alignment = Alignment(text_rotation=90, horizontal='center', vertical='center')
            cell.font = Font(bold=True, size=9)
        
        color_index += 1
        start_col += 2
    
    for row in [5, 6, 7]:
        for col in range(1, 4):
            cell = ws.cell(row=row, column=col)
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        
        gpa_col = 4
        for prev_sem in previous_semesters:
            cell = ws.cell(row=row, column=gpa_col)
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            gpa_col += 1
        
        for col in range(len(headers)-2, len(headers)+1):
            cell = ws.cell(row=row, column=col)
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
    
    row_idx = 8
    failed_counts = {course: 0 for course in all_courses}
    
    start_col = 4
    if previous_semesters:
        start_col += len(previous_semesters)
    
    for student in carryover_data:
        exam_no = student['EXAM NUMBER']
        
        ws.cell(row=row_idx, column=1, value=row_idx-7)
        ws.cell(row=row_idx, column=2, value=student['EXAM NUMBER'])
        ws.cell(row=row_idx, column=3, value=student['NAME'])
        
        gpa_col = 4
        for prev_sem in previous_semesters:
            gpa_value = student.get(f'GPA_{prev_sem}', '')
            ws.cell(row=row_idx, column=gpa_col, value=gpa_value)
            gpa_col += 1
        
        course_col = gpa_col
        color_index = 0
        for course in sorted(all_courses):
            for offset in [0, 1]:
                cell = ws.cell(row=row_idx, column=course_col + offset)
                cell.fill = PatternFill(start_color=course_colors[color_index % len(course_colors)], 
                                      end_color=course_colors[color_index % len(course_colors)], 
                                      fill_type="solid")
            
            if course in student['RESIT_COURSES']:
                course_data = student['RESIT_COURSES'][course]
                
                orig_cell = ws.cell(row=row_idx, column=course_col, value=course_data['original_score'])
                if course_data['original_score'] < DEFAULT_PASS_THRESHOLD:
                    orig_cell.fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
                
                resit_cell = ws.cell(row=row_idx, column=course_col+1, value=course_data['resit_score'])
                if course_data['resit_score'] >= DEFAULT_PASS_THRESHOLD:
                    resit_cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
                else:
                    resit_cell.fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
                    failed_counts[course] += 1
            else:
                ws.cell(row=row_idx, column=course_col, value='')
                ws.cell(row=row_idx, column=course_col+1, value='')
            
            color_index += 1
            course_col += 2
        
        ws.cell(row=row_idx, column=course_col, value=student['CURRENT_GPA'])
        ws.cell(row=row_idx, column=course_col+1, value=student['CURRENT_CGPA'])
        
        remarks = generate_remarks(student['RESIT_COURSES'])
        ws.cell(row=row_idx, column=course_col+2, value=remarks)
        
        row_idx += 1
    
    failed_row_idx = row_idx
    ws.cell(row=failed_row_idx, column=1, value="FAILED COUNT BY COURSE:").font = Font(bold=True)
    
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=failed_row_idx, column=col)
        cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    
    course_col = gpa_col
    for course in sorted(all_courses):
        count_cell = ws.cell(row=failed_row_idx, column=course_col+1, value=failed_counts[course])
        count_cell.font = Font(bold=True)
        count_cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        course_col += 2
    
    summary_start_row = failed_row_idx + 2
    
    total_students = len(carryover_data)
    passed_all = sum(1 for student in carryover_data 
                    if all(course_data['resit_score'] >= DEFAULT_PASS_THRESHOLD 
                          for course_data in student['RESIT_COURSES'].values()))
    
    carryover_count = total_students - passed_all
    total_failed_attempts = sum(failed_counts.values())
    
    summary_data = [
        ["CARRYOVER SUMMARY"],
        [f"A total of {total_students} students registered and sat for the Carryover Examination"],
        [f"A total of {passed_all} students passed all carryover courses"],
        [f"A total of {carryover_count} students failed one or more carryover courses and must repeat them"],
        [f"Total failed resit attempts: {total_failed_attempts} across all courses"],
        [f"Carryover processing completed on {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}"],
        [""],
        [""],
        ["", ""],
        ["________________________", "________________________"],
        ["Mrs. Abini Hauwa", "Mrs. Olukemi Ogunleye"],
        ["Head of Exams", "Chairman, ND Program C'tee"]
    ]
    
    for i, row_data in enumerate(summary_data):
        row_num = summary_start_row + i
        if len(row_data) == 1:
            if row_data[0]:
                ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=10)
                cell = ws.cell(row=row_num, column=1, value=row_data[0])
                if i == 0:
                    cell.font = Font(bold=True, size=12, underline='single')
                else:
                    cell.font = Font(bold=False, size=11)
                cell.alignment = Alignment(horizontal='left', vertical='center')
        elif len(row_data) == 2:
            left_cell = ws.cell(row=row_num, column=1, value=row_data[0])
            right_cell = ws.cell(row=row_num, column=4, value=row_data[1])
            
            if i >= len(summary_data) - 3:
                left_cell.alignment = Alignment(horizontal='left')
                right_cell.alignment = Alignment(horizontal='left')
                left_cell.font = Font(bold=True, size=11)
                right_cell.font = Font(bold=True, size=11)
    
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    for row in ws.iter_rows(min_row=7, max_row=row_idx-1, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border
    
    ws.freeze_panes = 'D8'
    
    for row in ws.iter_rows():
        for cell in row:
            if cell.font is None or not cell.font.bold:
                cell.font = Font(name='Calibri', size=11)
    
    for col_idx, column in enumerate(ws.columns, 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        
        for cell in column:
            try:
                if cell.value is not None:
                    cell_value = str(cell.value)
                    cell_length = len(cell_value)
                    
                    if cell.row == 5 and cell.alignment.text_rotation == 90:
                        cell_length = max(cell_length, 10)
                    
                    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                        cell_length = max(cell_length, 8)
                    
                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass
        
        adjusted_width = min(max_length + 2, 50)
        
        if col_idx == 1:
            adjusted_width = 8
        elif col_idx == 2:
            adjusted_width = 18
        elif col_idx == 3:
            adjusted_width = 35
        elif col_idx >= 4 and col_idx <= (4 + len(previous_semesters) - 1):
            adjusted_width = 15
        elif col_idx >= len(headers) - 2:
            adjusted_width = 15
        else:
            adjusted_width = min(max(adjusted_width, 12), 25)
        
        ws.column_dimensions[column_letter].width = adjusted_width
    
    for row_idx in range(8, row_idx):
        if row_idx % 2 == 0:
            for cell in ws[row_idx]:
                if (cell.fill.start_color.index == '00000000' or 
                    cell.fill.start_color.index == '00FFFFFF'):
                    cell.fill = PatternFill(start_color="F8F8F8", end_color="F8F8F8", fill_type="solid")
    
    gpa_fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
    if previous_semesters:
        for row in range(8, row_idx):
            for col in range(4, 4 + len(previous_semesters)):
                cell = ws.cell(row=row, column=col)
                if cell.fill.start_color.index == '00000000':
                    cell.fill = gpa_fill
    
    final_gpa_fill = PatternFill(start_color="E0FFFF", end_color="E0FFFF", fill_type="solid")
    for row in range(8, row_idx):
        for col in range(len(headers)-2, len(headers)+1):
            cell = ws.cell(row=row, column=col)
            if cell.fill.start_color.index == '00000000':
                cell.fill = final_gpa_fill
    
    filename = f"CARRYOVER_mastersheet_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    
    print(f"✅ ND CARRYOVER mastersheet generated: {filepath}")
    return filepath

def generate_individual_reports(carryover_data, output_dir, semester_key, set_name, timestamp, cgpa_data):
    """Generate individual ND student reports."""
    reports_dir = os.path.join(output_dir, "INDIVIDUAL_REPORTS")
    os.makedirs(reports_dir, exist_ok=True)
    
    for student in carryover_data:
        exam_no = student['EXAM NUMBER']
        safe_exam_no = sanitize_filename(exam_no)
        
        filename = f"carryover_report_{safe_exam_no}_{timestamp}.csv"
        filepath = os.path.join(reports_dir, filename)
        
        report_data = []
        report_data.append(["ND CARRYOVER RESULT REPORT"])
        report_data.append(["FCT COLLEGE OF NURSING SCIENCES"])
        report_data.append([f"ND Set: {set_name}"])
        report_data.append([f"ND Semester: {semester_key}"])
        report_data.append([])
        report_data.append(["ND STUDENT INFORMATION"])
        report_data.append(["Exam Number:", student['EXAM NUMBER']])
        report_data.append(["Name:", student['NAME']])
        report_data.append([])
        
        report_data.append(["ND PREVIOUS GPAs"])
        for key in sorted([k for k in student.keys() if k.startswith('GPA_')]):
            semester = key.replace('GPA_', '')
            report_data.append([f"{semester}:", student[key]])
        report_data.append([])
        
        report_data.append(["ND CURRENT ACADEMIC RECORD"])
        report_data.append(["Current GPA:", student['CURRENT_GPA']])
        report_data.append(["Current CGPA:", student['CURRENT_CGPA']])
        report_data.append([])
        
        report_data.append(["ND RESIT COURSES"])
        report_data.append(["Course Code", "Course Title", "Credit Unit", "Original Score", "Resit Score", "Status"])
        
        for course_code, course_data in student['RESIT_COURSES'].items():
            status = "PASSED" if course_data['resit_score'] >= DEFAULT_PASS_THRESHOLD else "FAILED"
            course_title = course_data.get('course_title', course_code)
            credit_unit = course_data.get('credit_unit', 0)
            report_data.append([
                course_code, 
                course_title,
                credit_unit,
                course_data['original_score'], 
                course_data['resit_score'], 
                status
            ])
        
        try:
            df = pd.DataFrame(report_data)
            df.to_csv(filepath, index=False, header=False)
            print(f"✅ Generated ND report for: {exam_no}")
        except Exception as e:
            print(f"❌ Error generating ND report for {exam_no}: {e}")
    
    print(f"✅ Generated {len(carryover_data)} ND individual student reports in {reports_dir}")

def create_carryover_zip(source_dir, zip_path):
    """Create ZIP file of ND carryover results."""
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, source_dir)
                    zipf.write(file_path, arcname)
        print(f"✅ ND ZIP file created: {zip_path}")
        return True
    except Exception as e:
        print(f"❌ Error creating ND ZIP: {e}")
        return False

def main():
    """Main function to process ND carryover results."""
    print("=" * 60)
    print("🎯 ND CARRYOVER RESULT PROCESSOR")
    print("=" * 60)
    
    # Get environment variables
    set_name = os.getenv("SELECTED_SET", "")
    semester_key = os.getenv("SELECTED_SEMESTERS", "")
    resit_file_path = os.getenv("RESIT_FILE_PATH", "")
    base_result_path = os.getenv("BASE_RESULT_PATH", "")
    output_dir_env = os.getenv("OUTPUT_DIR", "")
    pass_threshold = float(os.getenv("PASS_THRESHOLD", str(DEFAULT_PASS_THRESHOLD)))
    
    # Validate inputs
    if not set_name:
        print("❌ ERROR: SELECTED_SET not provided")
        print("💡 Set the SELECTED_SET environment variable")
        return
    
    if not semester_key:
        print("❌ ERROR: SELECTED_SEMESTERS not provided")
        print("💡 Set the SELECTED_SEMESTERS environment variable")
        return
    
    if not resit_file_path or not os.path.exists(resit_file_path):
        print(f"❌ ERROR: RESIT_FILE_PATH not provided or doesn't exist: {resit_file_path}")
        print("💡 Set the RESIT_FILE_PATH environment variable to a valid resit file")
        return
    
    # Validate ND set
    if not set_name.startswith("ND-"):
        print(f"❌ ERROR: Invalid ND set name: {set_name}")
        print(f"💡 ND set names must start with 'ND-' (e.g., ND-2024, ND-2025)")
        return
    
    print(f"✅ Processing ND Set: {set_name}")
    print(f"✅ Processing Semester: {semester_key}")
    print(f"✅ Resit file: {resit_file_path}")
    
    # Find directories
    clean_dir = None
    output_dir = None
    
    possible_base_dirs = [
        BASE_DIR,
        os.path.join(BASE_DIR, "EXAMS_INTERNAL"),
        os.path.join(os.path.expanduser('~'), 'student_result_cleaner', 'EXAMS_INTERNAL'),
    ]
    
    for base in possible_base_dirs:
        test_clean_dir = os.path.join(base, "ND", set_name, "CLEAN_RESULTS")
        
        if os.path.exists(test_clean_dir):
            clean_dir = test_clean_dir
            output_dir = output_dir_env if output_dir_env else test_clean_dir
            print(f"✅ Found ND clean directory: {clean_dir}")
            break
    
    if not clean_dir:
        print(f"❌ ERROR: ND clean directory not found for {set_name}")
        print("💡 Please check:")
        print(f"   - Set name: {set_name}")
        print(f"   - Base directory: {BASE_DIR}")
        print(f"💡 Run the ND regular result processor first to generate clean results")
        return
    
    print(f"📁 Clean directory: {clean_dir}")
    print(f"📁 Output directory: {output_dir}")
    
    # Find mastersheet source
    if base_result_path and os.path.exists(base_result_path):
        print(f"✅ Using provided base result path: {base_result_path}")
        source_path = base_result_path
        source_type = 'zip' if base_result_path.endswith('.zip') else 'folder'
    else:
        print(f"🔍 Looking for mastersheet in: {clean_dir}")
        source_path, source_type = find_latest_mastersheet_source(clean_dir, set_name)
        
        if not source_path:
            print(f"❌ ERROR: No ZIP files or result folders found in {clean_dir}")
            print(f"💡 Run the ND regular result processor first to generate clean results")
            return
    
    print(f"✅ Using source: {source_path} (type: {source_type})")
    
    # Process carryover results
    success = process_carryover_results(
        resit_file_path=resit_file_path,
        source_path=source_path,
        source_type=source_type,
        semester_key=semester_key,
        set_name=set_name,
        pass_threshold=pass_threshold,
        output_dir=output_dir
    )
    
    if success:
        print("\n" + "=" * 60)
        print("✅ ND CARRYOVER PROCESSING COMPLETED")
        print("=" * 60)
        print(f"📁 Check the CLEAN_RESULTS directory for the CARRYOVER output")
        print(f"📦 Location: {output_dir}")
    else:
        print("\n" + "=" * 60)
        print("❌ ND CARRYOVER PROCESSING FAILED")
        print("=" * 60)
        print("💡 Check the error messages above for details")

if __name__ == "__main__":
    main()