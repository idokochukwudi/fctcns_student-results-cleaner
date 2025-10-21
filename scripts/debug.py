#!/usr/bin/env python3
"""
DEBUG VERSION - Semester Detection Fix
"""

import os
import sys
import re
import pandas as pd
from datetime import datetime

# ----------------------------
# Configuration
# ----------------------------
def get_base_directory():
    return os.path.join(os.path.expanduser('~'), 'student_result_cleaner', 'EXAMS_INTERNAL')

BASE_DIR = get_base_directory()
ND_BASE_DIR = os.path.join(BASE_DIR, "ND")

# Define semester processing order
SEMESTER_ORDER = [
    "ND-FIRST-YEAR-FIRST-SEMESTER",
    "ND-FIRST-YEAR-SECOND-SEMESTER", 
    "ND-SECOND-YEAR-FIRST-SEMESTER",
    "ND-SECOND-YEAR-SECOND-SEMESTER"
]

def detect_semester_from_filename(filename):
    """
    SIMPLIFIED and IMPROVED semester detection
    """
    filename_upper = filename.upper().replace('_', '-').replace(' ', '-')
    
    print(f"🔍 DEBUG FILENAME: '{filename}' -> '{filename_upper}'")
    
    # Exact matches first
    exact_matches = {
        'FIRST-YEAR-FIRST-SEMESTER': "ND-FIRST-YEAR-FIRST-SEMESTER",
        'FIRST-YEAR-SECOND-SEMESTER': "ND-FIRST-YEAR-SECOND-SEMESTER", 
        'SECOND-YEAR-FIRST-SEMESTER': "ND-SECOND-YEAR-FIRST-SEMESTER",
        'SECOND-YEAR-SECOND-SEMESTER': "ND-SECOND-YEAR-SECOND-SEMESTER"
    }
    
    for pattern, semester_key in exact_matches.items():
        if pattern in filename_upper:
            print(f"   ✅ EXACT MATCH: '{pattern}' -> '{semester_key}'")
            return semester_key, 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
    
    # Try partial matches as fallback
    if 'FIRST' in filename_upper and 'SECOND' not in filename_upper and 'YEAR' in filename_upper:
        result = "ND-FIRST-YEAR-FIRST-SEMESTER", 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
        print(f"   ✅ PARTIAL MATCH: First Year First Semester")
    elif 'FIRST' in filename_upper and 'SECOND' in filename_upper and 'YEAR' in filename_upper:
        result = "ND-FIRST-YEAR-SECOND-SEMESTER", 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI"
        print(f"   ✅ PARTIAL MATCH: First Year Second Semester")
    elif 'SECOND' in filename_upper and 'FIRST' in filename_upper and 'YEAR' in filename_upper:
        result = "ND-SECOND-YEAR-FIRST-SEMESTER", 2, 1, "YEAR TWO", "FIRST SEMESTER", "NDII"
        print(f"   ✅ PARTIAL MATCH: Second Year First Semester")
    elif 'SECOND' in filename_upper and 'YEAR' in filename_upper:
        result = "ND-SECOND-YEAR-SECOND-SEMESTER", 2, 2, "YEAR TWO", "SECOND SEMESTER", "NDII"
        print(f"   ✅ PARTIAL MATCH: Second Year Second Semester")
    else:
        # Final fallback - check individual components
        if 'FIRST-SEMESTER' in filename_upper:
            result = "ND-FIRST-YEAR-FIRST-SEMESTER", 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
            print(f"   ✅ FALLBACK MATCH: First Semester")
        elif 'SECOND-SEMESTER' in filename_upper:
            result = "ND-FIRST-YEAR-SECOND-SEMESTER", 1, 2, "YEAR ONE", "SECOND SEMESTER", "NDI"
            print(f"   ✅ FALLBACK MATCH: Second Semester")
        else:
            result = "ND-FIRST-YEAR-FIRST-SEMESTER", 1, 1, "YEAR ONE", "FIRST SEMESTER", "NDI"
            print(f"   ⚠️  DEFAULT FALLBACK: Using First Year First Semester")
    
    return result

def debug_semester_detection(directory_path):
    """Test semester detection on actual files"""
    print(f"\n{'='*60}")
    print(f"DEBUG SEMESTER DETECTION")
    print(f"Directory: {directory_path}")
    print(f"{'='*60}")
    
    if not os.path.exists(directory_path):
        print(f"❌ Directory not found: {directory_path}")
        return
    
    files = [f for f in os.listdir(directory_path) if f.endswith(('.xlsx', '.xls'))]
    print(f"Found {len(files)} Excel files:")
    
    for file in files:
        print(f"\n📄 File: {file}")
        detected_sem, year, sem_num, level, sem_display, set_code = detect_semester_from_filename(file)
        print(f"   🎯 Result: {detected_sem}")

def get_available_sets():
    """Get available ND sets"""
    nd_dir = os.path.join(BASE_DIR, "ND")
    if not os.path.exists(nd_dir):
        print(f"❌ ND directory not found: {nd_dir}")
        return []
        
    sets = []
    for item in os.listdir(nd_dir):
        item_path = os.path.join(nd_dir, item)
        if os.path.isdir(item_path) and item.upper().startswith("ND-"):
            sets.append(item)
    return sorted(sets)

def test_individual_semester_processing(set_name, semester_key):
    """Test processing for a specific semester"""
    print(f"\n{'='*60}")
    print(f"TESTING: {set_name} - {semester_key}")
    print(f"{'='*60}")
    
    raw_dir = os.path.join(BASE_DIR, "ND", set_name, "RAW_RESULTS")
    if not os.path.exists(raw_dir):
        print(f"❌ Raw directory not found: {raw_dir}")
        return
    
    files = [f for f in os.listdir(raw_dir) if f.endswith(('.xlsx', '.xls'))]
    print(f"All files in {set_name}: {files}")
    
    # Filter files for the specific semester
    semester_files = []
    for file in files:
        detected_sem, _, _, _, _, _ = detect_semester_from_filename(file)
        print(f"   {file} -> {detected_sem}")
        if detected_sem == semester_key:
            semester_files.append(file)
            print(f"   ✅ MATCH - will process")
        else:
            print(f"   ❌ NO MATCH - skipping")
    
    print(f"\n🎯 Files to process for {semester_key}: {semester_files}")
    return semester_files

def main():
    print("🔧 DEBUG SEMESTER DETECTION SCRIPT")
    
    # Get available sets
    available_sets = get_available_sets()
    print(f"Available sets: {available_sets}")
    
    if not available_sets:
        print("No sets found. Exiting.")
        return
    
    # Test each set
    for set_name in available_sets:
        raw_dir = os.path.join(BASE_DIR, "ND", set_name, "RAW_RESULTS")
        debug_semester_detection(raw_dir)
    
    # Test individual semester processing
    print(f"\n{'='*60}")
    print("TESTING INDIVIDUAL SEMESTER PROCESSING")
    print(f"{'='*60}")
    
    test_semesters = [
        "ND-FIRST-YEAR-FIRST-SEMESTER",
        "ND-FIRST-YEAR-SECOND-SEMESTER", 
        "ND-SECOND-YEAR-FIRST-SEMESTER",
        "ND-SECOND-YEAR-SECOND-SEMESTER"
    ]
    
    for set_name in available_sets:
        for semester in test_semesters:
            test_individual_semester_processing(set_name, semester)

if __name__ == "__main__":
    main()