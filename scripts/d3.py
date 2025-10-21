#!/usr/bin/env python3
"""
Debug course matching
"""

import os
import pandas as pd
import re

def get_base_directory():
    return os.path.join(os.path.expanduser('~'), 'student_result_cleaner', 'EXAMS_INTERNAL')

BASE_DIR = get_base_directory()
ND_COURSES_DIR = os.path.join(BASE_DIR, "ND", "ND-COURSES")

def normalize_course_name(name):
    """Simple normalization for course title matching."""
    return re.sub(
        r'\s+',
        ' ',
        str(name).strip().lower()).replace(
        'coomunication',
        'communication')

def load_course_data():
    """
    Reads course-code-creditUnit.xlsx and returns:
      (semester_course_maps, semester_credit_units,
       semester_lookup, semester_course_titles)
    """
    course_file = os.path.join(ND_COURSES_DIR, "course-code-creditUnit.xlsx")
    print(f"Loading course data from: {course_file}")
    if not os.path.exists(course_file):
        raise FileNotFoundError(f"Course file not found: {course_file}")

    xl = pd.ExcelFile(course_file)
    semester_course_maps = {}
    semester_credit_units = {}
    semester_lookup = {}
    semester_course_titles = {}  # code -> title mapping

    for sheet in xl.sheet_names:
        df = pd.read_excel(
            course_file,
            sheet_name=sheet,
            engine='openpyxl',
            header=0)
        df.columns = [str(c).strip() for c in df.columns]
        expected = ['COURSE CODE', 'COURSE TITLE', 'CU']
        if not all(col in df.columns for col in expected):
            print(
                f"Warning: sheet '{sheet}' missing expected columns {expected} — skipped")
            continue
        dfx = df.dropna(subset=['COURSE CODE', 'COURSE TITLE'])
        dfx = dfx[~dfx['COURSE CODE'].astype(
            str).str.contains('TOTAL', case=False, na=False)]
        valid_mask = dfx['CU'].astype(str).str.replace(
            '.', '', regex=False).str.isdigit()
        dfx = dfx[valid_mask]
        if dfx.empty:
            print(
                f"Warning: sheet '{sheet}' has no valid rows after cleaning — skipped")
            continue
        codes = dfx['COURSE CODE'].astype(str).str.strip().tolist()
        titles = dfx['COURSE TITLE'].astype(str).str.strip().tolist()
        cus = dfx['CU'].astype(float).astype(int).tolist()

        semester_course_maps[sheet] = dict(zip(titles, codes))
        semester_credit_units[sheet] = dict(zip(codes, cus))
        semester_course_titles[sheet] = dict(zip(codes, titles))

    if not semester_course_maps:
        raise ValueError("No course data loaded from course workbook")
    print(f"Loaded course sheets: {list(semester_course_maps.keys())}")
    return semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles

def debug_course_matching(raw_file_path, semester_key):
    """
    Debug the course matching for a given raw file and semester key.
    """
    print(f"\n🔍 DEBUG COURSE MATCHING for {semester_key} in file {raw_file_path}")
    
    # Load course data
    semester_course_maps, semester_credit_units, semester_lookup, semester_course_titles = load_course_data()
    
    if semester_key not in semester_course_maps:
        print(f"❌ Semester {semester_key} not found in course data.")
        return
        
    course_map = semester_course_maps[semester_key]
    print(f"📚 Course map for {semester_key}:")
    for title, code in course_map.items():
        print(f"   {title} -> {code}")
    
    # Load the raw file
    xl = pd.ExcelFile(raw_file_path)
    sheets = ['CA', 'OBJ', 'EXAM']
    
    for sheet_name in sheets:
        if sheet_name not in xl.sheet_names:
            print(f"❌ Sheet {sheet_name} not found in {raw_file_path}")
            continue
            
        df = pd.read_excel(raw_file_path, sheet_name=sheet_name)
        print(f"\n📊 Sheet {sheet_name} has columns: {df.columns.tolist()}")
        
        # Check each column in the sheet against the course map
        for col in df.columns:
            if col in ['S/N', 'REG. No', 'NAME']:
                continue
                
            norm_col = normalize_course_name(col)
            print(f"   Column: '{col}' -> normalized: '{norm_col}'")
            
            matched = False
            for course_title in course_map.keys():
                norm_title = normalize_course_name(course_title)
                if norm_col == norm_title:
                    print(f"      ✅ Matched to course: '{course_title}' -> code: {course_map[course_title]}")
                    matched = True
                    break
                    
            if not matched:
                print(f"      ❌ No match found in course map")

if __name__ == "__main__":
    # Test for a specific file and semester
    test_set = "ND-2024"
    test_semester = "ND-FIRST-YEAR-FIRST-SEMESTER"
    test_file = "FIRST-YEAR-FIRST-SEMESTER.xlsx"
    
    raw_file_path = os.path.join(BASE_DIR, "ND", test_set, "RAW_RESULTS", test_file)
    
    if not os.path.exists(raw_file_path):
        print(f"❌ Raw file not found: {raw_file_path}")
    else:
        debug_course_matching(raw_file_path, test_semester)