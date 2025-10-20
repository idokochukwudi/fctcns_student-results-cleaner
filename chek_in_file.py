#!/usr/bin/env python3
"""
Debug script to check ND file structures and course matching
"""

import pandas as pd
import os

def check_nd_files():
    # Check ND-2024 files
    print('=== ND-2024 FILES ===')
    nd2024_dir = 'EXAMS_INTERNAL/ND/ND-2024/RAW_RESULTS'
    if os.path.exists(nd2024_dir):
        for file in os.listdir(nd2024_dir):
            if file.endswith('.xlsx'):
                print(f'\nFile: {file}')
                try:
                    xl = pd.ExcelFile(os.path.join(nd2024_dir, file))
                    print(f'Sheets: {xl.sheet_names}')
                    for sheet in xl.sheet_names:
                        df = pd.read_excel(os.path.join(nd2024_dir, file), sheet_name=sheet, nrows=3)
                        print(f'  {sheet} columns: {df.columns.tolist()}')
                        # Show first row of data to see actual values
                        if not df.empty:
                            print(f'  First row data:')
                            for col in df.columns:
                                if col not in ['REG. No', 'NAME', 'Exam No']:
                                    val = df[col].iloc[0] if len(df) > 0 else 'N/A'
                                    print(f'    {col}: {val}')
                except Exception as e:
                    print(f'  ERROR reading {file}: {e}')
    else:
        print(f'Directory not found: {nd2024_dir}')

    # Check ND-2025 files for comparison
    print('\n=== ND-2025 FILES (for comparison) ===')
    nd2025_dir = 'EXAMS_INTERNAL/ND/ND-2025/RAW_RESULTS'
    if os.path.exists(nd2025_dir):
        for file in os.listdir(nd2025_dir)[:1]:  # Just check first file
            if file.endswith('.xlsx'):
                print(f'\nFile: {file}')
                try:
                    xl = pd.ExcelFile(os.path.join(nd2025_dir, file))
                    print(f'Sheets: {xl.sheet_names}')
                    for sheet in xl.sheet_names[:1]:  # Just first sheet
                        df = pd.read_excel(os.path.join(nd2025_dir, file), sheet_name=sheet, nrows=3)
                        print(f'  {sheet} columns: {df.columns.tolist()}')
                        if not df.empty:
                            print(f'  First row data:')
                            for col in df.columns[:3]:  # Just first 3 columns
                                val = df[col].iloc[0] if len(df) > 0 else 'N/A'
                                print(f'    {col}: {val}')
                except Exception as e:
                    print(f'  ERROR reading {file}: {e}')
    else:
        print(f'Directory not found: {nd2025_dir}')

    # Check course file
    print('\n=== COURSE FILE ===')
    course_file = 'EXAMS_INTERNAL/ND/ND-COURSES/course-code-creditUnit.xlsx'
    if os.path.exists(course_file):
        try:
            xl = pd.ExcelFile(course_file)
            for sheet in xl.sheet_names:
                print(f'\nSheet: {sheet}')
                df = pd.read_excel(course_file, sheet_name=sheet)
                if 'COURSE TITLE' in df.columns and 'COURSE CODE' in df.columns:
                    print('First 5 courses:')
                    for i, row in df.head().iterrows():
                        title = row['COURSE TITLE']
                        code = row['COURSE CODE']
                        print(f'  "{title}" -> {code}')
                else:
                    print(f'  Required columns not found. Available: {df.columns.tolist()}')
        except Exception as e:
            print(f'ERROR reading course file: {e}')
    else:
        print(f'Course file not found: {course_file}')

if __name__ == "__main__":
    check_nd_files()