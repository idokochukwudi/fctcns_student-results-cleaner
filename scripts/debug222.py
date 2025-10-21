#!/usr/bin/env python3
"""
QUICK TEST - Find where processing fails
"""

import os
import sys
import pandas as pd
from datetime import datetime

def get_base_directory():
    return os.path.join(os.path.expanduser('~'), 'student_result_cleaner', 'EXAMS_INTERNAL')

BASE_DIR = get_base_directory()

def test_file_loading(file_path):
    """Test if we can load and read the Excel file"""
    print(f"\n🔍 TESTING FILE: {file_path}")
    
    if not os.path.exists(file_path):
        print("❌ File does not exist")
        return False
        
    try:
        # Try to open the Excel file
        xl = pd.ExcelFile(file_path)
        print(f"✅ File opened successfully")
        print(f"📋 Sheets: {xl.sheet_names}")
        
        # Try to read each expected sheet
        expected_sheets = ['CA', 'OBJ', 'EXAM']
        for sheet in expected_sheets:
            if sheet in xl.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet)
                    print(f"✅ Sheet '{sheet}': {df.shape} rows, {df.columns.tolist()}")
                    
                    # Show first few rows for debugging
                    if not df.empty:
                        print(f"   Sample data:")
                        for i in range(min(2, len(df))):
                            # Show first 3 columns
                            sample_cols = min(3, len(df.columns))
                            sample_data = {df.columns[j]: df.iloc[i][j] for j in range(sample_cols)}
                            print(f"   Row {i}: {sample_data}")
                    else:
                        print(f"   ⚠️ Sheet '{sheet}' is empty")
                        
                except Exception as e:
                    print(f"❌ Error reading sheet '{sheet}': {e}")
            else:
                print(f"❌ Sheet '{sheet}' not found")
                
        return True
        
    except Exception as e:
        print(f"❌ Error opening file: {e}")
        return False

def main():
    print("🔧 QUICK FILE PROCESSING TEST")
    
    # Test a specific file
    test_sets = ['ND-2024', 'ND-2025']
    test_semester = "ND-FIRST-YEAR-FIRST-SEMESTER"
    
    for set_name in test_sets:
        file_path = os.path.join(BASE_DIR, "ND", set_name, "RAW_RESULTS", "FIRST-YEAR-FIRST-SEMESTER.xlsx")
        print(f"\n{'='*60}")
        print(f"TESTING: {set_name} - {test_semester}")
        print(f"{'='*60}")
        
        success = test_file_loading(file_path)
        if success:
            print(f"✅ {set_name}: File loading successful")
        else:
            print(f"❌ {set_name}: File loading failed")

if __name__ == "__main__":
    main()