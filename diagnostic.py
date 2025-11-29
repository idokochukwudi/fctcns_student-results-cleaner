#!/usr/bin/env python3
"""
Diagnostic tool to inspect the resit file structure
"""

import pandas as pd
import sys
import os

def diagnose_resit_file(file_path):
    """Diagnose the structure of the resit file"""
    print("="*80)
    print("🔍 RESIT FILE DIAGNOSTIC TOOL")
    print("="*80)
    
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        return
    
    print(f"📁 File: {file_path}")
    print()
    
    # Try reading with different header rows
    for header_row in range(0, 5):
        print(f"\n{'='*80}")
        print(f"📊 ATTEMPTING TO READ WITH HEADER ROW: {header_row}")
        print(f"{'='*80}")
        
        try:
            df = pd.read_excel(file_path, header=header_row)
            
            print(f"\n✅ Successfully read file with header row {header_row}")
            print(f"📊 Shape: {df.shape[0]} rows × {df.shape[1]} columns")
            print()
            
            # Print column names
            print("📋 COLUMN NAMES:")
            for idx, col in enumerate(df.columns):
                print(f"  [{idx}] '{col}'")
            print()
            
            # Print first 3 rows
            print("📄 FIRST 3 ROWS OF DATA:")
            print(df.head(3).to_string())
            print()
            
            # Check for EXAM NUMBER column
            print("🔍 SEARCHING FOR EXAM NUMBER COLUMN:")
            exam_col = None
            for col in df.columns:
                col_upper = str(col).upper()
                if "EXAM" in col_upper and "NUMBER" in col_upper:
                    exam_col = col
                    print(f"  ✅ FOUND: '{col}' (column index: {df.columns.get_loc(col)})")
                    break
            
            if not exam_col:
                print("  ❌ NOT FOUND - No column contains 'EXAM NUMBER'")
                print(f"  💡 Available columns: {list(df.columns)}")
            
            # Print sample values from EXAM NUMBER column if found
            if exam_col:
                print(f"\n📊 SAMPLE VALUES FROM '{exam_col}':")
                print(df[exam_col].head(5).to_string(index=False))
            
            print()
            
        except Exception as e:
            print(f"❌ Failed to read with header row {header_row}: {e}")
            continue
    
    print("\n" + "="*80)
    print("✅ DIAGNOSTIC COMPLETE")
    print("="*80)


if __name__ == "__main__":
    # Get file path from command line or environment variable
    file_path = sys.argv[1] if len(sys.argv) > 1 else os.getenv("RESIT_FILE_PATH")
    
    if not file_path:
        print("❌ ERROR: Please provide resit file path")
        print("Usage: python diagnostic.py <path_to_resit_file>")
        sys.exit(1)
    
    diagnose_resit_file(file_path)