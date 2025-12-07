#!/usr/bin/env python3
"""
Script: clean_viva.py
Description: Clean viva data by extracting MAT numbers and scores from the provided format.
Usage: python clean_viva.py [input_file] [output_file]
"""

import pandas as pd
import re
import sys
from pathlib import Path

def extract_mat_number(full_name):
    """
    Extract MAT number from full name string using pattern BN/A23/ followed by digits.
    
    Args:
        full_name (str): Full name string containing MAT number
        
    Returns:
        tuple: (cleaned_name, mat_number) or (cleaned_name, None) if not found
    """
    # Pattern for MAT number: BN/A23/ followed by digits
    pattern = r'(BN/A23/\d{3})'
    match = re.search(pattern, str(full_name))
    
    if match:
        mat_number = match.group(1)
        # Remove the MAT number from the full name
        cleaned_name = str(full_name).replace(mat_number, '').strip()
        return cleaned_name, mat_number
    return str(full_name), None

def clean_viva_data(input_file=None, output_file=None):
    """
    Clean viva data by extracting MAT numbers and reorganizing columns.
    
    Args:
        input_file (str): Path to input CSV/TSV file. If None, will try to find one.
        output_file (str): Path to output CSV file. If None, will create one.
    """
    # Try to find input file if not specified
    if input_file is None:
        # Look for CSV/TSV files in current directory
        csv_files = list(Path('.').glob('*.csv')) + list(Path('.').glob('*.tsv'))
        if not csv_files:
            print("No CSV or TSV files found in current directory!")
            return
        input_file = str(csv_files[0])
        print(f"Using input file: {input_file}")
    
    # Set output file if not specified
    if output_file is None:
        input_path = Path(input_file)
        output_file = f"cleaned_{input_path.stem}.csv"
    
    try:
        print(f"Processing file: {input_file}")
        print("-" * 50)
        
        # First, let's inspect the file to determine the separator
        with open(input_file, 'r', encoding='utf-8') as f:
            first_lines = [f.readline() for _ in range(3)]
        
        # Check if it's tab-separated (based on your data format)
        if '\t' in first_lines[0]:
            print("Detected tab-separated values (TSV format)")
            df = pd.read_csv(input_file, sep='\t', encoding='utf-8')
        else:
            print("Detected comma-separated values (CSV format)")
            df = pd.read_csv(input_file, encoding='utf-8')
        
        # Display file information
        print(f"\nOriginal shape: {df.shape[0]} rows × {df.shape[1]} columns")
        print(f"Columns found: {list(df.columns)}")
        
        # Clean column names (remove trailing/leading whitespace and normalize)
        df.columns = df.columns.str.strip()
        
        print("\nFirst 5 rows of original data:")
        print("-" * 50)
        print(df.head())
        print("-" * 50)
        
        # Check for required columns
        required_columns = ['User full name']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"\n❌ Error: Missing required column: 'User full name'")
            print(f"Available columns: {list(df.columns)}")
            
            # Try to find similar column names
            for col in df.columns:
                if 'user' in col.lower() or 'name' in col.lower() or 'full' in col.lower():
                    print(f"  Suggestion: Use '{col}' instead of 'User full name'")
            return
        
        # Find score column
        score_column = None
        # List of possible score column names (case-insensitive)
        score_patterns = [
            'enter student score below',
            'enter student score',
            'student score below',
            'score below',
            'score'
        ]
        
        for col in df.columns:
            col_lower = col.lower().strip()
            for pattern in score_patterns:
                if pattern in col_lower:
                    score_column = col
                    break
            if score_column:
                break
        
        if score_column is None:
            print(f"\n❌ Error: Could not find score column!")
            print(f"Available columns: {list(df.columns)}")
            print("\nThe score column might be named differently. Please check your file.")
            print("Looking for columns containing: 'score', 'enter student score', etc.")
            return
        
        print(f"\n✓ Using score column: '{score_column}'")
        
        # Create lists for new columns
        cleaned_names = []
        mat_numbers = []
        scores = []
        
        # Process each row
        print(f"\nProcessing {len(df)} rows...")
        
        for idx, row in df.iterrows():
            # Get full name
            full_name = row['User full name']
            
            # Get score
            score = row[score_column]
            
            # Extract MAT number
            cleaned_name, mat_number = extract_mat_number(full_name)
            
            # Append to lists
            cleaned_names.append(cleaned_name)
            mat_numbers.append(mat_number if mat_number else 'N/A')
            
            # Clean and convert score
            if pd.isna(score):
                scores.append('')
            else:
                # Convert to string and strip whitespace
                score_str = str(score).strip()
                # Try to convert to numeric if possible
                try:
                    # Check if it's a number
                    if score_str.replace('.', '', 1).isdigit():
                        scores.append(float(score_str))
                    else:
                        scores.append(score_str)
                except:
                    scores.append(score_str)
        
        # Create new DataFrame
        cleaned_df = pd.DataFrame({
            'Full Name': cleaned_names,
            'MAT Number': mat_numbers,
            'Score': scores
        })
        
        # Save to CSV
        cleaned_df.to_csv(output_file, index=False, encoding='utf-8')
        
        print(f"\n" + "="*50)
        print(f"✓ SUCCESS: Cleaned data saved to: {output_file}")
        print("="*50)
        
        print(f"\n📊 CLEANED DATA SUMMARY:")
        print(f"Total records: {len(cleaned_df)}")
        
        # Count valid MAT numbers
        valid_mats = cleaned_df[cleaned_df['MAT Number'] != 'N/A']
        print(f"Valid MAT numbers extracted: {len(valid_mats)}")
        
        # Check for rows without MAT numbers
        no_mats = cleaned_df[cleaned_df['MAT Number'] == 'N/A']
        if len(no_mats) > 0:
            print(f"⚠️  Rows without MAT numbers: {len(no_mats)}")
            print("These rows might need manual checking:")
            print(no_mats[['Full Name', 'Score']].head())
        
        print(f"\n📋 FIRST 10 ROWS OF CLEANED DATA:")
        print("-" * 60)
        print(cleaned_df.head(10).to_string(index=False))
        print("-" * 60)
        
        # Show score statistics if scores are numeric
        numeric_scores = pd.to_numeric(cleaned_df['Score'], errors='coerce')
        if numeric_scores.notna().any():
            print(f"\n📈 SCORE STATISTICS:")
            print(f"  Average score: {numeric_scores.mean():.2f}")
            print(f"  Highest score: {numeric_scores.max():.2f}")
            print(f"  Lowest score:  {numeric_scores.min():.2f}")
            print(f"  Score range:   {numeric_scores.min():.2f} - {numeric_scores.max():.2f}")
        
        print(f"\n💾 File saved successfully!")
        print(f"   Output file: {output_file}")
        print(f"   File size: {Path(output_file).stat().st_size / 1024:.2f} KB")
        
    except FileNotFoundError:
        print(f"\n❌ Error: File '{input_file}' not found!")
        print("Please check the file path and try again.")
    except pd.errors.EmptyDataError:
        print(f"\n❌ Error: File '{input_file}' is empty!")
    except pd.errors.ParserError as e:
        print(f"\n❌ Error: Could not parse file '{input_file}'!")
        print(f"Details: {e}")
        print("\nTry saving your file as a proper CSV or TSV file.")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()

def main():
    """
    Main function to handle command line arguments.
    """
    print("\n" + "="*60)
    print("VIVA DATA CLEANER")
    print("Extracts MAT numbers and scores from viva data")
    print("="*60 + "\n")
    
    # Get command line arguments
    if len(sys.argv) >= 2:
        input_file = sys.argv[1]
    else:
        input_file = None
    
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        output_file = None
    
    # Run the cleaner
    clean_viva_data(input_file, output_file)
    
    print("\n" + "="*60)
    print("Process completed!")
    print("="*60)

if __name__ == "__main__":
    main()