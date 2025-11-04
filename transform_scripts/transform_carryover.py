import pandas as pd
import random
import sys
import os
from datetime import datetime
from pathlib import Path
from collections import defaultdict

def detect_file_format(file_path):
    """Detect if file is Excel or CSV based on extension."""
    extension = Path(file_path).suffix.lower()
    if extension in ['.xlsx', '.xls', '.xlsm']:
        return 'excel'
    elif extension == '.csv':
        return 'csv'
    else:
        return None

def load_data(file_path):
    """Load data from Excel or CSV file with automatic detection."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    file_format = detect_file_format(file_path)
    
    if file_format == 'excel':
        df = pd.read_excel(file_path)
    elif file_format == 'csv':
        # Try different encodings and delimiters
        try:
            df = pd.read_csv(file_path, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_path, encoding='latin-1')
            except:
                df = pd.read_csv(file_path, encoding='cp1252')
    else:
        raise ValueError(f"Unsupported file format. Please use .xlsx, .xls, .xlsm, or .csv files.")
    
    return df

def detect_columns(df):
    """Automatically detect column names for exam number, name, and course code."""
    columns = df.columns.str.strip().str.upper()
    df.columns = columns
    
    # Possible column name variations
    exam_variations = ['EXAM NUMBER', 'EXAMS NUMBER', 'EXAM NO', 'EXAMNO', 'EXAM_NUMBER', 
                       'EXAMS_NUMBER', 'STUDENT ID', 'STUDENT_ID', 'STUDENTID', 'ID', 
                       'MATRIC NO', 'MATRIC_NO', 'MATRICNO']
    name_variations = ['NAME', 'STUDENT NAME', 'STUDENT_NAME', 'STUDENTNAME', 'FULL NAME', 'FULLNAME']
    course_variations = ['COURSE CODE', 'COURSE_CODE', 'COURSECODE', 'COURSE', 'CODE', 'SUBJECT CODE', 'SUBJECT']
    
    exam_col = None
    name_col = None
    course_col = None
    
    for col in df.columns:
        if col in exam_variations and exam_col is None:
            exam_col = col
        if col in name_variations and name_col is None:
            name_col = col
        if col in course_variations and course_col is None:
            course_col = col
    
    return exam_col, name_col, course_col

def find_data_files(directory):
    """Find all Excel and CSV files in the directory, excluding transformed files."""
    data_files = []
    supported_extensions = ['.xlsx', '.xls', '.xlsm', '.csv']
    
    if not os.path.exists(directory):
        print(f"❌ Directory not found: {directory}")
        return []
    
    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)
        
        # Skip if not a file
        if not os.path.isfile(file_path):
            continue
        
        # Skip if already transformed
        if file.startswith('transformed_'):
            continue
        
        # Skip if it's a Python script
        if file.endswith('.py'):
            continue
        
        # Check if it's a supported format
        if any(file.lower().endswith(ext) for ext in supported_extensions):
            data_files.append(file_path)
    
    return sorted(data_files)

def get_file_type(filename):
    """Determine file type based on name for grouping."""
    lower_name = filename.lower()
    if 'nd' in lower_name:
        return 'nd'
    elif 'set' in lower_name:
        return 'set'
    else:
        return 'other'

def convert_carryover_to_resit_format(carryover_file_path, min_score=50, max_score=80, 
                                     output_format='excel', output_dir=None):
    """
    Convert carryover long format to resit wide format with random passing scores.
    
    Args:
        carryover_file_path: Path to input file (Excel or CSV)
        min_score: Minimum random score to generate (default: 50)
        max_score: Maximum random score to generate (default: 80)
        output_format: 'excel' or 'csv' (default: 'excel')
        output_dir: Directory to save output (default: same as input)
    """
    
    print(f"\n{'='*70}")
    print(f"📂 Processing: {Path(carryover_file_path).name}")
    print(f"{'='*70}")
    
    # Load the data
    try:
        df = load_data(carryover_file_path)
        print(f"✅ Loaded file successfully: {len(df)} records")
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return None
    
    # Detect column names
    exam_col, name_col, course_col = detect_columns(df)
    
    if not all([exam_col, name_col, course_col]):
        print("❌ Could not detect required columns automatically.")
        print(f"   Available columns: {list(df.columns)}")
        print("\n📋 Required columns:")
        print("   - Exam Number/Student ID")
        print("   - Student Name")
        print("   - Course Code")
        print("\n⏭️  Skipping this file...\n")
        return None
    
    print(f"🔍 Detected columns:")
    print(f"   Exam Number: {exam_col}")
    print(f"   Name: {name_col}")
    print(f"   Course Code: {course_col}")
    
    # Standardize column names for processing
    df = df.rename(columns={
        exam_col: 'EXAM_NUMBER',
        name_col: 'NAME',
        course_col: 'COURSE_CODE'
    })
    
    # Clean data
    df['EXAM_NUMBER'] = df['EXAM_NUMBER'].astype(str).str.strip()
    df['NAME'] = df['NAME'].astype(str).str.strip()
    df['COURSE_CODE'] = df['COURSE_CODE'].astype(str).str.strip()
    
    # Remove any rows with missing critical data
    original_count = len(df)
    df = df.dropna(subset=['EXAM_NUMBER', 'NAME', 'COURSE_CODE'])
    if len(df) < original_count:
        print(f"⚠️  Removed {original_count - len(df)} rows with missing data")
    
    # Get unique courses
    courses = sorted(df['COURSE_CODE'].unique())
    print(f"📚 Found {len(courses)} unique courses: {courses}")
    
    # Get unique students
    students = df[['EXAM_NUMBER', 'NAME']].drop_duplicates()
    print(f"👨‍🎓 Found {len(students)} unique students")
    
    # Count failures per course
    course_failures = df['COURSE_CODE'].value_counts()
    print("\n📊 Failures per course:")
    for course, count in course_failures.items():
        print(f"   {course}: {count} students")
    
    # Create the wide format resit file
    resit_data = []
    
    for _, student in students.iterrows():
        exam_no = student['EXAM_NUMBER']
        name = student['NAME']
        
        # Get all failed courses for this student
        student_failed_courses = df[df['EXAM_NUMBER'] == exam_no]
        failed_course_codes = student_failed_courses['COURSE_CODE'].tolist()
        
        # Create row with exam number and name
        row = {'EXAM NUMBER': exam_no, 'NAME': name}
        
        # Add random passing scores for each course the student failed
        for course in courses:
            if course in failed_course_codes:
                random_score = random.randint(min_score, max_score)
                row[course] = random_score
            else:
                row[course] = ''
        
        resit_data.append(row)
    
    # Create DataFrame
    resit_df = pd.DataFrame(resit_data)
    
    # Reorder columns: EXAM NUMBER, NAME, then all courses
    columns_order = ['EXAM NUMBER', 'NAME'] + list(courses)
    resit_df = resit_df[columns_order]
    
    # Generate output filename
    input_path = Path(carryover_file_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Determine output directory
    if output_dir:
        output_directory = Path(output_dir)
        output_directory.mkdir(parents=True, exist_ok=True)
    else:
        output_directory = input_path.parent
    
    if output_format.lower() == 'csv':
        output_file = output_directory / f"transformed_{input_path.stem}_{timestamp}.csv"
        resit_df.to_csv(output_file, index=False)
    else:
        output_file = output_directory / f"transformed_{input_path.stem}_{timestamp}.xlsx"
        resit_df.to_excel(output_file, index=False)
    
    print(f"\n✅ Converted resit file saved: {output_file.name}")
    print(f"📊 Final format: {len(resit_df)} students × {len(courses)} courses")
    
    # Show summary of generated scores
    print(f"\n🎲 Score Summary (Randomly generated {min_score}-{max_score}):")
    for course in courses:
        course_scores = resit_df[course].replace('', pd.NA).dropna()
        if len(course_scores) > 0:
            avg_score = course_scores.mean()
            min_gen = course_scores.min()
            max_gen = course_scores.max()
            print(f"   {course}: {len(course_scores)} students, Avg: {avg_score:.1f}, Range: {min_gen}-{max_gen}")
    
    return resit_df

def process_directory(directory, min_score=50, max_score=80, output_format='excel'):
    """Process all data files in a directory, grouping by type into timestamped folders."""
    
    print(f"\n🔍 Scanning directory: {directory}")
    data_files = find_data_files(directory)
    
    if not data_files:
        print(f"❌ No data files found in {directory}")
        print("   Looking for: .xlsx, .xls, .xlsm, or .csv files")
        print("   Excluding: transformed_* files and .py scripts")
        return
    
    print(f"✅ Found {len(data_files)} file(s) to process:")
    for i, file in enumerate(data_files, 1):
        print(f"   {i}. {Path(file).name}")
    
    # Group files by type
    groups = defaultdict(list)
    for file_path in data_files:
        filename = Path(file_path).name
        ftype = get_file_type(filename)
        groups[ftype].append(file_path)
    
    # Get timestamp once for all folders
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    input_dir = Path(directory)
    
    # Process each group
    successful = 0
    failed = 0
    
    for ftype, files in groups.items():
        if not files:
            continue
        
        if ftype == 'other':
            output_dir = input_dir / f"transformed_{timestamp}"
        else:
            output_dir = input_dir / f"{ftype}_transformed_{timestamp}"
        
        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"\n📂 Created output folder for {ftype}: {output_dir.name}")
        
        for file_path in files:
            try:
                result = convert_carryover_to_resit_format(
                    file_path,
                    min_score=min_score,
                    max_score=max_score,
                    output_format=output_format,
                    output_dir=str(output_dir)
                )
                if result is not None:
                    successful += 1
                else:
                    failed += 1
            except Exception as e:
                print(f"\n❌ Error processing {Path(file_path).name}: {e}")
                failed += 1
    
    # Summary
    print(f"\n{'='*70}")
    print("📋 PROCESSING SUMMARY")
    print(f"{'='*70}")
    print(f"✅ Successfully processed: {successful} file(s)")
    if failed > 0:
        print(f"❌ Failed: {failed} file(s)")
    print(f"{'='*70}\n")

# Main execution
if __name__ == "__main__":
    print("🎯 UNIVERSAL CARRYOVER TO RESIT CONVERTER")
    print("=" * 70)
    print("Supports: Excel (.xlsx, .xls, .xlsm) and CSV files")
    print("Auto-detects files in current directory")
    print("=" * 70)
    
    # Determine default directory intelligently
    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_dir_name = os.path.basename(script_dir)
    
    # Check if we're in transform_scripts directory
    if script_dir_name == "transform_scripts":
        default_directory = "."  # Current directory
    else:
        # Check if transform_scripts exists as subdirectory
        if os.path.exists(os.path.join(script_dir, "transform_scripts")):
            default_directory = "transform_scripts"
        else:
            default_directory = "."  # Use current directory
    
    # Parse command line arguments
    if len(sys.argv) == 1:
        # No arguments - process directory mode
        abs_directory = os.path.abspath(default_directory)
        print(f"\n🤖 AUTO MODE: Processing all files in current location")
        
        min_score = 50
        max_score = 80
        output_format = 'excel'
        
        print(f"\n⚙️ Default Settings:")
        print(f"   Directory: {abs_directory}")
        print(f"   Score range: {min_score}-{max_score}")
        print(f"   Output format: {output_format}")
        
        process_directory(default_directory, min_score, max_score, output_format)
        
    elif sys.argv[1] in ['-h', '--help', 'help']:
        print("\n📖 USAGE:")
        print(f"\n1️⃣ Auto-process all files in transform_scripts/:")
        print(f"   python {sys.argv[0]}")
        
        print(f"\n2️⃣ Process specific file:")
        print(f"   python {sys.argv[0]} <input_file> [min_score] [max_score] [output_format]")
        
        print(f"\n3️⃣ Process specific directory:")
        print(f"   python {sys.argv[0]} --dir <directory> [min_score] [max_score] [output_format]")
        
        print("\n💡 EXAMPLES:")
        print(f"   python {sys.argv[0]}")
        print(f"   python {sys.argv[0]} data/carryover.xlsx")
        print(f"   python {sys.argv[0]} data/carryover.csv 45 75")
        print(f"   python {sys.argv[0]} --dir my_data/ 50 80 csv")
        
    elif sys.argv[1] in ['--dir', '-d']:
        # Directory mode with custom directory
        if len(sys.argv) < 3:
            print("❌ Error: Please specify a directory")
            print(f"   Usage: python {sys.argv[0]} --dir <directory>")
            sys.exit(1)
        
        directory = sys.argv[2]
        min_score = int(sys.argv[3]) if len(sys.argv) > 3 else 50
        max_score = int(sys.argv[4]) if len(sys.argv) > 4 else 80
        output_format = sys.argv[5] if len(sys.argv) > 5 else 'excel'
        
        if min_score >= max_score:
            print(f"❌ Error: min_score ({min_score}) must be less than max_score ({max_score})")
            sys.exit(1)
        
        print(f"\n⚙️ Settings:")
        print(f"   Directory: {directory}")
        print(f"   Score range: {min_score}-{max_score}")
        print(f"   Output format: {output_format}")
        
        process_directory(directory, min_score, max_score, output_format)
        
    else:
        # Single file mode
        input_file = sys.argv[1]
        min_score = int(sys.argv[2]) if len(sys.argv) > 2 else 50
        max_score = int(sys.argv[3]) if len(sys.argv) > 3 else 80
        output_format = sys.argv[4] if len(sys.argv) > 4 else 'excel'
        
        if min_score >= max_score:
            print(f"❌ Error: min_score ({min_score}) must be less than max_score ({max_score})")
            sys.exit(1)
        
        print(f"\n⚙️ Settings:")
        print(f"   Input file: {input_file}")
        print(f"   Score range: {min_score}-{max_score}")
        print(f"   Output format: {output_format}")
        
        try:
            # Determine type and create folder
            input_path = Path(input_file)
            ftype = get_file_type(input_path.name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            if ftype == 'other':
                output_dir = input_path.parent / f"transformed_{timestamp}"
            else:
                output_dir = input_path.parent / f"{ftype}_transformed_{timestamp}"
            
            output_dir.mkdir(parents=True, exist_ok=True)
            print(f"\n📂 Created output folder for {ftype}: {output_dir.name}")
            
            result = convert_carryover_to_resit_format(
                input_file,
                min_score=min_score,
                max_score=max_score,
                output_format=output_format,
                output_dir=str(output_dir)
            )
            
            if result is not None:
                print("\n✨ Conversion completed successfully!")
            else:
                print("\n❌ Conversion failed.")
            
        except Exception as e:
            print(f"\n❌ Error during conversion: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)