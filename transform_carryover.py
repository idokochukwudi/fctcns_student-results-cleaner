import pandas as pd
import random
import sys
from datetime import datetime

def convert_carryover_to_resit_format(carryover_file_path, min_score=50, max_score=80):
    """
    Convert carryover long format to resit wide format with random passing scores.
    Outputs a file with prefix 'transformed_'.
    """
    
    # Determine output file path
    output_file_path = f"transformed_{carryover_file_path}"
    
    # Read the carryover file
    try:
        df = pd.read_excel(carryover_file_path)
        print(f"✅ Loaded carryover file: {len(df)} records from {carryover_file_path}")
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return None
    
    # Get unique courses
    courses = sorted(df['COURSE CODE'].unique())
    print(f"📚 Found {len(courses)} unique courses: {courses}")
    
    # Get unique students
    students = df[['EXAM NUMBER', 'NAME']].drop_duplicates()
    print(f"👨‍🎓 Found {len(students)} unique students")
    
    # Count failures per course
    course_failures = df['COURSE CODE'].value_counts()
    print("\n📊 Failures per course:")
    for course, count in course_failures.items():
        print(f"   {course}: {count} students")
    
    # Create the wide format resit file
    resit_data = []
    
    for _, student in students.iterrows():
        exam_no = student['EXAM NUMBER']
        name = student['NAME']
        
        # Get all failed courses for this student
        student_failed_courses = df[df['EXAM NUMBER'] == exam_no]
        failed_course_codes = student_failed_courses['COURSE CODE'].tolist()
        
        # Create row with exam number and name
        row = {'EXAM NUMBER': exam_no, 'NAME': name}
        
        # Add random passing scores for each course the student failed
        for course in courses:
            if course in failed_course_codes:
                # Generate random passing score between min_score and max_score
                random_score = random.randint(min_score, max_score)
                row[course] = random_score
            else:
                # Student didn't fail this course, leave blank
                row[course] = ''
        
        resit_data.append(row)
    
    # Create DataFrame and save
    resit_df = pd.DataFrame(resit_data)
    
    # Reorder columns: EXAM NUMBER, NAME, then all courses
    columns_order = ['EXAM NUMBER', 'NAME'] + list(courses)
    resit_df = resit_df[columns_order]
    
    # Add timestamp to filename to avoid overwriting
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file_path = output_file_path.replace('.xlsx', f'_{timestamp}.xlsx')
    
    # Save to Excel
    resit_df.to_excel(output_file_path, index=False)
    
    print(f"\n✅ Converted resit file saved: {output_file_path}")
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

# Main execution
if __name__ == "__main__":
    print("🎯 CARRYOVER TO RESIT CONVERTER")
    print("=" * 50)
    
    # Get input file from command line argument or use default
    if len(sys.argv) > 1:
        carryover_file = sys.argv[1]
    else:
        # Default file name - replace with your common file name if needed
        carryover_file = "co_student_ND-2025_ND-SECOND-YEAR-FIRST-SEMESTER_20251029_110724.xlsx"
        print(f"⚠️ No file specified, using default: {carryover_file}")
    
    # Convert to resit format
    print("\n🔄 Converting to resit format with random passing scores...")
    convert_carryover_to_resit_format(carryover_file, min_score=50, max_score=80)
    
    print("\n🎯 NEXT STEPS:")
    print("1. To process a different file, run: python script.py your_file.xlsx")
    print("2. The script assigns random scores between 50-80 for failed courses.")
    print("3. Output file has prefix 'transformed_' and includes a timestamp.")