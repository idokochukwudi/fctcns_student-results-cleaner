# Save as test_detection.py
import re

def detect_bn_semester_from_filename(filename):
    """Test version"""
    filename_upper = filename.upper()
    
    semester_patterns = {
        r'(?:N[-_\s])?(?:THIRD|3RD|3)[-_\s]*(?:YEAR|YR)[-_\s]*(?:SECOND|2ND|2)[-_\s]*(?:SEMESTER|SEM)': "N-THIRD-YEAR-SECOND-SEMESTER",
    }
    
    for pattern, semester_key in semester_patterns.items():
        match = re.search(pattern, filename_upper)
        if match:
            print(f"✅ MATCHED: '{filename}' → '{semester_key}'")
            print(f"   Pattern: {pattern}")
            print(f"   Matched text: '{match.group()}'")
            return semester_key
    
    print(f"❌ NO MATCH for: '{filename}'")
    return None

# Test with your actual filename
test_files = [
    "Third Year Second Semester.xlsx",
    "N-Third-Year-Second-Semester.xlsx",
    "3rd Year 2nd Semester.xlsx",
    "THIRD_YEAR_SECOND_SEMESTER.xlsx",
    "Year 3 Semester 2.xlsx",
    # Add your ACTUAL filename here:
    "YOUR_ACTUAL_FILENAME.xlsx"
]

for test_file in test_files:
    print(f"\nTesting: {test_file}")
    detect_bn_semester_from_filename(test_file)
    print("-" * 60)