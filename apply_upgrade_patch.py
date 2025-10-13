#!/usr/bin/env python3
"""
apply_upgrade_patch.py

Safely patch exam_result_processor.py:
 - Makes a backup (exam_result_processor.py.bak_TIMESTAMP)
 - Locates the first occurrence of the `for code in ordered_codes if ordered_codes else []:` loop
 - Replaces that loop block up to and including the `sn += 1` line with a corrected, indented block
 - Leaves all other content untouched
"""

import io
import os
import re
from datetime import datetime

# Path to the target file in your repo (adjust if needed)
FILE_PATH = os.path.join("scripts", "exam_result_processor.py")

if not os.path.exists(FILE_PATH):
    print(f"ERROR: File not found: {FILE_PATH}")
    raise SystemExit(1)

# Read file
with io.open(FILE_PATH, "r", encoding="utf-8") as f:
    lines = f.readlines()

# Find the start index of the target for-loop
start_idx = None
pattern_for = re.compile(r'^\s*for\s+code\s+in\s+ordered_codes\s+if\s+ordered_codes\s+else\s+\[\]\s*:')
for i, ln in enumerate(lines):
    if pattern_for.match(ln):
        start_idx = i
        break

if start_idx is None:
    print("ERROR: Could not find 'for code in ordered_codes if ordered_codes else []:' in file.")
    raise SystemExit(1)

# Find the end index: the first subsequent line that contains 'sn += 1' (we will include that line)
end_idx = None
for j in range(start_idx+1, len(lines)):
    if re.match(r'^\s*sn\s*\+=\s*1\s*$', lines[j]):
        end_idx = j
        break

if end_idx is None:
    print("ERROR: Could not find the loop end marker 'sn += 1' after the for-loop. Aborting.")
    raise SystemExit(1)

# Prepare backup
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
backup_path = FILE_PATH + f".bak_{timestamp}"
with io.open(backup_path, "w", encoding="utf-8") as f:
    f.writelines(lines)
print(f"Backup created: {backup_path}")

# Build replacement block
# Note: This block uses center_align_style and left_align_style - if your file uses different names adjust manually.
replacement_block = [
    lines[start_idx]  # keep the original 'for...' line as-is (preserves indentation)
]

# We'll detect the indentation prefix used for the body (count spaces after for-line)
for_line = lines[start_idx]
m = re.match(r'^(\s*)for\s', for_line)
indent = m.group(1) + "    "  # body should be 4 spaces deeper than for-line

# Construct the new body with the detected indentation
body = f"""{indent}score = r.get(code)
{indent}if pd.isna(score) or score == "":
{indent}    continue

{indent}try:
{indent}    score_val = float(score)

{indent}    # -----------------------
{indent}    # FIX: Auto-upgrade borderline scores when threshold upgrade applies
{indent}    if THRESHOLD_UPGRADED and 47.0 <= score_val < 50.0:
{indent}        # log upgrade for debugging
{indent}        print(f"🔼 Upgraded score for {{r.get('EXAMS NUMBER', '')}} - {{code}}: {{score_val}} → 50.0")
{indent}        score_val = 50.0
{indent}    # -----------------------

{indent}    score_display = str(int(round(score_val)))
{indent}    grade = get_grade(score_val)
{indent}    grade_point = get_grade_point(score_val)
{indent}except Exception:
{indent}    score_display = str(score)
{indent}    grade = "F"
{indent}    grade_point = 0.0

{indent}cu = filtered_credit_units.get(code, 0) if filtered_credit_units else 0
{indent}course_title = course_titles_map.get(code, code) if course_titles_map else code

{indent}total_grade_points += grade_point * cu
{indent}total_units += cu

{indent}if score_val >= pass_threshold:
{indent}    total_units_passed += cu
{indent}else:
{indent}    total_units_failed += cu
{indent}    failed_courses_list.append(code)

{indent}course_data.append([
{indent}    Paragraph(str(sn), center_align_style),
{indent}    Paragraph(code, left_align_style),
{indent}    Paragraph(course_title, left_align_style),
{indent}    Paragraph(str(cu), center_align_style),
{indent}    Paragraph(score_display, center_align_style),
{indent}    Paragraph(grade, center_align_style)
{indent}])
{indent}sn += 1
"""

replacement_block.extend(body.splitlines(keepends=True))

# Compose new file content
new_lines = lines[:start_idx] + replacement_block + lines[end_idx+1:]

# Write back
with io.open(FILE_PATH, "w", encoding="utf-8") as f:
    f.writelines(new_lines)

print(f"Patched file written: {FILE_PATH}")
print("Done. Please run your script and verify upgrades are logged and mastersheet updated.")
