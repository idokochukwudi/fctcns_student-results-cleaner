#!/usr/bin/env python3
"""
exam_result_processor.py

Processes examination results for ND students from EXAMS_INTERNAL/ND-*/RAW_RESULTS,
using course data from EXAMS_INTERNAL/ND-COURSES/course-code-creditUnit.xlsx.

Output mastersheets saved to EXAMS_INTERNAL/ND-*/CLEAN_RESULTS/ with separate sheets for each semester.

Features:
 - Detects courses and maps to codes/credit units from ND-COURSES per semester.
 - Creates separate sheets for each semester in the mastersheet.
 - Computes total scores with CA (20%) and OBJ+EXAM (80%) weighting, normalized from raw scores.
 - Computes remarks and derived metrics (CU Passed, TCPE, GPA).
 - Applies formatting: green for pass (≥50), bold red for fail (<50).
 - Includes summary and fails per course per semester.
"""

from openpyxl.cell import MergedCell
import os
import re
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import platform

BASE_DIR = "/mnt/c/Users/MTECH COMPUTERS/Documents/PROCESS_RESULT/EXAMS_INTERNAL"
ND_COURSES_DIR = os.path.join(BASE_DIR, "ND-COURSES")
DEFAULT_PASS_THRESHOLD = 50.0
TIMESTAMP_FMT = "%Y-%m-%d_%H:%M:%S"

def normalize_path(path):
    path = os.path.expanduser(path)
    path = os.path.normpath(path)
    if platform.system() == "Windows":
        path = path.replace("/", "\\")
    return os.path.abspath(path)

def normalize_course_name(name):
    return re.sub(r'\s+', ' ', str(name).strip().lower()).replace('coomunication', 'communication')

def load_course_data():
    course_file = os.path.join(ND_COURSES_DIR, "course-code-creditUnit.xlsx")
    print(f"Attempting to load: {course_file}")
    print(f"Directory contents: {os.listdir(ND_COURSES_DIR)}")
    if not os.path.exists(course_file):
        raise FileNotFoundError(f"Course file not found: {course_file}")
    
    try:
        xl = pd.ExcelFile(course_file)
        semester_course_maps = {}
        semester_credit_units = {}

        for sheet_name in xl.sheet_names:
            print(f"Processing sheet: {sheet_name}")
            df = pd.read_excel(course_file, sheet_name=sheet_name, engine='openpyxl', header=0)
            print(f"Raw data shape for {sheet_name}: {df.shape}")
            print(f"Raw data head for {sheet_name}:\n{df.head()}")
            print(f"Raw data tail for {sheet_name}:\n{df.tail()}")
            print(f"Column names for {sheet_name}: {df.columns.tolist()}")
            
            df.columns = [col.strip() for col in df.columns]
            print(f"Normalized column names for {sheet_name}: {df.columns.tolist()}")
            
            expected_cols = ['COURSE CODE', 'COURSE TITLE', 'CU']
            if not all(col in df.columns for col in expected_cols):
                missing_cols = [col for col in expected_cols if col not in df.columns]
                print(f"Warning: Missing columns {missing_cols} in {sheet_name}. Skipping this sheet.")
                continue
            
            data_df = df.dropna(subset=['COURSE CODE', 'COURSE TITLE'])
            print(f"After dropping NaN in CODE/TITLE for {sheet_name}, shape: {data_df.shape}")
            if data_df.empty:
                print(f"Warning: No valid data remaining in {sheet_name} after dropping NaN in COURSE CODE or COURSE TITLE. Skipping this sheet.")
                continue
            print(f"Data after dropna for {sheet_name}:\n{data_df}")
            
            data_df = data_df[~data_df['COURSE CODE'].str.contains('TOTAL', case=False, na=False)]
            print(f"After excluding totals for {sheet_name}, shape: {data_df.shape}")
            if data_df.empty:
                print(f"Warning: No valid data remaining in {sheet_name} after excluding totals. Skipping this sheet.")
                continue
            print(f"Data after totals filter for {sheet_name}:\n{data_df}")
            
            invalid_cu_rows = data_df[~data_df['CU'].astype(str).str.replace('.', '').str.isdigit()]
            if not invalid_cu_rows.empty:
                print(f"Warning: The following rows in {sheet_name} have invalid credit units (e.g., '-'):")
                for index, row in invalid_cu_rows.iterrows():
                    print(f"  - {row['COURSE CODE']}: {row['COURSE TITLE']} (CU: {row['CU']})")
                try:
                    response = input(f"Would you like to assign a default credit unit value of 0 to these courses in {sheet_name}? (y/n): ").lower()
                    if response not in ['y', 'n']:
                        print("Invalid input. Defaulting to skipping invalid rows.")
                        response = 'n'
                except EOFError:
                    print("No input received. Defaulting to skipping invalid rows.")
                    response = 'n'
                if response == 'y':
                    data_df.loc[invalid_cu_rows.index, 'CU'] = 0
                    print(f"Assigned default credit unit value of 0 to invalid entries in {sheet_name}.")
                else:
                    data_df = data_df[data_df['CU'].astype(str).str.replace('.', '').str.isdigit()]
                    print(f"Skipped rows with invalid credit units in {sheet_name}.")
            
            data_df = data_df.dropna(subset=['CU'])
            if data_df.empty:
                print(f"Warning: No valid data remaining in {sheet_name} after dropping NaN in CU. Skipping this sheet.")
                continue
            sheet_codes = data_df['COURSE CODE'].str.strip().tolist()
            sheet_names = data_df['COURSE TITLE'].str.strip().tolist()
            sheet_credit_units = data_df['CU'].astype(float).astype(int).tolist()
            semester_course_maps[sheet_name] = dict(zip(sheet_names, sheet_codes))
            semester_credit_units[sheet_name] = dict(zip(sheet_codes, sheet_credit_units))
        
        if not semester_course_maps:
            raise ValueError("No valid course data found across all sheets.")
        print(f"Successfully processed sheets: {list(semester_course_maps.keys())}")
        return semester_course_maps, semester_credit_units
    except Exception as e:
        raise Exception(f"Error reading course file: {str(e)}")

def find_column_by_names(df, candidate_names):
    norm_map = {col: re.sub(r'\s+', ' ', str(col).strip().lower()) for col in df.columns}
    candidates = [re.sub(r'\s+', ' ', c.strip().lower()) for c in candidate_names]
    for cand in candidates:
        for col, ncol in norm_map.items():
            if ncol == cand:
                return col
    return None

def process_file(path, output_dir, ts, pass_threshold, semester_course_maps, semester_credit_units):
    fname = os.path.basename(path)
    print(f"\nProcessing: {fname}")
    print(f"Full path: {path}")

    try:
        xl = pd.ExcelFile(path)
        sheets = ['CA', 'OBJ', 'EXAM']
        dfs = {sheet: pd.read_excel(path, sheet_name=sheet, dtype=str) for sheet in sheets if sheet in xl.sheet_names}
        if not dfs:
            print(f"Skipping {fname}: No valid sheets found")
            return None
    except Exception as e:
        print(f"Error accessing {fname}: {e}")
        return None

    # Determine semester from file name and match with course data
    semester_pattern = re.compile(r'(FIRST|SECOND)-YEAR-(FIRST|SECOND|THIRD)-SEMESTER', re.IGNORECASE)
    match = semester_pattern.search(fname)
    if not match:
        print(f"Warning: Could not determine semester from file name {fname}. Using first available semester.")
        semester_key = list(semester_course_maps.keys())[0]
    else:
        semester_key = f"ND-{match.group(0).upper()}"  # Match exact sheet names
        if semester_key not in semester_course_maps:
            print(f"Warning: Semester {semester_key} not found in course data. Using first available semester.")
            semester_key = list(semester_course_maps.keys())[0]

    course_map = semester_course_maps[semester_key]
    credit_units = semester_credit_units[semester_key]

    for sheet, df in dfs.items():
        print(f"Columns in {sheet} sheet: {df.columns.tolist()}")

    reg_no_cols = {sheet: find_column_by_names(df, ["REG. No", "Reg No", "Registration Number", "Mat No", "Exam No", "Student ID", "Reg. Number", "ID"]) for sheet, df in dfs.items()}
    name_cols = {sheet: find_column_by_names(df, ["NAME", "Full Name", "Candidate Name"]) for sheet, df in dfs.items()}
    for sheet in dfs:
        if not reg_no_cols[sheet]:
            print(f"Warning: Missing REG. No column in sheet {sheet}. Using first column as fallback.")
            reg_no_cols[sheet] = df.columns[0] if df.columns else None
        if not name_cols[sheet]:
            print(f"Warning: Missing NAME column in sheet {sheet}. Using second column as fallback or skipping if unavailable.")
            name_cols[sheet] = df.columns[1] if len(df.columns) > 1 else None

    merged = None
    for sheet, df in dfs.items():
        reg_no_col = reg_no_cols[sheet]
        name_col = name_cols[sheet]
        if not reg_no_col:
            print(f"Skipping {sheet} sheet: No valid REG. No column found.")
            continue
        df["REG. No"] = df[reg_no_col].astype(str).str.strip()
        if name_col:
            df["NAME"] = df[name_col].astype(str).str.strip()
        else:
            df["NAME"] = pd.NA
        columns_to_drop = [c for c in [reg_no_col, name_col] if c and c not in ["REG. No", "NAME"]]
        df.drop(columns=columns_to_drop, errors="ignore", inplace=True)
        print(f"Columns after standardization in {sheet}: {df.columns.tolist()}")
        course_cols = [c for c in df.columns if c not in ["REG. No", "NAME"]]
        for col in course_cols:
            norm_col = normalize_course_name(col)
            course_code = next((code for name, code in course_map.items() if normalize_course_name(name) == norm_col), None)
            if course_code:
                print(f"Renaming '{col}' to '{course_code}_{sheet.upper()}'")
                df.rename(columns={col: f"{course_code}_{sheet.upper()}"}, inplace=True)
        current_df = df[["REG. No"] + (["NAME"] if "NAME" in df.columns else []) + [c for c in df.columns if c.endswith(f"_{sheet.upper()}")]]
        if merged is None:
            merged = current_df
        else:
            merged = merged.merge(current_df, on="REG. No", how="outer", suffixes=('', '_dup'))
            if "NAME_dup" in merged.columns:
                merged["NAME"] = merged["NAME"].combine_first(merged["NAME_dup"])
                merged.drop(columns=["NAME_dup"], inplace=True)

    if merged is None or merged.empty:
        print(f"No valid data after merging sheets in {fname}")
        return None

    mastersheet = merged[["REG. No", "NAME"]].copy()
    mastersheet["S/N"] = range(1, len(mastersheet) + 1)
    for code in course_map.values():
        ca_col = f"{code}_CA"
        obj_col = f"{code}_OBJ"
        exam_col = f"{code}_EXAM"
        print(f"Processing {code}: CA={ca_col}, OBJ={obj_col}, EXAM={exam_col}")
        print(f"CA data sample: {merged[ca_col].head().tolist() if ca_col in merged.columns else 'Not found'}")
        print(f"OBJ data sample: {merged[obj_col].head().tolist() if obj_col in merged.columns else 'Not found'}")
        print(f"EXAM data sample: {merged[exam_col].head().tolist() if exam_col in merged.columns else 'Not found'}")
        ca_score = pd.to_numeric(merged[ca_col], errors="coerce") if ca_col in merged.columns else pd.Series([0] * len(merged), index=merged.index)
        obj_score = pd.to_numeric(merged[obj_col], errors="coerce") if obj_col in merged.columns else pd.Series([0] * len(merged), index=merged.index)
        exam_score = pd.to_numeric(merged[exam_col], errors="coerce") if exam_col in merged.columns else pd.Series([0] * len(merged), index=merged.index)
        # Normalize scores to 100-point scale
        ca_score = (ca_score / 20) * 100  # CA out of 20
        obj_score = (obj_score / 20) * 100  # OBJ out of 20
        exam_score = (exam_score / 80) * 100  # EXAM out of 80
        ca_score = ca_score.fillna(0).clip(upper=100)
        obj_score = obj_score.fillna(0).clip(upper=100)
        exam_score = exam_score.fillna(0).clip(upper=100)
        # Debug output for verification
        print(f"Debug - {code}: Normalized CA={ca_score.head().tolist()}, OBJ={obj_score.head().tolist()}, EXAM={exam_score.head().tolist()}")
        total = (ca_score * 0.2) + (((obj_score + exam_score) / 2) * 0.8)
        mastersheet[code] = total.round(1).clip(upper=100)

    mastersheet["REMARKS"] = mastersheet[list(course_map.values())].apply(
        lambda row: "Passed" if all(row[c] >= pass_threshold for c in course_map.values() if c in row.index) else
        "Failed: " + ", ".join(sorted(c for c in course_map.values() if c in row.index and row[c] < pass_threshold)),
        axis=1
    )
    mastersheet["CU Passed"] = mastersheet[list(course_map.values())].apply(
        lambda row: sum(credit_units.get(c, 0) for c in course_map.values() if c in row.index and row[c] >= pass_threshold),
        axis=1
    )
    mastersheet["CU Failed"] = mastersheet[list(course_map.values())].apply(
        lambda row: sum(credit_units.get(c, 0) for c in course_map.values() if c in row.index and row[c] < pass_threshold),
        axis=1
    )
    mastersheet["TCPE"] = mastersheet[list(course_map.values())].apply(
        lambda row: sum((row[c] / 100 * credit_units.get(c, 0)) for c in course_map.values() if c in row.index),
        axis=1
    ).round(2)
    total_cu = sum(credit_units.values())
    mastersheet["GPA"] = (mastersheet["TCPE"] / total_cu).round(2)
    mastersheet["AVERAGE"] = mastersheet[list(course_map.values())].mean(axis=1).round(1)

    # Sort REMARKS: "Passed" first, then by number of failed courses, then alphabetically by failed courses
    def sort_key(remark):
        if remark == "Passed":
            return (0, "")  # Passed comes first
        else:
            failed_courses = remark.replace("Failed: ", "").split(", ")
            return (1, len(failed_courses), ",".join(sorted(failed_courses)))  # Sort by count, then alphabetically

    mastersheet = mastersheet.sort_values(by="REMARKS", key=lambda x: x.map(sort_key)).reset_index(drop=True)
    mastersheet["S/N"] = range(1, len(mastersheet) + 1)

    out_cols = ["S/N", "REG. No", "NAME"] + list(course_map.values()) + ["REMARKS", "CU Passed", "CU Failed", "TCPE", "GPA", "AVERAGE"]
    mastersheet = mastersheet[out_cols]

    # Save to a workbook with a sheet for this semester
    output_subdir = os.path.join(output_dir, f"ND_RESULT-{ts}")
    os.makedirs(output_subdir, exist_ok=True)
    out_xlsx = os.path.join(output_subdir, f"mastersheet_{ts}.xlsx")
    if not os.path.exists(out_xlsx):
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet
    else:
        wb = load_workbook(out_xlsx)

    if semester_key not in wb.sheetnames:
        ws = wb.create_sheet(title=semester_key)
    else:
        ws = wb[semester_key]

    # Clear existing content by resetting the sheet
    ws.delete_rows(1, ws.max_row)
    ws.delete_cols(1, ws.max_column)

    ws.insert_rows(1, 2)
    ws.merge_cells("A1:Q1")
    title_cell = ws["A1"]
    title_cell.value = "FCT COLLEGE OF NURSING SCIENCES, GWAGWALADA-ABUJA"
    title_cell.font = Font(bold=True, size=16, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    border = Border(left=Side(style="medium"), right=Side(style="medium"), top=Side(style="medium"), bottom=Side(style="medium"))
    title_cell.border = border

    ws.merge_cells("A2:Q2")
    subtitle_cell = ws["A2"]
    subtitle_cell.value = f"2024/2025 SESSION  NATIONAL DIPLOMA I {semester_key.replace('ND-', '').upper().replace(' ', '')} EXAMINATIONS RESULT, OCTOBER 11, 2025"
    subtitle_cell.font = Font(bold=True, size=14, color="000000")
    subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Course names and CU with line breaks for long names and individual row height adjustment
    course_names = []
    for name in course_map.keys():
        if len(name.split()) > 3:  # Split long names into lines
            words = name.split()
            lines = [words[i:i + 2] for i in range(0, len(words), 2)]
            formatted_name = "\n".join(" ".join(line) for line in lines)
            course_names.append(formatted_name)
        else:
            course_names.append(name)
    ws.append([""] * 3 + course_names + [""] * 5)  # Row 3: Course names
    ws.append([""] * 3 + [credit_units[code] for code in course_map.values()] + [""] * 5)  # Row 4: CU
    for i, cell in enumerate(ws[3][3:3+len(course_map)], start=3):  # Row 3: Course names
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.font = Font(bold=True, size=12)
        # Set row height for each cell based on its content
        lines = str(cell.value).count("\n") + 1 if cell.value else 1
        ws.row_dimensions[3].height = max(ws.row_dimensions[3].height or 15, lines * 15)  # Adjust height per cell's lines
    for cell in ws[4][3:3+len(course_map)]:  # Row 4: CU
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(bold=True, size=12, color="000000")  # Dark bold text
        cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Light gray

    headers = ["S/N", "EXAMS NUMBER", "NAME"] + list(course_map.values()) + ["REMARKS", "CU Passed", "CU Failed", "TCPE", "GPA", "AVERAGE"]
    ws.append(headers)
    for cell in ws[5]:
        cell.font = Font(bold=True, size=12, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")  # Professional blue
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    for index, row in mastersheet.iterrows():
        ws.append(row.tolist())

    ws.freeze_panes = "A6"

    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in r:
            cell.border = thin_border
            if cell.row >= 6:  # Data rows
                col_idx = cell.column
                if col_idx in [2, 3] + [i for i in range(4, 13) if headers[i-1] in ["GNS102", "NUS111", "NUS112", "NUS113", "NUS114", "NUR111", "NUS115", "NUS116"]] + [12, 13, 14, 15, 16, 17]:
                    cell.alignment = Alignment(horizontal="left")
                else:
                    cell.alignment = Alignment(horizontal="center")

    for col_idx, col in enumerate(mastersheet.columns, 1):
        if col in course_map.values():
            for row_idx in range(6, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                try:
                    val = float(cell.value) if cell.value is not None else 0
                    if val >= pass_threshold:
                        green_intensity = min(1.0, val / 100)
                        green_value = int(128 + 127 * (1 - green_intensity))
                        start_color = f"{green_value:02x}FF00"
                        cell.fill = PatternFill(start_color=start_color, end_color="C6EFCE", fill_type="solid")  # Light green
                        cell.font = Font(color="006100")  # Green text
                    else:
                        cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # No fill
                        cell.font = Font(color="FF0000", bold=True)  # Bold red text
                except (ValueError, TypeError):
                    continue
        elif col in ["S/N"]:
            for row_idx in range(6, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(horizontal="center")

    fails_per_course = mastersheet[list(course_map.values())].apply(lambda x: (x < pass_threshold).sum()).tolist()
    ws.append([""] * 3 + fails_per_course + [""] * 5)
    ws.append(["FAILS PER COURSE:"] + [""] * 2 + fails_per_course + [""] * 5)
    for cell in ws[ws.max_row-1]:
        if 4 <= cell.column < 4 + len(course_map):
            cell.fill = PatternFill(start_color="F0E68C", end_color="F0E68C", fill_type="solid")  # Light yellow
            cell.font = Font(bold=True, size=12)
            cell.alignment = Alignment(horizontal="center")
    for cell in ws[ws.max_row]:
        if cell.column == 1:
            cell.font = Font(bold=True, size=12)
            cell.alignment = Alignment(horizontal="center")

    total_students = len(mastersheet)
    passed_all = len(mastersheet[mastersheet["REMARKS"] == "Passed"])
    failed_gpa_above_2 = len(mastersheet[(mastersheet["GPA"] >= 2.00) & (mastersheet["REMARKS"] != "Passed") & (mastersheet["CU Passed"] >= 0.45 * total_cu)])
    failed_gpa_below_2 = len(mastersheet[(mastersheet["GPA"] < 2.00) & (mastersheet["REMARKS"] != "Passed") & (mastersheet["CU Passed"] >= 0.45 * total_cu)])
    failed_over_45 = len(mastersheet[mastersheet["CU Failed"] > 0.45 * total_cu])
    ws.append([])
    ws.append(["SUMMARY"])
    ws.append([f"A total of {total_students} students registered and sat for the Examination"])
    ws.append([f"A total of {passed_all} students passed in all courses registered and are to proceed to {semester_key.replace('ND-', '').replace('First', 'Second').replace('Second', 'Next') if 'First' in semester_key else 'Next'} Semester, ND I"])
    ws.append([f"A total of {failed_gpa_above_2} students with Grade Point Average (GPA) of 2.00 and above failed various courses, but passed at least 45% of the total registered credit units, and are to carry these courses over to the next session."])
    ws.append([f"A total of {failed_gpa_below_2} students with Grade Point Average (GPA) below 2.00 failed various courses, but passed at least 45% of the total registered credit units, and are placed on Probation, to carry these courses over to the next session."])
    ws.append([f"A total of {failed_over_45} students failed in more than 45% of their registered credit units in various courses and have been advised to withdraw"])
    ws.append(["The above decisions are in line with the provisions of the General Information Section of the NMCN/NBTE Examinations Regulations (Pg 4) adopted by the College."])
    ws.append([])
    ws.append(["________________________", "", "", "", "", "", "", "", "", "", "", "", "", "", "________________________"])
    ws.append(["Mrs. Abini Hauwa", "", "", "", "", "", "", "", "", "", "", "", "", "", "Mrs. Olukemi Ogunleye"])
    ws.append(["Head of Exams", "", "", "", "", "", "", "", "", "", "", "", "", "", "Chairman, ND/HND Program C'tee"])
    for row in ws[ws.max_row-6:ws.max_row+1]:
        for cell in row:
            if cell.row == ws.max_row - 6:
                cell.font = Font(bold=True, size=14)
                cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Light gray
            elif cell.row in [ws.max_row - 5, ws.max_row - 4, ws.max_row - 3, ws.max_row - 2]:
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(horizontal="left")
            elif cell.row in [ws.max_row - 1, ws.max_row]:
                if cell.column in [1, 15]:  # Signatories
                    cell.font = Font(italic=True, size=12)
                    cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

    # Auto-fit column widths with minimum adjustments
    for col_idx, col in enumerate(ws.columns, 1):
        max_length = 0
        column = get_column_letter(col_idx)
        if column == "A":  # S/N column
            max_length = max(len(str(ws.cell(row=5, column=col_idx).value)), 3)
            adjusted_width = min(max_length + 2, 6)
        else:
            for cell in col:
                if not isinstance(cell, MergedCell) and cell.value is not None:
                    try:
                        cell_length = len(str(cell.value).replace("\n", ""))  # Count without newlines for width
                        max_length = max(max_length, cell_length)
                    except:
                        pass
            adjusted_width = max(max_length + 2, 10)  # Minimum width of 10 for readability
        ws.column_dimensions[column].width = adjusted_width

    wb.save(out_xlsx)
    print(f"Saved mastersheet: {os.path.basename(out_xlsx)}")

    return mastersheet

def main():
    print("Starting ND Examination Results Processing...")
    ts = datetime.now().strftime(TIMESTAMP_FMT)

    semester_course_maps = {}
    semester_credit_units = {}

    try:
        semester_course_maps, semester_credit_units = load_course_data()
    except Exception as e:
        print(f"❌ Failed to load course data: {e}")
        print("Attempting to proceed with partial course data if available.")

    if not semester_course_maps or not semester_credit_units:
        print("Warning: No valid course data loaded. Processing aborted.")
        return

    for year_dir in ["ND-2024", "ND-2025"]:
        raw_dir = normalize_path(os.path.join(BASE_DIR, year_dir, "RAW_RESULTS"))
        clean_dir = normalize_path(os.path.join(BASE_DIR, year_dir, "CLEAN_RESULTS"))
        os.makedirs(raw_dir, exist_ok=True)
        os.makedirs(clean_dir, exist_ok=True)

        files = [f for f in os.listdir(raw_dir) if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
        if not files:
            print(f"No raw files found in {raw_dir}")
            continue

        for f in files:
            process_file(os.path.join(raw_dir, f), clean_dir, ts, DEFAULT_PASS_THRESHOLD, semester_course_maps, semester_credit_units)

    print("\n✅ Processing completed successfully.")

if __name__ == "__main__":
    main()