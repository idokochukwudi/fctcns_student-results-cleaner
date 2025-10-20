#!/usr/bin/env python3
"""
caosce_result.py

CAOSCE cleaning script (updated with robust FULL NAME extraction, proper S/N sorting, and timestamped output folders).

Railway + Local compatible with new folder structure.
"""

import os
import re
from datetime import datetime
import math
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------
# Environment Detection & Directory Configuration
# ---------------------------
IS_RAILWAY = os.getenv('RAILWAY_ENVIRONMENT') is not None

if IS_RAILWAY:
    # Railway: BASE_DIR is set to /app/EXAMS_INTERNAL
    BASE_DIR = os.getenv('BASE_DIR', '/app/EXAMS_INTERNAL')
    print("🚂 Running on Railway")
    print(f"   BASE_DIR: {BASE_DIR}")
else:
    # Local: use home directory structure
    BASE_DIR = os.path.join(os.path.expanduser('~'), 'student_result_cleaner', 'EXAMS_INTERNAL')
    print("💻 Running locally")
    print(f"   BASE_DIR: {BASE_DIR}")

# UPDATED: Directories now under CAOSCE_RESULT
DEFAULT_BASE_DIR = os.path.join(BASE_DIR, "CAOSCE_RESULT")
DEFAULT_RAW_DIR = os.path.join(DEFAULT_BASE_DIR, "RAW_CAOSCE_RESULT")
DEFAULT_CLEAN_DIR = os.path.join(DEFAULT_BASE_DIR, "CLEAN_CAOSCE_RESULT")

TIMESTAMP_FMT = "%Y-%m-%d_%H:%M:%S"
OUTPUT_BASENAME = "CAOSCE_CLEANED"

STATION_COLUMN_MAP = {
    "procedure_station_one": "PS1_Score/",
    "procedure_station_three": "PS3_Score/",
    "procedure_station_five": "PS5_Score/",
    "question_station_two": "QS2_Score/",
    "question_station_four": "QS4_Score/",
    "question_station_six": "QS6_Score/",
    "viva": "VIVA/10",
}

UNWANTED_COL_PATTERNS = [
    r"phone",
    r"department",
    r"city",
    r"town",
    r"state",
    r"started on",
    r"completed",
    r"time taken",
    r"q\.\s*\d+",
    r"q\s*\.\s*\d+",
]

NO_SCORE_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
NO_SCORE_FONT = Font(bold=True, color="9C0006")
HEADER_FONT = Font(bold=True)


# ---------------------------
# Helpers
# ---------------------------
def find_first_col(df, candidates):
    cols = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        cand_norm = cand.strip().lower()
        if cand_norm in cols:
            return cols[cand_norm]
        for c in df.columns:
            if cand_norm in str(c).strip().lower():
                return c
    return None


def find_username_col(df):
    return find_first_col(
        df,
        [
            "username",
            "user name",
            "exam no",
            "registration no",
            "reg no",
            "mat no",
            "matno",
            "regnum",
        ],
    )


def find_fullname_col(df):
    return find_first_col(
        df, ["full name", "user full name", "name", "candidate name", "student name"]
    )


def find_viva_score_col(df):
    return find_first_col(df, ["enter student's score", "enter student score", "score"])


def find_grade_column(df):
    for c in df.columns:
        cn = str(c).strip().lower()
        if (
            cn.startswith("grade/")
            or cn == "grade"
            or cn == "total"
            or "grade" in cn
            or "total" in cn
        ):
            return c
    return None


def sanitize_exam_no(v):
    if pd.isna(v):
        return ""
    s = str(v).strip()
    s = re.sub(r"\.0+$", "", s)
    return s


def numeric_safe(v):
    try:
        if pd.isna(v):
            return None
        v2 = str(v).strip()
        if v2 == "":
            return None
        v2 = v2.replace(",", "")
        return float(v2)
    except Exception:
        return None


def auto_column_width(ws, min_width=8, max_width=60):
    for i, col in enumerate(ws.columns, 1):
        max_len = max(
            (len(str(cell.value)) for cell in col if cell.value is not None), default=0
        )
        ws.column_dimensions[get_column_letter(i)].width = min(
            max_width, max(min_width, max_len + 2)
        )


# ---------------------------
# Main processing
# ---------------------------
def process_files():
    print("Starting CAOSCE Results Cleaning...\n")

    # Use the configured directories
    RAW_DIR = DEFAULT_RAW_DIR
    BASE_CLEAN_DIR = DEFAULT_CLEAN_DIR

    # Create timestamped output directory
    ts = datetime.now().strftime(TIMESTAMP_FMT)
    output_dir = os.path.join(BASE_CLEAN_DIR, f"CAOSCE_RESULT_{ts}")
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(RAW_DIR, exist_ok=True)

    files = [
        f for f in os.listdir(RAW_DIR) if f.lower().endswith((".xlsx", ".xls", ".csv"))
    ]
    if not files:
        print(f"No raw files found in {RAW_DIR}")
        return

    results = {}

    for fname in sorted(files):
        path = os.path.join(RAW_DIR, fname)
        lower = fname.lower()

        # Determine station type
        station_key = None
        if "procedure" in lower:
            if "one" in lower or "station 1" in lower:
                station_key = "procedure_station_one"
            elif "three" in lower or "station 3" in lower:
                station_key = "procedure_station_three"
            elif "five" in lower or "station 5" in lower:
                station_key = "procedure_station_five"
        if "question" in lower:
            if "two" in lower or "station 2" in lower:
                station_key = "question_station_two"
            elif "four" in lower or "station 4" in lower:
                station_key = "question_station_four"
            elif "six" in lower or "station 6" in lower:
                station_key = "question_station_six"
        if "viva" in lower:
            station_key = "viva"

        # Load file
        try:
            if fname.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str)
            else:
                df = pd.read_excel(path, dtype=str)
        except Exception as e:
            print(f"Error reading {fname}: {e}")
            continue

        df.rename(columns=lambda c: str(c).strip(), inplace=True)

        username_col = find_username_col(df)
        fullname_col = find_fullname_col(df)
        grade_col = find_grade_column(df)
        viva_col = find_viva_score_col(df)

        # Drop unwanted columns
        for pattern in UNWANTED_COL_PATTERNS:
            df.drop(
                columns=[
                    c for c in df.columns if re.search(pattern, str(c), flags=re.I)
                ],
                inplace=True,
                errors="ignore",
            )

        rows_added = 0
        for _, row in df.iterrows():
            exam_no = sanitize_exam_no(row.get(username_col) if username_col else None)
            if not exam_no:
                continue

            if exam_no not in results:
                results[exam_no] = {
                    "EXAM NO.": exam_no,
                    "FULL NAME": None,
                    "PS1_Score/": None,
                    "PS3_Score/": None,
                    "PS5_Score/": None,
                    "QS2_Score/": None,
                    "QS4_Score/": None,
                    "QS6_Score/": None,
                    "VIVA/10": None,
                }

            # Robust FULL NAME extraction
            full_name_candidate = None
            if fullname_col and pd.notna(row.get(fullname_col)):
                val = str(row.get(fullname_col)).strip()
                if val and re.search(r"[A-Za-z]{2,}", val):
                    full_name_candidate = val

            if not full_name_candidate:
                for c in df.columns:
                    if c == username_col:
                        continue
                    v = str(row.get(c, "")).strip()
                    if (
                        v
                        and re.search(r"[A-Za-z]{2,}", v)
                        and not re.fullmatch(r"\d+", v)
                    ):
                        full_name_candidate = v
                        break

            if full_name_candidate and not results[exam_no]["FULL NAME"]:
                results[exam_no]["FULL NAME"] = full_name_candidate

            # Score assignment
            score_val = None
            if station_key == "viva":
                if viva_col:
                    score_val = numeric_safe(row.get(viva_col))
                elif grade_col:
                    score_val = numeric_safe(row.get(grade_col))
                if score_val is not None:
                    results[exam_no]["VIVA/10"] = round(score_val, 2)
            else:
                if grade_col:
                    score_val = numeric_safe(row.get(grade_col))
                if score_val is not None:
                    out_col = STATION_COLUMN_MAP.get(station_key)
                    results[exam_no][out_col] = round(score_val, 2)

            rows_added += 1

        print(f"Processed {fname} ({rows_added} rows read)")

    # Final DataFrame
    final_cols = [
        "EXAM NO.",
        "FULL NAME",
        "PS1_Score/",
        "PS3_Score/",
        "PS5_Score/",
        "QS2_Score/",
        "QS4_Score/",
        "QS6_Score/",
        "VIVA/10",
    ]
    df_out = pd.DataFrame(list(results.values()))

    df_out["FULL NAME"] = df_out.groupby("EXAM NO.")["FULL NAME"].transform(
        lambda x: x.ffill().bfill()
    )

    df_out["__exam_num_sort"] = pd.to_numeric(df_out["EXAM NO."], errors="coerce")
    df_out.sort_values(
        by=["__exam_num_sort", "EXAM NO."],
        ascending=[True, True],
        inplace=True,
        na_position="last",
    )
    df_out.drop(columns="__exam_num_sort", inplace=True)
    df_out.reset_index(drop=True, inplace=True)

    df_out.insert(0, "S/N", range(1, len(df_out) + 1))

    for col in final_cols[2:]:
        df_out[col] = df_out[col].apply(
            lambda v: (
                "NO SCORE"
                if v is None or (isinstance(v, float) and math.isnan(v))
                else v
            )
        )

    # Save to timestamped directory
    out_csv = os.path.join(output_dir, f"{OUTPUT_BASENAME}_{ts}.csv")
    out_xlsx = os.path.join(output_dir, f"{OUTPUT_BASENAME}_{ts}.xlsx")
    df_out.to_csv(out_csv, index=False)
    df_out.to_excel(out_xlsx, index=False, engine="openpyxl")
    print(f"Saved processed file: {os.path.basename(out_csv)}")
    print(f"Saved processed file: {os.path.basename(out_xlsx)}")

    # Format Excel
    try:
        wb = load_workbook(out_xlsx)
        ws = wb.active
        for cell in ws[1]:
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"

        thin = Side(border_style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for row_cells in ws.iter_rows(
            min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
        ):
            for cell in row_cells:
                cell.border = border

        score_cols_idx = [
            i + 1 for i, c in enumerate(df_out.columns) if c in final_cols[2:]
        ]
        for r in range(2, ws.max_row + 1):
            for col_idx in score_cols_idx:
                cell = ws.cell(row=r, column=col_idx)
                val = cell.value
                if val == "NO SCORE" or val is None:
                    cell.value = "NO SCORE"
                    cell.fill = NO_SCORE_FILL
                    cell.font = NO_SCORE_FONT
                    cell.alignment = Alignment(horizontal="center")
                else:
                    try:
                        cell.value = float(val)
                        cell.number_format = "0.00"
                        cell.alignment = Alignment(horizontal="center")
                    except Exception:
                        cell.fill = NO_SCORE_FILL
                        cell.font = NO_SCORE_FONT

        auto_column_width(ws)
        wb.save(out_xlsx)
        print(f"Saved processed file with formatting: {os.path.basename(out_xlsx)}")
    except Exception as e:
        print(f"Error formatting XLSX: {e}")

    print("\nProcessing completed successfully.")
    print(f"Files saved in: {output_dir}")


if __name__ == "__main__":
    process_files()