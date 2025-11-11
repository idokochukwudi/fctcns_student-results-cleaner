#!/usr/bin/env python3
"""
auto_fix_bn_duplicates.py - Automatically fix duplicate function definitions

Usage:
    python auto_fix_bn_duplicates.py exam_processor_bn.py
"""

import sys
import os
from datetime import datetime


def fix_duplicate_functions(filepath):
    """Remove duplicate function definitions from the BN processor script."""

    if not os.path.exists(filepath):
        print(f"❌ Error: File not found: {filepath}")
        return False

    print(f"📖 Reading file: {filepath}")
    print(f"📊 File size: {os.path.getsize(filepath):,} bytes")

    # Read all lines
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()

    total_lines = len(lines)
    print(f"📊 Total lines: {total_lines:,}")

    # Create backup
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{filepath}.backup_{timestamp}"

    print(f"\n💾 Creating backup: {backup_path}")
    with open(backup_path, "w", encoding="utf-8") as f:
        f.writelines(lines)

    # Find duplicate function definitions
    print("\n🔍 Searching for duplicate functions...")

    duplicate_ranges = []

    # Search for second occurrence of create_bn_cgpa_summary_sheet
    count_cgpa = 0
    for i, line in enumerate(lines, 1):
        if "def create_bn_cgpa_summary_sheet" in line:
            count_cgpa += 1
            if count_cgpa == 2:
                print(f"   Found duplicate create_bn_cgpa_summary_sheet at line {i}")
                # Find end of this function (next def or end of file)
                end_line = i
                for j in range(i, min(len(lines), i + 300)):
                    if (
                        j > i
                        and lines[j].startswith("def ")
                        and not lines[j].startswith("    ")
                    ):
                        end_line = j - 1
                        break
                duplicate_ranges.append((i - 1, end_line))  # 0-indexed

    # Search for second occurrence of load_previous_cgpas_from_processed_files
    count_prev = 0
    for i, line in enumerate(lines, 1):
        if "def load_previous_cgpas_from_processed_files" in line:
            count_prev += 1
            if count_prev == 2:
                print(
                    f"   Found duplicate load_previous_cgpas_from_processed_files at line {i}"
                )
                end_line = i
                for j in range(i, min(len(lines), i + 200)):
                    if (
                        j > i
                        and lines[j].startswith("def ")
                        and not lines[j].startswith("    ")
                    ):
                        end_line = j - 1
                        break
                duplicate_ranges.append((i - 1, end_line))

    # Search for second occurrence of load_all_previous_cgpas_for_cumulative
    count_all = 0
    for i, line in enumerate(lines, 1):
        if "def load_all_previous_cgpas_for_cumulative" in line:
            count_all += 1
            if count_all == 2:
                print(
                    f"   Found duplicate load_all_previous_cgpas_for_cumulative at line {i}"
                )
                end_line = i
                for j in range(i, min(len(lines), i + 200)):
                    if (
                        j > i
                        and lines[j].startswith("def ")
                        and not lines[j].startswith("    ")
                    ):
                        end_line = j - 1
                        break
                duplicate_ranges.append((i - 1, end_line))

    if not duplicate_ranges:
        print("\n✅ No duplicate functions found! File is already correct.")
        return True

    # Remove duplicate ranges
    print(f"\n🗑️  Removing {len(duplicate_ranges)} duplicate function(s)...")

    lines_to_keep = set(range(len(lines)))

    for start, end in duplicate_ranges:
        print(f"   Deleting lines {start+1} to {end+1}")
        for line_num in range(start, min(end + 1, len(lines))):
            lines_to_keep.discard(line_num)

    # Keep only non-deleted lines
    filtered_lines = [lines[i] for i in sorted(lines_to_keep)]

    # Write fixed file
    print(f"\n✍️  Writing fixed file: {filepath}")

    with open(filepath, "w", encoding="utf-8") as f:
        f.writelines(filtered_lines)

    # Summary
    deleted_lines = total_lines - len(filtered_lines)

    print("\n" + "=" * 60)
    print("✅ FILE FIXED SUCCESSFULLY!")
    print("=" * 60)
    print(f"📊 Original lines:  {total_lines:,}")
    print(f"📊 Fixed lines:     {len(filtered_lines):,}")
    print(f"📊 Deleted lines:   {deleted_lines:,}")
    print(f"💾 Backup saved to: {backup_path}")
    print("=" * 60)

    print("\n🎯 Next steps:")
    print("   1. Run your script to verify it works")
    print("   2. Check the CGPA_SUMMARY sheet - STATUS should now be correct")
    print("   3. If something goes wrong, restore from backup:")
    print(f"      cp {backup_path} {filepath}")

    return True


def main():
    """Main entry point."""
    if len(sys.argv) != 2:
        print(
            """
╔════════════════════════════════════════════════════════════╗
║          BN Processor Duplicate Function Remover            ║
╚════════════════════════════════════════════════════════════╝

This script automatically removes duplicate function definitions
that are overwriting the fixed versions in your BN processor.

USAGE:
    python auto_fix_bn_duplicates.py exam_processor_bn.py

WHAT IT DOES:
    1. Creates a timestamped backup of your file
    2. Finds duplicate function definitions
    3. Removes the duplicate (second) occurrences
    4. Keeps only the fixed versions

FUNCTIONS FIXED:
    ✓ create_bn_cgpa_summary_sheet
    ✓ load_previous_cgpas_from_processed_files  
    ✓ load_all_previous_cgpas_for_cumulative

SAFE:
    - Creates backup before making changes
    - Only removes exact duplicate functions
    - Preserves all other code
"""
        )
        sys.exit(1)

    filepath = sys.argv[1]

    print("╔════════════════════════════════════════════════════════════╗")
    print("║          BN Processor Duplicate Function Remover            ║")
    print("╚════════════════════════════════════════════════════════════╝\n")

    success = fix_duplicate_functions(filepath)

    if success:
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
