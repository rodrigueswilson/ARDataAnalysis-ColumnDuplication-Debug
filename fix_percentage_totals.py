#!/usr/bin/env python3
"""
Fix Percentage Calculation in Data Cleaning Sheet Totals
========================================================

This script fixes the Filter Application Rate (%) calculation in the Data Cleaning sheet totals
to properly include all excluded files (school_outliers + non_school_normal + non_school_outliers)
instead of just non_school_outliers.
"""

import re
import os
import sys

def fix_percentage_calculation():
    """Fix the percentage calculation in the Data Cleaning sheet implementation."""
    file_path = os.path.join('report_generator', 'sheet_creators', '__init__.py')
    
    if not os.path.exists(file_path):
        print(f"ERROR: File '{file_path}' not found.")
        return False
    
    # Read the file content
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Make a backup of the original file
    backup_path = f"{file_path}.bak"
    with open(backup_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Backup created: {backup_path}")
    
    # Find and fix Table 1 total row calculation
    table1_pattern = re.compile(r"""(\s+# Add formatted total row at the specific position
\s+total_row_values = \{
\s+\d+: totals\['total_files'\],
\s+\d+: totals\['school_outliers'\],
\s+\d+: totals\['non_school_normal'\],
\s+\d+: totals\['school_normal'\],
\s+\d+: totals\['non_school_outliers'\],
\s+\d+: totals\['school_normal'\],
\s+\d+: )totals\['non_school_outliers'\] / totals\['total_files'\]( if totals\['total_files'\] > 0 else 0,
\s+\d+: totals\['school_normal'\] / totals\['total_files'\]( if totals\['total_files'\] > 0 else 0
\s+\})""", re.MULTILINE)
    
    # Fix Table 1 calculation
    replacement = r"""\1(totals['school_outliers'] + totals['non_school_normal'] + totals['non_school_outliers']) / totals['total_files']\2\3"""
    fixed_content = re.sub(table1_pattern, replacement, content)
    
    # Find and fix Table 2 total row calculation (similar pattern but likely in a different location)
    table2_pattern = re.compile(r"""(\s+# Add formatted total row at the specific position
\s+total_row_values = \{
\s+\d+: totals\['total_files'\],
\s+\d+: totals\['school_outliers'\],
\s+\d+: totals\['non_school_normal'\],
\s+\d+: totals\['(?:non_school_normal|school_normal)'\],
\s+\d+: totals\['non_school_outliers'\],
\s+\d+: totals\['school_normal'\],
\s+\d+: )totals\['non_school_outliers'\] / totals\['total_files'\]( if totals\['total_files'\] > 0 else 0,
\s+\d+: totals\['school_normal'\] / totals\['total_files'\]( if totals\['total_files'\] > 0 else 0
\s+\})""", re.MULTILINE)
    
    # Fix Table 2 calculation
    fixed_content = re.sub(table2_pattern, replacement, fixed_content)
    
    # Also insert a calculation for total_excluded as a new variable before each total_row_values
    excluded_calc = "            # Calculate total excluded files for correct Filter Application Rate\n            total_excluded = totals['school_outliers'] + totals['non_school_normal'] + totals['non_school_outliers']\n            \n"
    
    # Replace for all occurrences of total_row_values
    fixed_content = re.sub(r"(\s+# Add formatted total row at the specific position\n)", r"\1" + excluded_calc, fixed_content)
    
    # Count the number of replacements made
    original_matches = len(re.findall(r"totals\['non_school_outliers'\] / totals\['total_files'\]", content))
    fixed_matches = len(re.findall(r"\(totals\['school_outliers'\] \+ totals\['non_school_normal'\] \+ totals\['non_school_outliers'\]\) / totals\['total_files'\]", fixed_content))
    
    # If we made replacements, write back the fixed content
    if original_matches > 0 and fixed_matches > 0:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(fixed_content)
        print(f"SUCCESS: Fixed {fixed_matches} percentage calculations in {file_path}")
        print("Filter Application Rate (%) now correctly includes all excluded files.")
        return True
    else:
        print("ERROR: Could not find or fix the percentage calculation pattern.")
        return False

if __name__ == "__main__":
    print("=" * 80)
    print("FIXING PERCENTAGE CALCULATION IN DATA CLEANING SHEET")
    print("=" * 80)
    
    success = fix_percentage_calculation()
    
    if success:
        print("\nThe fix has been applied successfully.")
        print("Run generate_report.py to generate a new report with corrected percentages.")
    else:
        print("\nThe fix could not be applied automatically.")
        print("Please make the following changes manually:")
        print("1. Find the total_row_values calculation in create_data_cleaning_sheet")
        print("2. Replace totals['non_school_outliers'] with (totals['school_outliers'] + totals['non_school_normal'] + totals['non_school_outliers'])")
        print("3. This ensures Filter Application Rate includes all excluded files")
    
    print("=" * 80)
