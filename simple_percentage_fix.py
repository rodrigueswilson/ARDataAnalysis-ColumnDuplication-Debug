#!/usr/bin/env python3
"""
Simple Fix for Percentage Calculation in Data Cleaning Sheet
===========================================================

This script fixes the Filter Application Rate (%) calculation in the Data Cleaning sheet
to properly include all excluded files instead of just non_school_outliers.
"""

import os
import shutil

def fix_percentage_calculation():
    """Fix the percentage calculation in the Data Cleaning sheet implementation."""
    file_path = os.path.join('report_generator', 'sheet_creators', '__init__.py')
    
    if not os.path.exists(file_path):
        print(f"ERROR: File '{file_path}' not found.")
        return False
    
    # Make a backup of the original file
    backup_path = f"{file_path}.bak"
    shutil.copy2(file_path, backup_path)
    print(f"Backup created: {backup_path}")
    
    # Read the file content
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    fixed_lines = []
    fix_count = 0
    total_excluded_added = False
    
    # Process the file line by line
    for i, line in enumerate(lines):
        # Check for the line with non_school_outliers calculation for percentages
        if "totals['non_school_outliers'] / totals['total_files']" in line:
            # Replace with the correct calculation using all excluded files
            new_line = line.replace(
                "totals['non_school_outliers'] / totals['total_files']",
                "(totals['school_outliers'] + totals['non_school_normal'] + totals['non_school_outliers']) / totals['total_files']"
            )
            fixed_lines.append(new_line)
            fix_count += 1
            print(f"Fixed line {i+1}: {line.strip()} -> {new_line.strip()}")
            
            # If we haven't added the total_excluded calculation before this point,
            # scan backward to find a good insertion point
            if not total_excluded_added:
                # Look for the previous line with "total_row_values"
                for j in range(i-1, max(0, i-5), -1):
                    if "total_row_values" in lines[j]:
                        # Insert the total_excluded calculation before total_row_values
                        insert_point = j
                        indent = len(lines[j]) - len(lines[j].lstrip())
                        indent_str = ' ' * indent
                        
                        # Insert the new calculation lines
                        fixed_lines.insert(
                            -1,  # Insert before the last added line
                            f"{indent_str}# Calculate total excluded files for correct Filter Application Rate\n"
                        )
                        fixed_lines.insert(
                            -1,
                            f"{indent_str}total_excluded = totals['school_outliers'] + totals['non_school_normal'] + totals['non_school_outliers']\n"
                        )
                        fixed_lines.insert(-1, f"{indent_str}\n")
                        
                        total_excluded_added = True
                        print(f"Added total_excluded calculation before line {j+1}")
                        break
        else:
            fixed_lines.append(line)
    
    # Write the fixed content back
    if fix_count > 0:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.writelines(fixed_lines)
        print(f"SUCCESS: Fixed {fix_count} percentage calculations in {file_path}")
        print("Filter Application Rate (%) now correctly includes all excluded files.")
        return True
    else:
        print("ERROR: Could not find any percentage calculations to fix.")
        return False

def run_verification():
    """Generate a test report to verify the fix."""
    import subprocess
    print("\nRunning quick verification...")
    result = subprocess.run(['python', 'generate_report.py'], 
                           capture_output=True, text=True)
    
    if result.returncode == 0:
        print("✅ Report generation succeeded!")
        print("The fix has been successfully applied.")
        return True
    else:
        print("❌ Report generation failed.")
        print(result.stderr)
        return False

if __name__ == "__main__":
    print("=" * 80)
    print("FIXING PERCENTAGE CALCULATION IN DATA CLEANING SHEET")
    print("=" * 80)
    
    success = fix_percentage_calculation()
    
    if success:
        print("\nThe fix has been applied successfully.")
        verification = input("\nWould you like to verify the fix by generating a test report? (y/n): ")
        if verification.lower() == 'y':
            run_verification()
    else:
        print("\nThe fix could not be applied automatically.")
        print("Please make the following changes manually:")
        print("1. Find the total_row_values calculation in create_data_cleaning_sheet")
        print("2. Replace totals['non_school_outliers'] with (totals['school_outliers'] + totals['non_school_normal'] + totals['non_school_outliers'])")
        print("3. This ensures Filter Application Rate includes all excluded files")
    
    print("=" * 80)
