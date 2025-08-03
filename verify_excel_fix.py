#!/usr/bin/env python3
"""
Verify Excel Fix

This script opens the most recent Excel report and checks if our fix worked:
1. Are the left-aligned numbers gone?
2. Does the total match the database total (9731)?
3. Are early September dates included properly?
"""

import sys
import os
import pandas as pd
from pathlib import Path
import glob

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def find_latest_excel_report():
    """Find the most recent Excel report file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    
    if not files:
        print("❌ No Excel reports found")
        return None
    
    # Sort by modification time, get the latest
    latest_file = max(files, key=os.path.getmtime)
    print(f"📊 Latest Excel report: {latest_file}")
    return latest_file

def verify_daily_counts_sheet(excel_file):
    """Verify the Daily Counts (ACF_PACF) sheet."""
    print("\nVERIFYING DAILY COUNTS (ACF_PACF) SHEET")
    print("=" * 50)
    
    try:
        # Read the specific sheet
        df = pd.read_excel(excel_file, sheet_name='Daily Counts (ACF_PACF)')
        
        print(f"Sheet dimensions: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        
        # Check if we have the expected structure
        if '_id' in df.columns and 'Total_Files' in df.columns:
            # Remove any NaN rows
            df_clean = df.dropna(subset=['_id'])
            
            print(f"Clean data rows: {len(df_clean)}")
            
            # Calculate total files
            total_files = df_clean['Total_Files'].sum()
            print(f"Total files in Excel: {total_files}")
            
            # Check if we have 333 rows (full collection days)
            if len(df_clean) == 333:
                print("✅ Has 333 rows (full collection day range)")
            else:
                print(f"❌ Has {len(df_clean)} rows (expected 333)")
            
            # Check if total matches expected
            if total_files == 9731:
                print("✅ Total files match database (9731)")
            else:
                print(f"❌ Total files don't match (expected 9731, got {total_files})")
            
            # Check early September dates
            early_sept_dates = [
                "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
                "2021-09-20", "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
            ]
            
            print("\nEarly September dates check:")
            early_sept_count = 0
            found_dates = 0
            
            for date_str in early_sept_dates:
                matching_rows = df_clean[df_clean['_id'] == date_str]
                if not matching_rows.empty:
                    count = matching_rows.iloc[0]['Total_Files']
                    early_sept_count += count
                    found_dates += 1
                    print(f"  ✅ {date_str}: {count} files")
                else:
                    print(f"  ❌ {date_str}: Not found")
            
            print(f"\nEarly September summary:")
            print(f"  Dates found: {found_dates}/{len(early_sept_dates)}")
            print(f"  Total files: {early_sept_count}")
            
            if found_dates == len(early_sept_dates) and early_sept_count == 462:
                print("✅ All early September dates included with correct counts")
                return True
            else:
                print("❌ Early September dates missing or incorrect counts")
                return False
                
        else:
            print("❌ Sheet doesn't have expected columns (_id, Total_Files)")
            print(f"Available columns: {list(df.columns)}")
            return False
            
    except Exception as e:
        print(f"❌ Error reading Excel sheet: {e}")
        return False

def check_for_left_aligned_numbers(excel_file):
    """Check if there are any left-aligned numbers at the top."""
    print("\nCHECKING FOR LEFT-ALIGNED NUMBERS")
    print("=" * 50)
    
    try:
        # Read the sheet with header=None to see raw structure
        df_raw = pd.read_excel(excel_file, sheet_name='Daily Counts (ACF_PACF)', header=None)
        
        print(f"Raw sheet dimensions: {df_raw.shape}")
        
        # Look at the first 20 rows to see if there are left-aligned numbers
        print("First 20 rows of raw data:")
        for i in range(min(20, len(df_raw))):
            row_data = df_raw.iloc[i].tolist()
            # Convert to string and truncate for display
            row_str = str(row_data)[:100] + "..." if len(str(row_data)) > 100 else str(row_data)
            print(f"  Row {i}: {row_str}")
        
        # Check if the first column has any numeric values in the first few rows
        # (which would indicate left-aligned numbers)
        first_col = df_raw.iloc[:, 0]
        numeric_in_first_col = []
        
        for i, val in enumerate(first_col.head(10)):
            if pd.notna(val) and str(val).isdigit():
                numeric_in_first_col.append((i, val))
        
        if numeric_in_first_col:
            print(f"\n❌ Found potential left-aligned numbers:")
            for row_idx, val in numeric_in_first_col:
                print(f"  Row {row_idx}: {val}")
            return False
        else:
            print("\n✅ No left-aligned numbers detected in first column")
            return True
            
    except Exception as e:
        print(f"❌ Error checking for left-aligned numbers: {e}")
        return False

def main():
    print("EXCEL FIX VERIFICATION")
    print("=" * 60)
    
    # Find the latest Excel report
    excel_file = find_latest_excel_report()
    if not excel_file:
        return
    
    # Verify the Daily Counts sheet
    daily_counts_ok = verify_daily_counts_sheet(excel_file)
    
    # Check for left-aligned numbers
    no_left_aligned = check_for_left_aligned_numbers(excel_file)
    
    # Final assessment
    print("\n" + "=" * 60)
    print("VERIFICATION RESULTS")
    print("=" * 60)
    
    if daily_counts_ok and no_left_aligned:
        print("🎉 SUCCESS! Excel fix is working correctly:")
        print("  ✅ Daily Counts sheet has correct structure and totals")
        print("  ✅ No left-aligned numbers detected")
        print("  ✅ Early September dates are included")
        print("  ✅ Total files match database (9731)")
    else:
        print("❌ Excel fix is NOT working correctly:")
        if not daily_counts_ok:
            print("  ❌ Daily Counts sheet has issues")
        if not no_left_aligned:
            print("  ❌ Left-aligned numbers still present")
        
        print("\n🔧 NEXT STEPS:")
        print("  1. Check if zero-fill method is actually being called")
        print("  2. Verify that the correct implementation is being used")
        print("  3. Add more debug logging to trace the issue")

if __name__ == "__main__":
    main()
