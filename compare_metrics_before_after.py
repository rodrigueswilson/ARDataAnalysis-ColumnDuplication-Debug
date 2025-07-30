#!/usr/bin/env python3
"""
Simple comparison of statistical metrics before and after the collection days fix.
"""

import pandas as pd
import glob

def compare_metrics():
    """Compare statistical metrics between reports to verify the fix."""
    print("=== STATISTICAL METRICS COMPARISON ===")
    
    # Find the reports before and after the fix
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    files.sort()
    
    if len(files) < 2:
        print("Need at least 2 reports to compare")
        return
    
    # Compare the last few reports
    print("Available reports:")
    for i, file in enumerate(files[-5:], 1):
        print(f"  {i}. {file}")
    
    # Use specific reports for comparison
    # Before fix: AR_Analysis_Report_20250726_222111.xlsx (had correct Total Collection Days)
    # After fix: AR_Analysis_Report_20250726_223327.xlsx (should have corrected statistical metrics)
    
    before_file = "AR_Analysis_Report_20250726_222111.xlsx"
    after_file = "AR_Analysis_Report_20250726_223327.xlsx"
    
    if before_file not in files or after_file not in files:
        print(f"Required files not found:")
        print(f"  Before: {before_file} - {'Found' if before_file in files else 'Missing'}")
        print(f"  After: {after_file} - {'Found' if after_file in files else 'Missing'}")
        return
    
    try:
        # Read both reports
        df_before = pd.read_excel(before_file, sheet_name='Summary Statistics')
        df_after = pd.read_excel(after_file, sheet_name='Summary Statistics')
        
        print(f"\nComparing metrics:")
        print(f"  Before fix: {before_file}")
        print(f"  After fix:  {after_file}")
        
        # Compare key metrics
        metrics_to_compare = [
            'Total Collection Days',
            'Mean Files per Day',
            'Median Files per Day', 
            'Standard Deviation',
            'Min Files per Day',
            'Max Files per Day',
            'Range (Max - Min)'
        ]
        
        print(f"\n{'Metric':<25} | {'Before 2021-22':<12} | {'After 2021-22':<12} | {'Before 2022-23':<12} | {'After 2022-23':<12} | {'Changed?'}")
        print("-" * 100)
        
        changes_detected = False
        
        for metric in metrics_to_compare:
            before_row = df_before[df_before['Metric'] == metric]
            after_row = df_after[df_after['Metric'] == metric]
            
            if not before_row.empty and not after_row.empty:
                before_2122 = before_row['2021-2022'].iloc[0]
                after_2122 = after_row['2021-2022'].iloc[0]
                before_2223 = before_row['2022-2023'].iloc[0]
                after_2223 = after_row['2022-2023'].iloc[0]
                
                # Check if values changed
                changed_2122 = abs(float(before_2122) - float(after_2122)) > 0.1
                changed_2223 = abs(float(before_2223) - float(after_2223)) > 0.1
                changed = changed_2122 or changed_2223
                
                if changed:
                    changes_detected = True
                
                change_indicator = "YES" if changed else "No"
                print(f"{metric:<25} | {before_2122:<12} | {after_2122:<12} | {before_2223:<12} | {after_2223:<12} | {change_indicator}")
        
        print(f"\n=== ANALYSIS ===")
        if changes_detected:
            print("✅ Changes detected in statistical metrics!")
            print("   The collection days filtering appears to be working.")
        else:
            print("⚠️  No changes detected in statistical metrics.")
            print("   This could mean:")
            print("   1. The filtering logic isn't working as expected, OR")
            print("   2. There are no non-collection days in the actual data being processed, OR") 
            print("   3. The non-collection days don't significantly affect the statistics")
        
        # Check specifically for the key indicator - Min Files per Day
        min_before = df_before[df_before['Metric'] == 'Min Files per Day']
        min_after = df_after[df_after['Metric'] == 'Min Files per Day']
        
        if not min_before.empty and not min_after.empty:
            min_val_before = min_before['Total'].iloc[0]
            min_val_after = min_after['Total'].iloc[0]
            
            print(f"\nKey indicator - Min Files per Day (Total):")
            print(f"  Before: {min_val_before}")
            print(f"  After:  {min_val_after}")
            
            if float(min_val_before) == 0.0 and float(min_val_after) == 0.0:
                print("  ⚠️  Still showing 0.0 - filtering may not be working as expected")
            elif float(min_val_after) > float(min_val_before):
                print("  ✅ Minimum increased - filtering is working!")
            else:
                print("  ❓ No change in minimum value")
    
    except Exception as e:
        print(f"Error comparing metrics: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    compare_metrics()
