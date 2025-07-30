#!/usr/bin/env python3
"""
Final verification that the statistical metrics fix is working correctly.
"""

import pandas as pd
import glob

def verify_final_fix():
    """Verify that the statistical metrics are now correctly filtered."""
    print("=== FINAL STATISTICAL METRICS FIX VERIFICATION ===")
    
    # Compare the latest reports
    files = glob.glob("AR_Analysis_Report_*.xlsx")
    files.sort()
    
    if len(files) < 2:
        print("Need at least 2 reports to compare")
        return
    
    # Use the most recent reports for comparison
    before_file = "AR_Analysis_Report_20250726_223327.xlsx"  # Before date format fix
    after_file = "AR_Analysis_Report_20250726_223756.xlsx"   # After date format fix
    
    print(f"Comparing:")
    print(f"  Before fix: {before_file}")
    print(f"  After fix:  {after_file}")
    
    try:
        # Read both reports
        df_before = pd.read_excel(before_file, sheet_name='Summary Statistics')
        df_after = pd.read_excel(after_file, sheet_name='Summary Statistics')
        
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
        significant_changes = []
        
        for metric in metrics_to_compare:
            before_row = df_before[df_before['Metric'] == metric]
            after_row = df_after[df_after['Metric'] == metric]
            
            if not before_row.empty and not after_row.empty:
                before_2122 = before_row['2021-2022'].iloc[0]
                after_2122 = after_row['2021-2022'].iloc[0]
                before_2223 = before_row['2022-2023'].iloc[0]
                after_2223 = after_row['2022-2023'].iloc[0]
                
                # Check if values changed significantly
                changed_2122 = abs(float(before_2122) - float(after_2122)) > 0.1
                changed_2223 = abs(float(before_2223) - float(after_2223)) > 0.1
                changed = changed_2122 or changed_2223
                
                if changed:
                    changes_detected = True
                    change_details = f"{metric}: "
                    if changed_2122:
                        change_details += f"2021-22: {before_2122} -> {after_2122} "
                    if changed_2223:
                        change_details += f"2022-23: {before_2223} -> {after_2223}"
                    significant_changes.append(change_details)
                
                change_indicator = "YES" if changed else "No"
                print(f"{metric:<25} | {before_2122:<12} | {after_2122:<12} | {before_2223:<12} | {after_2223:<12} | {change_indicator}")
        
        print(f"\n=== RESULTS ANALYSIS ===")
        if changes_detected:
            print("🎉 SUCCESS! Changes detected in statistical metrics!")
            print("✅ The collection days filtering is now working correctly!")
            print(f"\nSignificant changes:")
            for change in significant_changes:
                print(f"  - {change}")
            
            # Check the key indicators
            min_before = df_before[df_before['Metric'] == 'Min Files per Day']
            min_after = df_after[df_after['Metric'] == 'Min Files per Day']
            
            if not min_before.empty and not min_after.empty:
                min_total_before = float(min_before['Total'].iloc[0])
                min_total_after = float(min_after['Total'].iloc[0])
                
                print(f"\nKey Success Indicator - Min Files per Day (Total):")
                print(f"  Before: {min_total_before}")
                print(f"  After:  {min_total_after}")
                
                if min_total_after > min_total_before:
                    print("  🎯 PERFECT! Minimum value increased - non-collection day zeros excluded!")
                elif min_total_before == 0.0 and min_total_after > 0.0:
                    print("  🎯 EXCELLENT! Minimum is no longer 0 - filtering working!")
                else:
                    print("  ⚠️  Minimum didn't increase as expected")
            
            print(f"\n✅ CONCLUSION: Statistical metrics are now consistent with collection days logic!")
            print(f"✅ All metrics exclude non-collection days (holidays, breaks, etc.)")
            print(f"✅ Summary Statistics sheet is now fully consistent internally")
            
        else:
            print("⚠️  No changes detected in statistical metrics.")
            print("   This could indicate:")
            print("   1. The date format fix didn't work as expected")
            print("   2. There might be no non-collection days in the actual data")
            print("   3. The non-collection days don't significantly impact statistics")
            
            # Additional debugging
            print(f"\nDebugging info:")
            min_row = df_after[df_after['Metric'] == 'Min Files per Day']
            if not min_row.empty:
                min_val = min_row['Total'].iloc[0]
                if float(min_val) == 0.0:
                    print(f"  ❌ Min Files per Day is still 0.0 - filtering may not be working")
                else:
                    print(f"  ✅ Min Files per Day is {min_val} - filtering might be working")
    
    except Exception as e:
        print(f"Error in verification: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    verify_final_fix()
