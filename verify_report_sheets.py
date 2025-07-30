#!/usr/bin/env python3
"""
Verify what sheets are present in the latest report and check for any issues.
"""

import pandas as pd
import glob

def find_latest_report():
    """Find the most recent AR Analysis Report Excel file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort()
    return files[-1]

def verify_report_sheets():
    """Verify what sheets are present in the latest report."""
    latest_report = find_latest_report()
    if not latest_report:
        print("No report files found!")
        return
    
    print(f"Verifying sheets in: {latest_report}")
    print("\n=== SHEET VERIFICATION ===")
    
    try:
        # Get all sheet names
        excel_file = pd.ExcelFile(latest_report)
        sheet_names = excel_file.sheet_names
        
        print(f"Total sheets found: {len(sheet_names)}")
        print("\nSheet list:")
        for i, sheet_name in enumerate(sheet_names, 1):
            print(f"  {i:2d}. {sheet_name}")
        
        # Check specifically for Summary Statistics
        if 'Summary Statistics' in sheet_names:
            print(f"\n✅ Summary Statistics sheet found!")
            
            # Try to read it
            df = pd.read_excel(latest_report, sheet_name='Summary Statistics')
            print(f"   Dimensions: {len(df)} rows x {len(df.columns)} columns")
            print(f"   Columns: {list(df.columns)}")
            
            # Check for our metrics
            metrics_to_check = [
                'Total Collection Days',
                'Mean Files per Day',
                'Median Files per Day', 
                'Standard Deviation',
                'Min Files per Day',
                'Max Files per Day',
                'Range (Max - Min)'
            ]
            
            print(f"\n   Checking for key metrics:")
            for metric in metrics_to_check:
                if metric in df['Metric'].values:
                    print(f"     ✅ {metric}")
                else:
                    print(f"     ❌ {metric} - MISSING")
        else:
            print(f"\n❌ Summary Statistics sheet NOT found!")
            print(f"   Available sheets: {sheet_names}")
        
    except Exception as e:
        print(f"Error verifying sheets: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    verify_report_sheets()
