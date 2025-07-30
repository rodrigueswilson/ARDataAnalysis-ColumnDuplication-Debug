#!/usr/bin/env python3
"""
Examine the Summary Statistics sheet structure to find collection days data
and analyze for inconsistencies.
"""

import pandas as pd
import glob
from ar_utils import calculate_collection_days_for_period

def find_latest_report():
    """Find the most recent AR Analysis Report Excel file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort()
    return files[-1]

def examine_summary_metrics():
    """Examine the Summary Statistics sheet structure and find collection days data."""
    latest_report = find_latest_report()
    if not latest_report:
        print("No report files found!")
        return
    
    print(f"Examining Summary Statistics in: {latest_report}")
    print("\n=== SUMMARY STATISTICS STRUCTURE ===")
    
    try:
        # Read the Summary Statistics sheet
        df = pd.read_excel(latest_report, sheet_name='Summary Statistics')
        
        print(f"Sheet structure: {len(df)} rows x {len(df.columns)} columns")
        print(f"Columns: {list(df.columns)}")
        print()
        
        # Display all data to understand structure
        print("=== FULL SHEET DATA ===")
        print(df.to_string(index=False))
        print()
        
        # Look for collection days related metrics
        print("=== SEARCHING FOR COLLECTION DAYS METRICS ===")
        collection_metrics = []
        
        if 'Metric' in df.columns:
            for idx, row in df.iterrows():
                metric = str(row['Metric']).lower()
                if 'collection' in metric and 'days' in metric:
                    collection_metrics.append(idx)
                    print(f"Found collection days metric at row {idx}: '{row['Metric']}'")
        
        if not collection_metrics:
            print("No direct 'collection days' metrics found.")
            print("Searching for related metrics...")
            
            # Look for any day-related metrics
            day_metrics = []
            if 'Metric' in df.columns:
                for idx, row in df.iterrows():
                    metric = str(row['Metric']).lower()
                    if 'days' in metric or 'day' in metric:
                        day_metrics.append(idx)
                        print(f"Found day-related metric at row {idx}: '{row['Metric']}'")
            
            if day_metrics:
                print(f"\nAnalyzing {len(day_metrics)} day-related metrics:")
                for idx in day_metrics:
                    row = df.iloc[idx]
                    metric_name = row['Metric']
                    val_2122 = row.get('2021-2022', 'N/A')
                    val_2223 = row.get('2022-2023', 'N/A')
                    total_val = row.get('Total', 'N/A')
                    
                    print(f"\n  {metric_name}:")
                    print(f"    2021-2022: {val_2122}")
                    print(f"    2022-2023: {val_2223}")
                    print(f"    Total: {total_val}")
                    
                    # Check if this could be collection days
                    if 'collection' in metric_name.lower() or 'school' in metric_name.lower():
                        print(f"    *** This might be collection days data ***")
                        analyze_collection_days_metric(metric_name, val_2122, val_2223, total_val)
        else:
            print(f"\nAnalyzing {len(collection_metrics)} collection days metrics:")
            for idx in collection_metrics:
                row = df.iloc[idx]
                metric_name = row['Metric']
                val_2122 = row.get('2021-2022', 'N/A')
                val_2223 = row.get('2022-2023', 'N/A')
                total_val = row.get('Total', 'N/A')
                
                print(f"\n  {metric_name}:")
                print(f"    2021-2022: {val_2122}")
                print(f"    2022-2023: {val_2223}")
                print(f"    Total: {total_val}")
                
                analyze_collection_days_metric(metric_name, val_2122, val_2223, total_val)
    
    except Exception as e:
        print(f"Error examining Summary Statistics: {e}")
        import traceback
        traceback.print_exc()

def analyze_collection_days_metric(metric_name, val_2122, val_2223, total_val):
    """Analyze a potential collection days metric for consistency."""
    print(f"    === CONSISTENCY ANALYSIS ===")
    
    try:
        # Convert to numeric if possible
        num_2122 = float(val_2122) if pd.notna(val_2122) and str(val_2122) != 'N/A' else None
        num_2223 = float(val_2223) if pd.notna(val_2223) and str(val_2223) != 'N/A' else None
        num_total = float(total_val) if pd.notna(total_val) and str(total_val) != 'N/A' else None
        
        # Calculate expected values
        expected_2122 = 34 + 59 + 63  # SY 21-22: P1 + P2 + P3
        expected_2223 = 56 + 58 + 50  # SY 22-23: P1 + P2 + P3
        expected_total = expected_2122 + expected_2223
        
        print(f"    Expected 2021-2022: {expected_2122} days")
        print(f"    Expected 2022-2023: {expected_2223} days")
        print(f"    Expected Total: {expected_total} days")
        
        # Check consistency
        issues = []
        
        if num_2122 is not None:
            diff_2122 = num_2122 - expected_2122
            if abs(diff_2122) > 0.1:
                issues.append(f"2021-2022: Off by {diff_2122} days")
                print(f"    ⚠️  2021-2022: Reported={num_2122}, Expected={expected_2122}, Diff={diff_2122}")
            else:
                print(f"    ✅ 2021-2022: Consistent ({num_2122} days)")
        
        if num_2223 is not None:
            diff_2223 = num_2223 - expected_2223
            if abs(diff_2223) > 0.1:
                issues.append(f"2022-2023: Off by {diff_2223} days")
                print(f"    ⚠️  2022-2023: Reported={num_2223}, Expected={expected_2223}, Diff={diff_2223}")
            else:
                print(f"    ✅ 2022-2023: Consistent ({num_2223} days)")
        
        if num_total is not None:
            diff_total = num_total - expected_total
            if abs(diff_total) > 0.1:
                issues.append(f"Total: Off by {diff_total} days")
                print(f"    ⚠️  Total: Reported={num_total}, Expected={expected_total}, Diff={diff_total}")
            else:
                print(f"    ✅ Total: Consistent ({num_total} days)")
        
        # Check if reported total matches sum of reported years
        if num_2122 is not None and num_2223 is not None and num_total is not None:
            calculated_total = num_2122 + num_2223
            if abs(num_total - calculated_total) > 0.1:
                issues.append(f"Total doesn't match sum: {num_total} ≠ {calculated_total}")
                print(f"    ⚠️  Total calculation: {num_total} ≠ {num_2122} + {num_2223} = {calculated_total}")
            else:
                print(f"    ✅ Total matches sum: {num_total} = {calculated_total}")
        
        if issues:
            print(f"    🚨 INCONSISTENCIES FOUND:")
            for issue in issues:
                print(f"      - {issue}")
        else:
            print(f"    ✅ All values are consistent!")
    
    except Exception as e:
        print(f"    ❌ Error in analysis: {e}")

if __name__ == "__main__":
    examine_summary_metrics()
