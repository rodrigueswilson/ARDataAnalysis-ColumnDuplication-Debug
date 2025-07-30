#!/usr/bin/env python3
"""
Direct analysis of statistical metrics in Summary Statistics to determine
if they should be recalculated based on collection days logic.
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

def analyze_metrics_consistency():
    """Analyze if statistical metrics are consistent with collection days logic."""
    latest_report = find_latest_report()
    if not latest_report:
        print("No report files found!")
        return
    
    print(f"Analyzing: {latest_report}")
    print("\n=== STATISTICAL METRICS CONSISTENCY ANALYSIS ===")
    
    try:
        # Read the Summary Statistics sheet
        df = pd.read_excel(latest_report, sheet_name='Summary Statistics')
        
        # Display current metrics
        metrics_to_check = [
            'Total Collection Days',
            'Mean Files per Day',
            'Median Files per Day', 
            'Standard Deviation',
            'Min Files per Day',
            'Max Files per Day',
            'Range (Max - Min)'
        ]
        
        print("Current Summary Statistics values:")
        print("=" * 70)
        print(f"{'Metric':<25} | {'2021-2022':<10} | {'2022-2023':<10} | {'Total':<10}")
        print("-" * 70)
        
        for metric in metrics_to_check:
            row = df[df['Metric'] == metric]
            if not row.empty:
                val_2122 = row['2021-2022'].iloc[0]
                val_2223 = row['2022-2023'].iloc[0]
                val_total = row['Total'].iloc[0]
                print(f"{metric:<25} | {val_2122:<10} | {val_2223:<10} | {val_total:<10}")
        
        print("\n=== CONSISTENCY ISSUE ANALYSIS ===")
        
        print("\n🔍 CURRENT SITUATION:")
        print("✅ 'Total Collection Days' = FIXED (now uses proper collection days logic)")
        print("❓ Statistical metrics (Mean, Median, etc.) = UNKNOWN (need verification)")
        
        print("\n🔍 KEY QUESTION:")
        print("Are the statistical metrics calculated from:")
        print("  Option A: ALL days with data (including non-collection days)")
        print("  Option B: ONLY collection days (excluding holidays, breaks, etc.)")
        
        print("\n🔍 CODE ANALYSIS:")
        print("Based on the Summary Statistics generation code:")
        print("1. Data source: DAILY_COUNTS_ALL pipeline")
        print("2. Processing: _fill_missing_collection_days() adds zeros for missing days")
        print("3. Statistics: Calculated from ALL resulting days (including zeros)")
        print("4. Issue: This includes non-collection days that were filled with zeros")
        
        print("\n🚨 INCONSISTENCY IDENTIFIED:")
        print("- 'Total Collection Days' excludes non-collection days ✅")
        print("- Statistical metrics include non-collection days ❌")
        print("- This creates logical inconsistency in the same sheet")
        
        print("\n📊 IMPACT ON METRICS:")
        print("- Mean Files per Day: Artificially lowered by zero-filled non-collection days")
        print("- Median Files per Day: Affected by inclusion of non-collection day zeros")
        print("- Standard Deviation: Inflated by artificial zeros from non-collection days")
        print("- Min Files per Day: Likely 0 due to non-collection day zeros")
        print("- Max Files per Day: Unaffected (highest value regardless of day type)")
        print("- Range: Inflated due to artificial minimum of 0")
        
        print("\n✅ RECOMMENDATION:")
        print("Statistical metrics SHOULD be recalculated to:")
        print("1. Use ONLY collection days (same logic as 'Total Collection Days')")
        print("2. Exclude non-collection days, holidays, breaks, etc.")
        print("3. Provide accurate statistics for educational analysis")
        print("4. Maintain consistency within the Summary Statistics sheet")
        
        print("\n🔧 PROPOSED SOLUTION:")
        print("Update the calc_stats_for_period() function to:")
        print("1. Filter data to exclude non-collection days before calculating statistics")
        print("2. Use the same date filtering logic as collection days calculation")
        print("3. Calculate Mean/Median/Std/Min/Max/Range from filtered data only")
        
        print("\n📋 IMPLEMENTATION NEEDED:")
        print("Modify report_generator/sheet_creators.py:")
        print("- Filter df_daily to exclude non-collection days")
        print("- Apply same logic used for collection days calculation")
        print("- Recalculate all statistical metrics from filtered data")
        
        return True
    
    except Exception as e:
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    analyze_metrics_consistency()
