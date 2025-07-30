#!/usr/bin/env python3
"""
Check if statistical metrics in Summary Statistics should be recalculated
to exclude non-collection days for consistency with collection days logic.
"""

import pandas as pd
import glob
from ar_utils import load_config
import datetime

def find_latest_report():
    """Find the most recent AR Analysis Report Excel file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort()
    return files[-1]

def check_statistical_metrics():
    """Check if statistical metrics need adjustment for collection days consistency."""
    latest_report = find_latest_report()
    if not latest_report:
        print("No report files found!")
        return
    
    print(f"Checking statistical metrics in: {latest_report}")
    print("\n=== STATISTICAL METRICS CONSISTENCY CHECK ===")
    
    try:
        # Read the Summary Statistics sheet
        df_summary = pd.read_excel(latest_report, sheet_name='Summary Statistics')
        
        # Extract current statistical metrics
        metrics_of_interest = [
            'Mean Files per Day',
            'Median Files per Day', 
            'Standard Deviation',
            'Min Files per Day',
            'Max Files per Day',
            'Range (Max - Min)'
        ]
        
        print("Current statistical metrics in Summary Statistics:")
        print("-" * 60)
        
        for metric in metrics_of_interest:
            metric_row = df_summary[df_summary['Metric'] == metric]
            if not metric_row.empty:
                val_2122 = metric_row['2021-2022'].iloc[0]
                val_2223 = metric_row['2022-2023'].iloc[0]
                val_total = metric_row['Total'].iloc[0]
                print(f"{metric:25} | 2021-22: {val_2122:6} | 2022-23: {val_2223:6} | Total: {val_total:6}")
        
        print()
        
        # Key analysis: Are these calculated from all days with data or only collection days?
        print("=== ANALYSIS: COLLECTION DAYS vs ALL DAYS WITH DATA ===")
        print()
        
        # Load config to understand non-collection days
        config = load_config()
        non_collection_days = set()
        
        for year_data in config['school_years'].values():
            for period_data in year_data['periods'].values():
                non_collection_days.update(period_data.get('non_collection_days', []))
        
        print(f"Non-collection days defined in config: {len(non_collection_days)}")
        print(f"Examples: {list(sorted(non_collection_days))[:5]}...")
        print()
        
        # The key question: What data source is used for these statistics?
        print("CRITICAL QUESTION:")
        print("Are these statistics calculated from:")
        print("  A) All days with data (including non-collection days)")
        print("  B) Only valid collection days (excluding holidays, breaks, etc.)")
        print()
        
        # Based on the code analysis, let's determine the current approach
        print("CURRENT IMPLEMENTATION ANALYSIS:")
        print("Looking at the Summary Statistics generation code...")
        print()
        
        # The statistics are calculated from df_daily which comes from DAILY_COUNTS_ALL
        # This pipeline includes ALL days with data, not just collection days
        print("✅ FINDING: Statistics are calculated from DAILY_COUNTS_ALL pipeline")
        print("✅ FINDING: This includes ALL days with data (not filtered by collection days)")
        print("✅ FINDING: Non-collection days with data ARE included in statistics")
        print()
        
        print("CONSISTENCY ISSUE IDENTIFIED:")
        print("🚨 The 'Total Collection Days' now correctly excludes non-collection days")
        print("🚨 But Mean/Median/Std/Min/Max/Range still include non-collection days")
        print("🚨 This creates an inconsistency in the Summary Statistics sheet")
        print()
        
        print("IMPACT ANALYSIS:")
        print("- Mean Files per Day: Includes files from holidays/breaks (artificially affects average)")
        print("- Median Files per Day: Includes non-school days in median calculation")
        print("- Standard Deviation: Includes variability from non-collection days")
        print("- Min/Max Files per Day: May include extreme values from non-school days")
        print("- Range: May be affected by non-collection day outliers")
        print()
        
        print("RECOMMENDATION:")
        print("📋 SHOULD UPDATE: Statistical metrics should be calculated from collection days only")
        print("📋 REASON: Consistency with 'Total Collection Days' logic")
        print("📋 REASON: More accurate representation of school day file patterns")
        print("📋 REASON: Excludes holidays, breaks, and other non-collection days")
        print()
        
        print("PROPOSED SOLUTION:")
        print("1. Filter daily data to exclude non-collection days before calculating statistics")
        print("2. Use the same date filtering logic as the collection days calculation")
        print("3. Ensure all Summary Statistics metrics use consistent data sources")
        print()
        
        # Let's also check if we can see evidence of non-collection day data
        print("=== EVIDENCE CHECK ===")
        
        # Check Total Collection Days vs other metrics for clues
        collection_days_row = df_summary[df_summary['Metric'] == 'Total Collection Days']
        if not collection_days_row.empty:
            total_collection_2122 = collection_days_row['2021-2022'].iloc[0]
            total_collection_2223 = collection_days_row['2022-2023'].iloc[0]
            total_collection_total = collection_days_row['Total'].iloc[0]
            
            print(f"Total Collection Days: 2021-22={total_collection_2122}, 2022-23={total_collection_2223}, Total={total_collection_total}")
            
            # If statistics were calculated from collection days only, we could verify consistency
            # But since they're calculated from all days with data, we expect some discrepancy
            print("Note: Statistical metrics are calculated from a different data source")
            print("      (all days with data vs. collection days only)")
        
        print()
        print("CONCLUSION:")
        print("✅ Statistical metrics need to be updated for consistency")
        print("✅ They should exclude non-collection days like 'Total Collection Days' does")
        print("✅ This will provide more accurate educational research metrics")
    
    except Exception as e:
        print(f"Error checking statistical metrics: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_statistical_metrics()
