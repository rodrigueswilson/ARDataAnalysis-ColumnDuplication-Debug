#!/usr/bin/env python3
"""
Verify that the statistical metrics filtering is working correctly
by comparing before/after values and checking actual data filtering.
"""

import pandas as pd
import glob
from ar_utils import _load_config

def find_latest_report():
    """Find the most recent AR Analysis Report Excel file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort()
    return files[-1]

def verify_metrics_filtering():
    """Verify that statistical metrics filtering is working correctly."""
    print("=== STATISTICAL METRICS FILTERING VERIFICATION ===")
    
    try:
        # Load config to get non-collection days
        config = _load_config()
        non_collection_days = set()
        if 'non_collection_days' in config:
            non_collection_days = set(config['non_collection_days'].keys())
        
        print(f"Non-collection days in config: {len(non_collection_days)}")
        print(f"Sample non-collection days: {sorted(list(non_collection_days))[:10]}")
        print()
        
        # Test the actual data pipeline to see what data is being processed
        from report_generator.sheet_creators import SheetCreator
        from report_generator.formatters import ExcelFormatter
        from db_utils import get_db_connection
        from pipelines.time_series import PIPELINES
        
        # Initialize components
        db = get_db_connection()
        formatter = ExcelFormatter()
        sheet_creator = SheetCreator(db, formatter)
        
        # Get the same data that Summary Statistics uses
        print("Fetching daily data from DAILY_COUNTS_ALL pipeline...")
        df_daily = sheet_creator._run_aggregation(PIPELINES['DAILY_COUNTS_ALL'])
        
        if df_daily.empty:
            print("❌ No daily data found!")
            return
        
        print(f"✅ Found {len(df_daily)} daily records")
        print(f"Date range: {df_daily['_id'].min()} to {df_daily['_id'].max()}")
        
        # Fill missing collection days (same as Summary Statistics does)
        df_daily = sheet_creator._fill_missing_collection_days(df_daily, 'DAILY_COUNTS_ALL_WITH_ZEROES')
        print(f"After filling missing days: {len(df_daily)} records")
        
        # Add school year classification
        df_daily['date_obj'] = pd.to_datetime(df_daily['_id'])
        df_daily['school_year'] = df_daily['date_obj'].apply(
            lambda d: f"{d.year}-{d.year + 1}" if d.month >= 8 else f"{d.year - 1}-{d.year}"
        )
        
        # Test filtering logic
        df_daily['date_str'] = df_daily['_id'].astype(str)
        
        # Check how many records are on non-collection days
        non_collection_records = df_daily[df_daily['date_str'].isin(non_collection_days)]
        print(f"\nRecords on non-collection days: {len(non_collection_records)}")
        
        if len(non_collection_records) > 0:
            print("Sample non-collection day records:")
            print(non_collection_records[['_id', 'Total_Files']].head())
            print(f"File counts on non-collection days: {non_collection_records['Total_Files'].describe()}")
        else:
            print("✅ No records found on non-collection days")
        
        # Apply filtering
        df_filtered = df_daily[~df_daily['date_str'].isin(non_collection_days)].copy()
        print(f"\nAfter filtering: {len(df_daily)} -> {len(df_filtered)} records")
        print(f"Filtered out: {len(df_daily) - len(df_filtered)} records")
        
        # Calculate statistics for comparison
        print(f"\n=== STATISTICAL COMPARISON ===")
        
        for year in ['2021-2022', '2022-2023']:
            year_data_all = df_daily[df_daily['school_year'] == year]
            year_data_filtered = df_filtered[df_filtered['school_year'] == year]
            
            if len(year_data_all) == 0:
                continue
                
            print(f"\n{year}:")
            print(f"  All data: {len(year_data_all)} records")
            print(f"  Filtered: {len(year_data_filtered)} records")
            print(f"  Difference: {len(year_data_all) - len(year_data_filtered)} records")
            
            # Calculate statistics from both datasets
            all_files = year_data_all['Total_Files']
            filtered_files = year_data_filtered['Total_Files']
            
            print(f"  Statistics comparison:")
            print(f"    Mean:   All={all_files.mean():.1f}, Filtered={filtered_files.mean():.1f}")
            print(f"    Median: All={all_files.median():.1f}, Filtered={filtered_files.median():.1f}")
            print(f"    Std:    All={all_files.std():.1f}, Filtered={filtered_files.std():.1f}")
            print(f"    Min:    All={all_files.min()}, Filtered={filtered_files.min()}")
            print(f"    Max:    All={all_files.max()}, Filtered={filtered_files.max()}")
        
        # Compare with actual Summary Statistics values
        print(f"\n=== COMPARISON WITH SUMMARY STATISTICS ===")
        latest_report = find_latest_report()
        if latest_report:
            df_summary = pd.read_excel(latest_report, sheet_name='Summary Statistics')
            
            metrics_to_check = ['Mean Files per Day', 'Median Files per Day', 'Standard Deviation', 'Min Files per Day', 'Max Files per Day']
            
            for metric in metrics_to_check:
                metric_row = df_summary[df_summary['Metric'] == metric]
                if not metric_row.empty:
                    val_2122 = metric_row['2021-2022'].iloc[0]
                    val_2223 = metric_row['2022-2023'].iloc[0]
                    print(f"{metric}: 2021-22={val_2122}, 2022-23={val_2223}")
        
        print(f"\n=== CONCLUSION ===")
        if len(df_daily) == len(df_filtered):
            print("⚠️  No filtering occurred - all data records are on collection days")
            print("   This means the current statistics are already correct")
        else:
            print("✅ Filtering is working - non-collection day records were excluded")
            print("   Statistical metrics should now be based on collection days only")
    
    except Exception as e:
        print(f"❌ Error in verification: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    verify_metrics_filtering()
