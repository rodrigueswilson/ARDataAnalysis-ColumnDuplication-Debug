#!/usr/bin/env python3
"""
Analyze whether the statistical metrics (Mean, Median, Std, Min, Max, Range)
in Summary Statistics are calculated correctly based on collection days logic.
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

def analyze_statistical_metrics():
    """Analyze how statistical metrics are calculated in Summary Statistics."""
    latest_report = find_latest_report()
    if not latest_report:
        print("No report files found!")
        return
    
    print(f"Analyzing statistical metrics in: {latest_report}")
    print("\n=== STATISTICAL METRICS ANALYSIS ===")
    
    try:
        # Read the Summary Statistics sheet
        df_summary = pd.read_excel(latest_report, sheet_name='Summary Statistics')
        
        # Extract the statistical metrics
        metrics_of_interest = [
            'Mean Files per Day',
            'Median Files per Day', 
            'Standard Deviation',
            'Min Files per Day',
            'Max Files per Day',
            'Range (Max - Min)'
        ]
        
        print("Current values in Summary Statistics:")
        print("-" * 50)
        
        summary_values = {}
        for metric in metrics_of_interest:
            metric_row = df_summary[df_summary['Metric'] == metric]
            if not metric_row.empty:
                val_2122 = metric_row['2021-2022'].iloc[0]
                val_2223 = metric_row['2022-2023'].iloc[0]
                val_total = metric_row['Total'].iloc[0]
                summary_values[metric] = {
                    '2021-2022': val_2122,
                    '2022-2023': val_2223,
                    'Total': val_total
                }
                print(f"{metric}:")
                print(f"  2021-2022: {val_2122}")
                print(f"  2022-2023: {val_2223}")
                print(f"  Total: {val_total}")
                print()
        
        # Now let's analyze the underlying data to understand how these should be calculated
        print("=== UNDERLYING DATA ANALYSIS ===")
        
        # We need to examine what data is being used for these calculations
        # Let's look at the DAILY_COUNTS_ALL pipeline data
        from pipelines.time_series import PIPELINES
        from report_generator.sheet_creators import SheetCreator
        
        # Create a temporary sheet creator to access the data
        sheet_creator = SheetCreator()
        df_daily = sheet_creator._run_aggregation(PIPELINES['DAILY_COUNTS_ALL'])
        
        if df_daily.empty:
            print("No daily data found!")
            return
        
        # Fill missing collection days and add school year
        df_daily = sheet_creator._fill_missing_collection_days(df_daily, 'DAILY_COUNTS_ALL_WITH_ZEROES')
        df_daily['date_obj'] = pd.to_datetime(df_daily['_id'])
        df_daily['school_year'] = df_daily['date_obj'].apply(
            lambda d: f"{d.year}-{d.year + 1}" if d.month >= 8 else f"{d.year - 1}-{d.year}"
        )
        
        print(f"Total daily records: {len(df_daily)}")
        print(f"Date range: {df_daily['_id'].min()} to {df_daily['_id'].max()}")
        print()
        
        # Analyze by school year
        school_years = sorted(df_daily['school_year'].unique())
        
        for year in school_years:
            year_data = df_daily[df_daily['school_year'] == year]
            file_counts = year_data['Total_Files']
            
            print(f"=== {year} ANALYSIS ===")
            print(f"Days with data: {len(year_data)}")
            
            # Calculate current statistics (what's in the report)
            current_stats = {
                'mean': round(file_counts.mean(), 1),
                'median': round(file_counts.median(), 1),
                'std': round(file_counts.std(), 1),
                'min': file_counts.min(),
                'max': file_counts.max(),
                'range': file_counts.max() - file_counts.min()
            }
            
            print(f"Current calculation (all days with data):")
            print(f"  Mean: {current_stats['mean']}")
            print(f"  Median: {current_stats['median']}")
            print(f"  Std: {current_stats['std']}")
            print(f"  Min: {current_stats['min']}")
            print(f"  Max: {current_stats['max']}")
            print(f"  Range: {current_stats['range']}")
            
            # Check if these match the Summary Statistics
            reported_stats = {}
            for metric, key in [
                ('Mean Files per Day', 'mean'),
                ('Median Files per Day', 'median'),
                ('Standard Deviation', 'std'),
                ('Min Files per Day', 'min'),
                ('Max Files per Day', 'max'),
                ('Range (Max - Min)', 'range')
            ]:
                if metric in summary_values:
                    reported_stats[key] = summary_values[metric][year]
            
            print(f"\nReported in Summary Statistics:")
            for key, value in reported_stats.items():
                print(f"  {key.title()}: {value}")
            
            print(f"\nComparison:")
            discrepancies = []
            for key in current_stats:
                if key in reported_stats:
                    current_val = current_stats[key]
                    reported_val = reported_stats[key]
                    if abs(current_val - reported_val) > 0.1:  # Allow small floating point differences
                        discrepancies.append(f"{key}: Current={current_val}, Reported={reported_val}")
                        print(f"  ⚠️  {key}: Current={current_val}, Reported={reported_val}")
                    else:
                        print(f"  ✅ {key}: Consistent ({current_val})")
            
            if discrepancies:
                print(f"  🚨 Found {len(discrepancies)} discrepancies!")
            else:
                print(f"  ✅ All statistics are consistent!")
            
            print()
        
        # Key question: Should these statistics exclude non-collection days?
        print("=== COLLECTION DAYS CONSIDERATION ===")
        print("QUESTION: Should statistical metrics be calculated from:")
        print("  Option A: All days with data (current approach)")
        print("  Option B: Only valid collection days (excluding non-collection days)")
        print()
        
        print("ANALYSIS:")
        print("- Current approach includes ALL days with data, even non-collection days")
        print("- This means statistics include data from holidays, breaks, etc.")
        print("- For educational research, we might want statistics from ONLY school days")
        print()
        
        # Let's check what days are included that shouldn't be
        from ar_utils import load_config
        config = load_config()
        non_collection_days = set()
        
        for year_data in config['school_years'].values():
            for period_data in year_data['periods'].values():
                non_collection_days.update(period_data.get('non_collection_days', []))
        
        print(f"Non-collection days defined in config: {len(non_collection_days)}")
        
        # Check if any of our data falls on non-collection days
        df_daily['date_str'] = df_daily['_id'].astype(str)
        non_collection_data = df_daily[df_daily['date_str'].isin(non_collection_days)]
        
        if not non_collection_data.empty:
            print(f"⚠️  Found {len(non_collection_data)} records on non-collection days:")
            print(f"   Date range: {non_collection_data['_id'].min()} to {non_collection_data['_id'].max()}")
            print(f"   File counts: {non_collection_data['Total_Files'].min()} to {non_collection_data['Total_Files'].max()}")
            print(f"   These days are included in current statistical calculations!")
        else:
            print("✅ No data found on non-collection days")
        
        print()
        print("RECOMMENDATION:")
        if not non_collection_data.empty:
            print("📋 Statistical metrics SHOULD be recalculated to exclude non-collection days")
            print("   This would make them consistent with the collection days logic")
            print("   and provide more accurate statistics for educational analysis")
        else:
            print("✅ Current statistical metrics are appropriate (no non-collection day data)")
    
    except Exception as e:
        print(f"Error analyzing statistical metrics: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_statistical_metrics()
