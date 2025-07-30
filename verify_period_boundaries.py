#!/usr/bin/env python3
"""
Verification script to check if both Collection_Days and Days_With_Data columns
are using the correct period boundaries from config.yaml.
"""

import pandas as pd
import glob
import os
from pathlib import Path
import yaml
import datetime

def find_latest_report():
    """Find the most recent AR Analysis Report Excel file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        print("No report files found!")
        return None
    
    # Sort by filename (timestamp), get the latest
    files.sort()
    latest_file = files[-1]
    print(f"Latest report: {latest_file}")
    return latest_file

def load_config_periods():
    """Load period boundaries from config.yaml."""
    with open('config.yaml', 'r') as f:
        config = yaml.safe_load(f)
    
    periods = {}
    for year, details in config.get('school_calendar', {}).items():
        for period_name, dates in details.get('periods', {}).items():
            if isinstance(dates, list) and len(dates) == 2:
                start_date = datetime.date.fromisoformat(dates[0])
                end_date = datetime.date.fromisoformat(dates[1])
                periods[period_name] = (start_date, end_date)
    
    return periods

def count_weekdays_in_period(start_date, end_date, non_collection_days):
    """Count weekdays in a period, excluding non-collection days."""
    count = 0
    current_date = start_date
    
    while current_date <= end_date:
        # Check if it's a weekday (Monday=0, Sunday=6)
        if current_date.weekday() < 5:  # Monday through Friday
            # Check if it's not a non-collection day
            if current_date not in non_collection_days:
                count += 1
        current_date += datetime.timedelta(days=1)
    
    return count

def verify_period_boundaries():
    """Verify that both columns use correct period boundaries."""
    latest_report = find_latest_report()
    if not latest_report:
        return
    
    # Load config periods
    config_periods = load_config_periods()
    
    # Load non-collection days
    with open('config.yaml', 'r') as f:
        config = yaml.safe_load(f)
    
    non_collection_days = set()
    for date_str in config.get('non_collection_days', {}):
        non_collection_days.add(datetime.date.fromisoformat(date_str))
    
    try:
        # Read the Period Counts (ACF_PACF) sheet
        df = pd.read_excel(latest_report, sheet_name='Period Counts (ACF_PACF)')
        
        print(f"\n=== Period Boundary Verification ===")
        print(f"Config.yaml periods:")
        for period, (start, end) in config_periods.items():
            print(f"  {period}: {start} to {end}")
        
        print(f"\n=== Comparison: Expected vs Actual ===")
        print(f"{'Period':<12} {'Expected':<10} {'Collection_Days':<15} {'Days_With_Data':<15} {'Match?'}")
        print("-" * 70)
        
        all_match = True
        for _, row in df.iterrows():
            period = row['Period']
            collection_days = int(row['Collection_Days'])
            days_with_data = int(row['Days_With_Data'])
            
            if period in config_periods:
                start_date, end_date = config_periods[period]
                expected_collection_days = count_weekdays_in_period(start_date, end_date, non_collection_days)
                
                collection_match = collection_days == expected_collection_days
                # Days_With_Data should be <= Collection_Days and within reasonable range
                data_reasonable = days_with_data <= collection_days
                
                match_status = "✓" if collection_match and data_reasonable else "✗"
                if not (collection_match and data_reasonable):
                    all_match = False
                
                print(f"{period:<12} {expected_collection_days:<10} {collection_days:<15} {days_with_data:<15} {match_status}")
            else:
                print(f"{period:<12} {'Unknown':<10} {collection_days:<15} {days_with_data:<15} {'?'}")
                all_match = False
        
        print(f"\n=== Summary ===")
        if all_match:
            print("✓ All periods are using correct boundaries!")
            print("✓ Collection_Days matches expected values")
            print("✓ Days_With_Data is within reasonable bounds")
        else:
            print("✗ Some periods have boundary issues")
        
        # Show specific issue for SY 22-23 P1 if it exists
        sy2223p1_row = df[df['Period'] == 'SY 22-23 P1']
        if not sy2223p1_row.empty:
            row = sy2223p1_row.iloc[0]
            print(f"\n=== SY 22-23 P1 Details ===")
            print(f"Collection_Days: {int(row['Collection_Days'])}")
            print(f"Days_With_Data: {int(row['Days_With_Data'])}")
            
            if 'SY 22-23 P1' in config_periods:
                start_date, end_date = config_periods['SY 22-23 P1']
                expected = count_weekdays_in_period(start_date, end_date, non_collection_days)
                print(f"Expected Collection_Days: {expected}")
                print(f"Period: {start_date} to {end_date}")
                
                if int(row['Days_With_Data']) > int(row['Collection_Days']):
                    print("⚠️  WARNING: Days_With_Data > Collection_Days indicates data outside period boundaries!")
                elif int(row['Days_With_Data']) == int(row['Collection_Days']):
                    print("✓ Perfect match - all collection days have data")
                else:
                    missing_days = int(row['Collection_Days']) - int(row['Days_With_Data'])
                    print(f"ℹ️  {missing_days} collection days without data (normal)")
        
    except Exception as e:
        print(f"Error reading the Excel file: {e}")

if __name__ == "__main__":
    verify_period_boundaries()
