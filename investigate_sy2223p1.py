#!/usr/bin/env python3
"""
Investigation script to examine the actual dates in SY 22-23 P1 
to understand why Days_With_Data > Collection_Days.
"""

from db_utils import get_db_connection
from collections import Counter
import datetime

def investigate_sy2223p1_dates():
    """Investigate the actual dates in SY 22-23 P1 data."""
    
    # Connect to MongoDB using the correct database connection
    db = get_db_connection()
    collection = db["media_records"]
    
    print("=== SY 22-23 P1 Date Investigation ===")
    print("Period boundaries: 2022-09-12 to 2022-12-12")
    print()
    
    # Query for all SY 22-23 P1 records
    query = {"Collection_Period": "SY 22-23 P1"}
    records = list(collection.find(query, {"ISO_Date": 1, "Collection_Period": 1, "_id": 0}))
    
    print(f"Total records found: {len(records)}")
    
    # Count unique dates
    dates = [record["ISO_Date"] for record in records]
    date_counts = Counter(dates)
    unique_dates = sorted(date_counts.keys())
    
    print(f"Unique dates: {len(unique_dates)}")
    print()
    
    # Define period boundaries
    period_start = datetime.date(2022, 9, 12)
    period_end = datetime.date(2022, 12, 12)
    
    # Analyze each date
    within_period = []
    outside_period = []
    
    for date_str in unique_dates:
        try:
            date_obj = datetime.date.fromisoformat(date_str)
            file_count = date_counts[date_str]
            
            if period_start <= date_obj <= period_end:
                within_period.append((date_str, file_count))
            else:
                outside_period.append((date_str, file_count))
                
        except ValueError:
            print(f"Invalid date format: {date_str}")
    
    print(f"=== Dates WITHIN period boundaries ({period_start} to {period_end}) ===")
    print(f"Count: {len(within_period)} days")
    for date_str, count in within_period[:10]:  # Show first 10
        print(f"  {date_str}: {count} files")
    if len(within_period) > 10:
        print(f"  ... and {len(within_period) - 10} more dates")
    
    print()
    print(f"=== Dates OUTSIDE period boundaries ===")
    print(f"Count: {len(outside_period)} days")
    for date_str, count in outside_period:
        print(f"  {date_str}: {count} files")
    
    print()
    print(f"=== Summary ===")
    print(f"Days within period: {len(within_period)}")
    print(f"Days outside period: {len(outside_period)}")
    print(f"Total unique days: {len(unique_dates)}")
    
    # Check if the outside dates are causing the discrepancy
    if len(outside_period) > 0:
        print(f"\n⚠️  Found {len(outside_period)} days with data outside the period boundaries!")
        print("This explains why Days_With_Data > Collection_Days")
        
        # Show the specific outside dates
        for date_str, count in outside_period:
            date_obj = datetime.date.fromisoformat(date_str)
            if date_obj < period_start:
                print(f"  {date_str}: {count} files (BEFORE period start)")
            elif date_obj > period_end:
                print(f"  {date_str}: {count} files (AFTER period end)")
    
    # Note: db_utils handles connection management

if __name__ == "__main__":
    investigate_sy2223p1_dates()
