#!/usr/bin/env python3
"""
Compare the actual days with data vs calculated collection days
to identify the specific discrepancy.
"""

from db_utils import get_db_connection
from ar_utils import calculate_collection_days_for_period, get_school_calendar, get_non_collection_days
import datetime
from collections import Counter

def compare_collection_vs_data_days():
    """Compare actual data days vs calculated collection days for SY 22-23 P1."""
    
    # Get database connection
    db = get_db_connection()
    collection = db["media_records"]
    
    # Get actual dates with data for SY 22-23 P1
    query = {"Collection_Period": "SY 22-23 P1"}
    records = list(collection.find(query, {"ISO_Date": 1, "_id": 0}))
    
    actual_dates = sorted(set(record["ISO_Date"] for record in records))
    actual_date_objects = [datetime.date.fromisoformat(date_str) for date_str in actual_dates]
    
    print("=== SY 22-23 P1: Collection Days vs Data Days Comparison ===")
    print(f"Period boundaries: 2022-09-12 to 2022-12-12")
    print()
    
    # Calculate collection days using my function
    calculated_collection_days = calculate_collection_days_for_period("SY 22-23 P1")
    print(f"Calculated Collection_Days: {calculated_collection_days}")
    print(f"Actual Days_With_Data: {len(actual_dates)}")
    print(f"Discrepancy: {len(actual_dates) - calculated_collection_days}")
    print()
    
    # Get period boundaries and non-collection days
    school_calendar = get_school_calendar()
    non_collection_days = get_non_collection_days()
    
    period_dates = None
    for year_data in school_calendar.values():
        if "SY 22-23 P1" in year_data.get('periods', {}):
            period_dates = year_data['periods']["SY 22-23 P1"]
            break
    
    if not period_dates:
        print("Could not find period dates!")
        return
    
    start_date, end_date = period_dates
    
    # Generate all expected collection days (my calculation logic)
    expected_collection_days = []
    current_date = start_date
    
    while current_date <= end_date:
        # Check if it's a weekday (Monday=0, Sunday=6)
        if current_date.weekday() < 5:  # Monday through Friday
            # Check if it's not a non-collection day
            if current_date not in non_collection_days:
                expected_collection_days.append(current_date)
        current_date += datetime.timedelta(days=1)
    
    print(f"=== Detailed Analysis ===")
    print(f"Expected collection days (my calculation): {len(expected_collection_days)}")
    print(f"Actual days with data: {len(actual_date_objects)}")
    print()
    
    # Find days with data that are NOT in my expected collection days
    actual_set = set(actual_date_objects)
    expected_set = set(expected_collection_days)
    
    data_but_not_expected = actual_set - expected_set
    expected_but_no_data = expected_set - actual_set
    
    if data_but_not_expected:
        print(f"=== Days with DATA but NOT in my Collection_Days calculation ===")
        for date_obj in sorted(data_but_not_expected):
            day_name = date_obj.strftime('%A')
            is_weekend = date_obj.weekday() >= 5
            is_non_collection = date_obj in non_collection_days
            
            reason = []
            if is_weekend:
                reason.append("Weekend")
            if is_non_collection:
                event = non_collection_days[date_obj].get('event', 'Unknown')
                reason.append(f"Non-collection ({event})")
            
            reason_str = ", ".join(reason) if reason else "Unknown reason"
            print(f"  {date_obj} ({day_name}): {reason_str}")
    
    if expected_but_no_data:
        print(f"\n=== Days in Collection_Days calculation but NO data ===")
        print(f"Count: {len(expected_but_no_data)} days")
        for date_obj in sorted(expected_but_no_data)[:5]:  # Show first 5
            day_name = date_obj.strftime('%A')
            print(f"  {date_obj} ({day_name}): Expected collection day but no data")
        if len(expected_but_no_data) > 5:
            print(f"  ... and {len(expected_but_no_data) - 5} more days")
    
    print(f"\n=== Summary ===")
    if len(data_but_not_expected) == 1 and len(expected_but_no_data) == 0:
        print("✓ Found the issue: 1 day has data but is excluded from Collection_Days calculation")
        print("  This explains why Days_With_Data > Collection_Days")
    elif len(data_but_not_expected) == 0 and len(expected_but_no_data) == 1:
        print("✓ Found the issue: 1 expected collection day has no data")
        print("  This explains why Collection_Days > Days_With_Data")
    else:
        print(f"Complex discrepancy: {len(data_but_not_expected)} extra data days, {len(expected_but_no_data)} missing data days")

if __name__ == "__main__":
    compare_collection_vs_data_days()
