#!/usr/bin/env python3
"""
Current Collection Day Logic Investigation

This script investigates the current precompute_collection_days function
to understand exactly why early September dates are being excluded.
"""

import sys
import os
from datetime import datetime, date

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ar_utils import get_school_calendar, get_non_collection_days, precompute_collection_days

def investigate_current_collection_logic():
    """Investigate the current collection day logic."""
    print("CURRENT COLLECTION DAY LOGIC INVESTIGATION")
    print("=" * 60)
    
    try:
        # Get the current configuration
        school_calendar = get_school_calendar()
        non_collection_days = get_non_collection_days()
        
        print("1. SCHOOL CALENDAR CONFIGURATION:")
        print("-" * 40)
        for year, config in school_calendar.items():
            print(f"  {year}:")
            print(f"    Start: {config.get('start_date')}")
            print(f"    End: {config.get('end_date')}")
            
            periods = config.get('periods', {})
            for period, dates in periods.items():
                if isinstance(dates, list) and len(dates) >= 2:
                    print(f"    {period}: {dates[0]} to {dates[1]}")
        
        print(f"\n2. NON-COLLECTION DAYS:")
        print("-" * 40)
        print(f"Total non-collection days: {len(non_collection_days)}")
        
        # Show early September non-collection days
        early_september = []
        for date_obj, info in non_collection_days.items():
            if date_obj >= date(2021, 9, 1) and date_obj <= date(2021, 9, 30):
                early_september.append((date_obj, info))
        
        early_september.sort()
        print("Early September non-collection days:")
        for date_obj, info in early_september:
            print(f"  {date_obj}: {info.get('type', 'Unknown')} - {info.get('event', 'Unknown')}")
        
        print(f"\n3. CURRENT COLLECTION DAY MAP:")
        print("-" * 40)
        
        # Get the current collection day map
        collection_day_map = precompute_collection_days(school_calendar, non_collection_days)
        
        # Find the date range
        all_dates = sorted(collection_day_map.keys())
        if all_dates:
            min_date = min(all_dates)
            max_date = max(all_dates)
            print(f"Current collection day range: {min_date} to {max_date}")
            print(f"Total collection days: {len(all_dates)}")
            
            # Check early September specifically
            early_sept_dates = [d for d in all_dates if d >= date(2021, 9, 13) and d <= date(2021, 9, 30)]
            print(f"Early September collection days: {len(early_sept_dates)}")
            
            print("First 10 collection days:")
            for i, date_obj in enumerate(all_dates[:10]):
                info = collection_day_map[date_obj]
                print(f"  {i+1}. {date_obj}: {info.get('School_Year', 'Unknown')} - {info.get('Period', 'Unknown')}")
            
            # Check if 2021-09-13 is included
            target_date = date(2021, 9, 13)
            if target_date in collection_day_map:
                info = collection_day_map[target_date]
                print(f"\n✅ 2021-09-13 IS in collection day map:")
                print(f"   School Year: {info.get('School_Year', 'Unknown')}")
                print(f"   Period: {info.get('Period', 'Unknown')}")
            else:
                print(f"\n❌ 2021-09-13 is NOT in collection day map!")
                
                # Check why it's missing
                print("Checking why 2021-09-13 is missing...")
                
                # Is it a weekday?
                weekday = target_date.weekday()
                print(f"   Weekday: {weekday} ({'Weekday' if weekday < 5 else 'Weekend'})")
                
                # Is it a non-collection day?
                if target_date in non_collection_days:
                    info = non_collection_days[target_date]
                    print(f"   Non-collection: {info.get('type', 'Unknown')} - {info.get('event', 'Unknown')}")
                else:
                    print(f"   Non-collection: No")
                
                # Is it in the period range?
                for year, config in school_calendar.items():
                    periods = config.get('periods', {})
                    for period, dates in periods.items():
                        if isinstance(dates, list) and len(dates) >= 2:
                            start_date = dates[0] if isinstance(dates[0], date) else datetime.strptime(dates[0], '%Y-%m-%d').date()
                            end_date = dates[1] if isinstance(dates[1], date) else datetime.strptime(dates[1], '%Y-%m-%d').date()
                            
                            if start_date <= target_date <= end_date:
                                print(f"   In period: {period} ({start_date} to {end_date})")
                                break
        
        return collection_day_map
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

def main():
    try:
        collection_map = investigate_current_collection_logic()
        
        if collection_map:
            print("\n" + "=" * 60)
            print("SUMMARY")
            print("=" * 60)
            
            # Check if the issue is in the precompute function or elsewhere
            target_dates = [
                date(2021, 9, 13), date(2021, 9, 14), date(2021, 9, 15),
                date(2021, 9, 16), date(2021, 9, 17), date(2021, 9, 20),
                date(2021, 9, 21), date(2021, 9, 22), date(2021, 9, 23),
                date(2021, 9, 24), date(2021, 9, 27)
            ]
            
            missing_dates = []
            present_dates = []
            
            for target_date in target_dates:
                if target_date in collection_map:
                    present_dates.append(target_date)
                else:
                    missing_dates.append(target_date)
            
            print(f"Target dates present in collection map: {len(present_dates)}")
            for date_obj in present_dates:
                print(f"  ✅ {date_obj}")
            
            print(f"\nTarget dates missing from collection map: {len(missing_dates)}")
            for date_obj in missing_dates:
                print(f"  ❌ {date_obj}")
            
            if missing_dates:
                print(f"\n🎯 ISSUE CONFIRMED: {len(missing_dates)} early September dates are missing from collection day map!")
                print("This explains why they appear as left-aligned in Excel.")
            else:
                print(f"\n❓ UNEXPECTED: All target dates are in collection day map.")
                print("The issue must be elsewhere in the zero-fill logic.")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
