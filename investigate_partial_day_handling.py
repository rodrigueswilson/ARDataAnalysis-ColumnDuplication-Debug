#!/usr/bin/env python3
"""
Partial Day Handling Investigation

Now that we found 2021-09-20 (104 files) as a "Partial" collection day,
we need to understand how Partial days are handled differently in Excel
generation vs database counting.
"""

import sys
import os

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def investigate_partial_days():
    """Investigate how partial collection days are handled."""
    print("PARTIAL DAY HANDLING INVESTIGATION")
    print("=" * 50)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Get all partial days from the early September range
    partial_dates_in_range = [
        "2021-09-20"  # We know this one is partial
    ]
    
    print("1. CHECKING PARTIAL DAYS IN EARLY SEPTEMBER")
    print("-" * 40)
    
    total_partial_files = 0
    for date in partial_dates_in_range:
        count = collection.count_documents({
            "ISO_Date": date,
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        print(f"{date}: {count} files (Partial day)")
        total_partial_files += count
    
    print(f"Total files on partial days: {total_partial_files}")
    
    # Now let's recalculate the exclusion
    print("\n2. RECALCULATING EXCLUSIONS")
    print("-" * 40)
    
    # Known excluded dates (from your Excel observation)
    regular_excluded = 358
    partial_excluded = total_partial_files
    
    print(f"Regular excluded dates: {regular_excluded} files")
    print(f"Partial day excluded: {partial_excluded} files")
    print(f"Total excluded: {regular_excluded + partial_excluded} files")
    
    # Check the math
    database_total = 9731
    excel_total = 9372
    expected_excluded = database_total - excel_total
    actual_excluded = regular_excluded + partial_excluded
    
    print(f"\nDatabase total: {database_total}")
    print(f"Excel total: {excel_total}")
    print(f"Expected excluded: {expected_excluded}")
    print(f"Actual excluded: {actual_excluded}")
    print(f"Difference: {expected_excluded - actual_excluded}")
    
    if expected_excluded == actual_excluded:
        print("✅ PERFECT MATCH! All exclusions accounted for.")
    else:
        print(f"❓ Still missing {expected_excluded - actual_excluded} files")
        
        # Let's check if there are more dates we're missing
        print("\n3. LOOKING FOR MORE EXCLUDED DATES")
        print("-" * 40)
        
        # Check a broader range around the problematic period
        all_dates_in_period = collection.distinct("ISO_Date", {
            "ISO_Date": {"$gte": "2021-09-13", "$lte": "2021-10-15"},
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        
        all_dates_in_period.sort()
        
        # Known dates
        known_excluded = [
            "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
            "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27",
            "2021-09-20"  # The partial day we found
        ]
        
        print("All dates in extended range (2021-09-13 to 2021-10-15):")
        
        additional_excluded = 0
        for date in all_dates_in_period:
            count = collection.count_documents({
                "ISO_Date": date,
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "is_collection_day": True,
                "Outlier_Status": False
            })
            
            is_known = date in known_excluded
            status = "✅ Known" if is_known else "🔍 NEW!"
            
            if not is_known and count > 0:
                additional_excluded += count
                print(f"  {date}: {count} files {status}")
            elif is_known:
                print(f"  {date}: {count} files {status}")
        
        if additional_excluded > 0:
            print(f"\n🎯 FOUND {additional_excluded} additional excluded files!")
            print(f"New total excluded: {actual_excluded + additional_excluded}")
            
            if expected_excluded == actual_excluded + additional_excluded:
                print("✅ NOW PERFECT MATCH!")
            else:
                print(f"❓ Still off by {expected_excluded - (actual_excluded + additional_excluded)}")

def main():
    try:
        investigate_partial_days()
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
