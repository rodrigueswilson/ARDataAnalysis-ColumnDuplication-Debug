#!/usr/bin/env python3
"""
Simple Date Exclusion Check

Quick check to identify why 358 files from early September 2021 are excluded.
"""

import sys
import os

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def main():
    print("SIMPLE DATE EXCLUSION CHECK")
    print("=" * 40)
    
    # The dates mentioned by user with their counts
    excluded_dates = {
        "2021-09-13": 63,
        "2021-09-14": 26,
        "2021-09-15": 69,
        "2021-09-16": 0,
        "2021-09-17": 37,
        "2021-09-21": 50,
        "2021-09-22": 1,
        "2021-09-23": 0,
        "2021-09-24": 80,
        "2021-09-27": 32
    }
    
    expected_total = sum(excluded_dates.values())
    print(f"Expected excluded total: {expected_total}")
    
    db = get_db_connection()
    collection = db.media_records
    
    # Check these specific dates in database
    actual_total = 0
    for date, expected_count in excluded_dates.items():
        actual_count = collection.count_documents({
            "ISO_Date": date,
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        
        status = "✅" if actual_count == expected_count else "❌"
        print(f"{date}: {actual_count:3d} (expected {expected_count:3d}) {status}")
        actual_total += actual_count
    
    print(f"\nActual total: {actual_total}")
    print(f"Expected: {expected_total}")
    print(f"Match: {'✅' if actual_total == expected_total else '❌'}")
    
    # Check if these dates are before a cutoff
    print(f"\nDiscrepancy check:")
    print(f"9731 - {actual_total} = {9731 - actual_total}")
    print(f"Expected result: 9372")
    print(f"Match: {'✅' if 9731 - actual_total == 9372 else '❌'}")
    
    # Quick check: are these dates collection days?
    try:
        from ar_utils import precompute_collection_days
        collection_days = precompute_collection_days()
        
        print(f"\nCollection days check:")
        for date in excluded_dates.keys():
            is_collection = date in collection_days
            print(f"{date}: {'✅ Collection day' if is_collection else '❌ Not collection day'}")
            
    except Exception as e:
        print(f"Could not check collection days: {e}")

if __name__ == "__main__":
    main()
