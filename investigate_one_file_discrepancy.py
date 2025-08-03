#!/usr/bin/env python3
"""
One-File Discrepancy Investigation

This script investigates the 1-file difference between:
- Expected discrepancy: 9,731 - 9,372 = 359 files
- Observed excluded files: 358 files (from Excel left-aligned rows)
- Missing: 1 file

We need to find where this 1 additional file is being excluded.
"""

import sys
import os

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def investigate_exact_counts():
    """Get exact counts to identify the 1-file discrepancy."""
    print("1. EXACT COUNT INVESTIGATION")
    print("-" * 40)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Get the exact count using Data Cleaning logic
    data_cleaning_count = collection.count_documents({
        "School_Year": {"$ne": "N/A"},
        "is_collection_day": True,
        "Outlier_Status": False
    })
    
    # Get the exact count using Daily Counts logic
    daily_counts_count = collection.count_documents({
        "School_Year": {"$ne": "N/A"},
        "file_type": {"$in": ["JPG", "MP3"]},
        "is_collection_day": True,
        "Outlier_Status": False
    })
    
    print(f"Data Cleaning count: {data_cleaning_count}")
    print(f"Daily Counts count: {daily_counts_count}")
    print(f"Difference: {data_cleaning_count - daily_counts_count}")
    
    # The observed Excel total
    excel_total = 9372
    
    print(f"\nExcel total: {excel_total}")
    print(f"Expected discrepancy: {data_cleaning_count - excel_total}")
    print(f"Observed excluded files: 358")
    print(f"Missing file count: {(data_cleaning_count - excel_total) - 358}")
    
    return data_cleaning_count, daily_counts_count, excel_total

def check_specific_excluded_dates():
    """Check the exact counts for the specific excluded dates."""
    print("\n2. SPECIFIC EXCLUDED DATES CHECK")
    print("-" * 40)
    
    db = get_db_connection()
    collection = db.media_records
    
    # The dates from user observation
    excluded_dates = [
        "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
        "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
    ]
    
    total_excluded = 0
    for date in excluded_dates:
        # Use the same filter as Daily Counts (with file_type)
        count = collection.count_documents({
            "ISO_Date": date,
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        print(f"{date}: {count} files")
        total_excluded += count
    
    print(f"\nTotal excluded from specific dates: {total_excluded}")
    
    # Check if there are any other dates in the same range
    print("\n3. CHECKING FOR ADDITIONAL DATES IN RANGE")
    print("-" * 40)
    
    # Check for any other dates between 2021-09-13 and 2021-09-27
    additional_dates = collection.distinct("ISO_Date", {
        "ISO_Date": {"$gte": "2021-09-13", "$lte": "2021-09-27"},
        "School_Year": {"$ne": "N/A"},
        "file_type": {"$in": ["JPG", "MP3"]},
        "is_collection_day": True,
        "Outlier_Status": False
    })
    
    additional_dates.sort()
    print(f"All dates in range 2021-09-13 to 2021-09-27:")
    
    range_total = 0
    for date in additional_dates:
        count = collection.count_documents({
            "ISO_Date": date,
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        
        is_known = date in excluded_dates
        status = "✅ Known" if is_known else "🔍 NEW!"
        print(f"  {date}: {count} files {status}")
        range_total += count
    
    print(f"\nTotal files in range: {range_total}")
    
    # Find the missing dates
    missing_dates = [d for d in additional_dates if d not in excluded_dates]
    if missing_dates:
        print(f"\n🎯 FOUND MISSING DATES: {missing_dates}")
        missing_total = 0
        for date in missing_dates:
            count = collection.count_documents({
                "ISO_Date": date,
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "is_collection_day": True,
                "Outlier_Status": False
            })
            missing_total += count
        print(f"Missing dates total: {missing_total}")
        return missing_total
    
    return 0

def check_edge_cases():
    """Check for edge cases that might explain the 1-file difference."""
    print("\n4. CHECKING EDGE CASES")
    print("-" * 40)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Check if there are files with different file_type values
    print("Checking file types in the clean dataset...")
    
    file_types = collection.distinct("file_type", {
        "School_Year": {"$ne": "N/A"},
        "is_collection_day": True,
        "Outlier_Status": False
    })
    
    print(f"File types in clean dataset: {file_types}")
    
    for file_type in file_types:
        count = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": True,
            "Outlier_Status": False,
            "file_type": file_type
        })
        
        in_filter = file_type in ["JPG", "MP3"]
        status = "✅ Included" if in_filter else "❌ EXCLUDED"
        print(f"  {file_type}: {count} files {status}")
        
        if not in_filter and count > 0:
            print(f"    🎯 FOUND: {count} files with type '{file_type}' excluded by file_type filter!")
            return count
    
    return 0

def main():
    print("ONE-FILE DISCREPANCY INVESTIGATION")
    print("=" * 50)
    print("Finding the missing 1 file in the 359 vs 358 discrepancy")
    print("=" * 50)
    
    try:
        # Get exact counts
        data_cleaning_count, daily_counts_count, excel_total = investigate_exact_counts()
        
        # Check specific excluded dates
        missing_from_dates = check_specific_excluded_dates()
        
        # Check edge cases
        missing_from_edge_cases = check_edge_cases()
        
        print("\n" + "=" * 50)
        print("INVESTIGATION SUMMARY")
        print("=" * 50)
        
        expected_discrepancy = data_cleaning_count - excel_total
        observed_excluded = 358
        actual_missing = expected_discrepancy - observed_excluded
        
        print(f"Expected discrepancy: {expected_discrepancy}")
        print(f"Observed excluded files: {observed_excluded}")
        print(f"Actual missing files: {actual_missing}")
        
        if missing_from_dates > 0:
            print(f"🎯 FOUND {missing_from_dates} missing files in date range!")
        
        if missing_from_edge_cases > 0:
            print(f"🎯 FOUND {missing_from_edge_cases} missing files in edge cases!")
        
        total_found = missing_from_dates + missing_from_edge_cases
        if total_found == actual_missing:
            print(f"✅ COMPLETE: All {actual_missing} missing files accounted for!")
        elif total_found > actual_missing:
            print(f"❓ OVER: Found {total_found} but only {actual_missing} missing")
        else:
            print(f"❓ UNDER: Found {total_found} but {actual_missing} still missing")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
