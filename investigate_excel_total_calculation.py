#!/usr/bin/env python3
"""
Excel Total Calculation Investigation

This script investigates the Excel sheet generation logic to understand
why the Daily Counts (ACF_PACF) sheet shows 9,372 in the total row
while the database contains 9,731 files.

Based on user observation:
- Total row (326) shows 9,372
- Left-aligned rows at top show 358 files that seem excluded
- These are dates from 2021-09-13 to 2021-09-27
"""

import sys
import os
import pandas as pd
from pathlib import Path

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def analyze_excluded_dates():
    """Analyze the specific dates that appear to be excluded."""
    print("1. ANALYZING EXCLUDED DATES")
    print("-" * 40)
    
    # The dates mentioned by user
    excluded_dates = [
        "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
        "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
    ]
    
    db = get_db_connection()
    collection = db.media_records
    
    # Check these specific dates
    total_excluded = 0
    for date in excluded_dates:
        count = collection.count_documents({
            "ISO_Date": date,
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        print(f"  {date}: {count} files")
        total_excluded += count
    
    print(f"\nTotal files in excluded dates: {total_excluded}")
    
    # Check if these are all from a specific school year
    print("\nSchool year breakdown for excluded dates:")
    pipeline = [
        {
            "$match": {
                "ISO_Date": {"$in": excluded_dates},
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "is_collection_day": True,
                "Outlier_Status": False
            }
        },
        {
            "$group": {
                "_id": "$School_Year",
                "Total_Files": {"$sum": 1}
            }
        }
    ]
    
    results = list(collection.aggregate(pipeline))
    for result in results:
        print(f"  {result['_id']}: {result['Total_Files']} files")
    
    return total_excluded

def check_date_range_filtering():
    """Check if there's date range filtering that excludes early dates."""
    print("\n2. CHECKING DATE RANGE FILTERING")
    print("-" * 40)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Get the earliest and latest dates in the clean dataset
    pipeline = [
        {
            "$match": {
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "is_collection_day": True,
                "Outlier_Status": False
            }
        },
        {
            "$group": {
                "_id": None,
                "min_date": {"$min": "$ISO_Date"},
                "max_date": {"$max": "$ISO_Date"},
                "total_files": {"$sum": 1}
            }
        }
    ]
    
    result = list(collection.aggregate(pipeline))[0]
    print(f"Date range: {result['min_date']} to {result['max_date']}")
    print(f"Total files in range: {result['total_files']}")
    
    # Check if there's a cutoff date being applied
    cutoff_candidates = ["2021-09-30", "2021-10-01", "2021-09-28"]
    
    for cutoff in cutoff_candidates:
        count_after = collection.count_documents({
            "ISO_Date": {"$gte": cutoff},
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        count_before = collection.count_documents({
            "ISO_Date": {"$lt": cutoff},
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        
        print(f"Files after {cutoff}: {count_after}")
        print(f"Files before {cutoff}: {count_before}")
        
        if count_after == 9372:
            print(f"🎯 FOUND: Cutoff date {cutoff} gives exactly 9,372 files!")
            return cutoff
    
    return None

def check_school_year_filtering():
    """Check if specific school years are being filtered out."""
    print("\n3. CHECKING SCHOOL YEAR FILTERING")
    print("-" * 40)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Get breakdown by school year
    pipeline = [
        {
            "$match": {
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "is_collection_day": True,
                "Outlier_Status": False
            }
        },
        {
            "$group": {
                "_id": "$School_Year",
                "Total_Files": {"$sum": 1}
            }
        },
        {
            "$sort": {"_id": 1}
        }
    ]
    
    results = list(collection.aggregate(pipeline))
    total_all_years = 0
    
    print("School year breakdown:")
    for result in results:
        print(f"  {result['_id']}: {result['Total_Files']} files")
        total_all_years += result['Total_Files']
    
    print(f"Total across all years: {total_all_years}")
    
    # Check if excluding first school year gives 9372
    if len(results) > 1:
        without_first = total_all_years - results[0]['Total_Files']
        print(f"Without first year ({results[0]['_id']}): {without_first} files")
        
        if without_first == 9372:
            print(f"🎯 FOUND: Excluding {results[0]['_id']} gives exactly 9,372 files!")
            return results[0]['_id']
    
    return None

def investigate_pipeline_date_logic():
    """Check if the pipeline has date filtering logic."""
    print("\n4. INVESTIGATING PIPELINE DATE LOGIC")
    print("-" * 40)
    
    try:
        from pipelines import PIPELINES
        
        pipeline_name = "DAILY_COUNTS_COLLECTION_ONLY"
        if pipeline_name in PIPELINES:
            pipeline = PIPELINES[pipeline_name]
            
            print(f"Pipeline {pipeline_name} stages:")
            for i, stage in enumerate(pipeline, 1):
                print(f"  Stage {i}: {list(stage.keys())[0]}")
                
                # Look for date filtering
                if '$match' in stage:
                    match_filter = stage['$match']
                    if 'ISO_Date' in match_filter:
                        print(f"    🎯 FOUND DATE FILTER: {match_filter['ISO_Date']}")
                        return match_filter['ISO_Date']
                    else:
                        print(f"    Filter: {match_filter}")
        
        return None
        
    except ImportError as e:
        print(f"Could not import pipelines: {e}")
        return None

def check_zero_fill_logic():
    """Check if zero-fill logic affects the totals."""
    print("\n5. CHECKING ZERO-FILL LOGIC")
    print("-" * 40)
    
    try:
        from ar_utils import precompute_collection_days
        
        # Get all collection days
        collection_days = precompute_collection_days()
        print(f"Total collection days from ar_utils: {len(collection_days)}")
        
        # Check if there's a date range in collection days
        if collection_days:
            min_date = min(collection_days)
            max_date = max(collection_days)
            print(f"Collection days range: {min_date} to {max_date}")
            
            # Check if the excluded dates are in collection days
            excluded_dates = [
                "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
                "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
            ]
            
            excluded_in_collection = [d for d in excluded_dates if d in collection_days]
            excluded_not_in_collection = [d for d in excluded_dates if d not in collection_days]
            
            print(f"Excluded dates in collection days: {len(excluded_in_collection)}")
            print(f"Excluded dates NOT in collection days: {len(excluded_not_in_collection)}")
            
            if excluded_not_in_collection:
                print(f"  Not in collection: {excluded_not_in_collection}")
            
            # Check if zero-fill starts from a specific date
            if min_date > "2021-09-13":
                print(f"🎯 FOUND: Zero-fill starts from {min_date}, excluding earlier dates!")
                return min_date
        
        return None
        
    except ImportError as e:
        print(f"Could not import ar_utils: {e}")
        return None

def main():
    print("EXCEL TOTAL CALCULATION INVESTIGATION")
    print("=" * 50)
    print("Investigating why 358 files are excluded from total")
    print("=" * 50)
    
    try:
        # Run all investigations
        excluded_count = analyze_excluded_dates()
        cutoff_date = check_date_range_filtering()
        excluded_year = check_school_year_filtering()
        pipeline_date_filter = investigate_pipeline_date_logic()
        zero_fill_start = check_zero_fill_logic()
        
        print("\n" + "=" * 50)
        print("INVESTIGATION SUMMARY")
        print("=" * 50)
        
        print(f"Excluded files count: {excluded_count}")
        print(f"Expected discrepancy: {9731 - 9372} = 359")
        print(f"Actual excluded count: {excluded_count}")
        
        if excluded_count == 359 or excluded_count == 358:
            print("✅ Excluded count matches discrepancy!")
        else:
            print("❓ Excluded count doesn't match discrepancy")
        
        # Identify the root cause
        root_causes = []
        
        if cutoff_date:
            root_causes.append(f"Date cutoff: {cutoff_date}")
        
        if excluded_year:
            root_causes.append(f"School year exclusion: {excluded_year}")
        
        if pipeline_date_filter:
            root_causes.append(f"Pipeline date filter: {pipeline_date_filter}")
        
        if zero_fill_start:
            root_causes.append(f"Zero-fill start date: {zero_fill_start}")
        
        if root_causes:
            print("\n🎯 POTENTIAL ROOT CAUSES:")
            for cause in root_causes:
                print(f"  - {cause}")
        else:
            print("\n❓ No clear root cause identified in date/year filtering")
            print("   The exclusion may be in Excel generation logic")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
