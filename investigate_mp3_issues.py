#!/usr/bin/env python3
"""
Investigate MP3 Duration Analysis issues:
1. Missing 2022-2023 school year row
2. Header formatting inconsistency
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection
import pandas as pd

def investigate_mp3_data():
    """Investigate MP3 data and school year issues."""
    
    print("=== MP3 Duration Analysis Investigation ===\n")
    
    db = get_db_connection()
    
    # 1. Check what school years exist for MP3 files
    print("1. Checking distinct School Years for MP3 files...")
    distinct_years = db.media_records.distinct('School_Year', {'file_type': 'MP3'})
    print(f"Distinct School Years: {distinct_years}")
    
    # 2. Check MP3 data by school year with filters
    print("\n2. Checking MP3 data by School Year with filters...")
    pipeline = [
        {"$match": {
            "file_type": "MP3",
            "is_collection_day": True,
            "Outlier_Status": False
        }},
        {"$group": {
            "_id": "$School_Year",
            "Total_MP3_Files": {"$sum": 1},
            "Total_Duration_Seconds": {"$sum": "$Duration_Seconds"}
        }},
        {"$sort": {"_id": 1}}
    ]
    
    result = list(db.media_records.aggregate(pipeline))
    print("MP3 data by School Year (with filters):")
    for item in result:
        school_year = item['_id']
        files = item['Total_MP3_Files']
        duration = item['Total_Duration_Seconds']
        hours = duration / 3600 if duration else 0
        print(f"  School Year: {school_year}, Files: {files}, Duration: {duration}s ({hours:.2f}h)")
    
    # 3. Check without filters to see if 2022-2023 data exists
    print("\n3. Checking MP3 data by School Year WITHOUT filters...")
    pipeline_no_filter = [
        {"$match": {"file_type": "MP3"}},
        {"$group": {
            "_id": "$School_Year",
            "Total_MP3_Files": {"$sum": 1},
            "With_Collection_Day": {"$sum": {"$cond": [{"$eq": ["$is_collection_day", True]}, 1, 0]}},
            "Non_Outliers": {"$sum": {"$cond": [{"$eq": ["$Outlier_Status", False]}, 1, 0]}},
            "Total_Duration_Seconds": {"$sum": "$Duration_Seconds"}
        }},
        {"$sort": {"_id": 1}}
    ]
    
    result_no_filter = list(db.media_records.aggregate(pipeline_no_filter))
    print("MP3 data by School Year (NO filters):")
    for item in result_no_filter:
        school_year = item['_id']
        total_files = item['Total_MP3_Files']
        collection_day_files = item['With_Collection_Day']
        non_outlier_files = item['Non_Outliers']
        duration = item['Total_Duration_Seconds']
        print(f"  School Year: {school_year}")
        print(f"    Total Files: {total_files}")
        print(f"    Collection Day Files: {collection_day_files}")
        print(f"    Non-Outlier Files: {non_outlier_files}")
        print(f"    Duration: {duration}s")
    
    # 4. Sample some 2022-2023 records if they exist
    print("\n4. Sample 2022-2023 MP3 records...")
    sample_2022_2023 = list(db.media_records.find(
        {
            'file_type': 'MP3', 
            'School_Year': '2022-2023'
        }, 
        {
            'School_Year': 1, 
            'is_collection_day': 1, 
            'Outlier_Status': 1, 
            'Duration_Seconds': 1,
            'ISO_Date': 1
        }
    ).limit(5))
    
    if sample_2022_2023:
        print("Found 2022-2023 MP3 records:")
        for record in sample_2022_2023:
            print(f"  {record}")
    else:
        print("No 2022-2023 MP3 records found!")

def check_header_formatting():
    """Check header formatting in the MP3 Duration Analysis sheet."""
    
    print("\n=== Header Formatting Investigation ===\n")
    
    # Look at the current header implementation
    from report_generator.sheet_creators.specialized import SpecializedSheetCreator
    
    print("Current header implementation in _add_duration_summary_table:")
    print("Headers: ['School Year', 'Total Files', 'Total Hours', 'Avg Duration', 'Min Duration', 'Max Duration']")
    
    # Compare with other sheets' header formatting
    print("\nComparing with other sheet headers...")
    print("Expected format should match other sheets in the system")
    
    # Check what the standard header format should be
    from report_generator.formatters import ExcelFormatter
    formatter = ExcelFormatter()
    
    print("Standard header formatting should use:")
    print("- formatter.apply_header_style() for consistent styling")
    print("- Bold font, background color, borders")
    print("- Consistent with other sheet headers")

if __name__ == "__main__":
    investigate_mp3_data()
    check_header_formatting()
