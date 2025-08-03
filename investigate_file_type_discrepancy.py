#!/usr/bin/env python3
"""
Investigate File Type Discrepancy Between Data Cleaning and Daily Counts (ACF_PACF)

This script investigates the discrepancy where:
- Data Cleaning sheet reports 9,731 clean files
- Daily Counts (ACF_PACF) sheet reports 9,372 clean files
- Difference: 359 files

Root cause hypothesis: Daily Counts includes file_type filter, Data Cleaning doesn't.
"""

import sys
import os
from pymongo import MongoClient
from collections import defaultdict, Counter

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def investigate_file_type_discrepancy():
    """
    Investigate the file type discrepancy between Data Cleaning and Daily Counts sheets.
    """
    print("=" * 80)
    print("INVESTIGATING FILE TYPE DISCREPANCY")
    print("=" * 80)
    
    # Get database connection
    db = get_db_connection()
    collection = db.media_records
    
    print(f"Connected to database: {db.name}")
    print(f"Collection: {collection.name}")
    print(f"Total documents: {collection.count_documents({})}")
    print()
    
    # 1. Data Cleaning Sheet Logic (NO file_type filter)
    print("1. DATA CLEANING SHEET LOGIC")
    print("-" * 40)
    data_cleaning_filter = {
        "School_Year": {"$ne": "N/A"},
        "is_collection_day": True,
        "Outlier_Status": False
    }
    
    data_cleaning_count = collection.count_documents(data_cleaning_filter)
    print(f"Clean files (Data Cleaning logic): {data_cleaning_count}")
    
    # Get file type breakdown for Data Cleaning logic
    data_cleaning_pipeline = [
        {"$match": data_cleaning_filter},
        {"$group": {"_id": "$file_type", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}}
    ]
    
    data_cleaning_types = list(collection.aggregate(data_cleaning_pipeline))
    print("File type breakdown (Data Cleaning logic):")
    for item in data_cleaning_types:
        file_type = item['_id'] if item['_id'] else 'NULL'
        count = item['count']
        print(f"  {file_type}: {count}")
    print()
    
    # 2. Daily Counts (ACF_PACF) Sheet Logic (WITH file_type filter)
    print("2. DAILY COUNTS (ACF_PACF) SHEET LOGIC")
    print("-" * 40)
    daily_counts_filter = {
        "School_Year": {"$ne": "N/A"},
        "file_type": {"$in": ["JPG", "MP3"]},  # Extra filter!
        "is_collection_day": True,
        "Outlier_Status": False
    }
    
    daily_counts_count = collection.count_documents(daily_counts_filter)
    print(f"Clean files (Daily Counts logic): {daily_counts_count}")
    
    # Get file type breakdown for Daily Counts logic
    daily_counts_pipeline = [
        {"$match": daily_counts_filter},
        {"$group": {"_id": "$file_type", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}}
    ]
    
    daily_counts_types = list(collection.aggregate(daily_counts_pipeline))
    print("File type breakdown (Daily Counts logic):")
    for item in daily_counts_types:
        file_type = item['_id'] if item['_id'] else 'NULL'
        count = item['count']
        print(f"  {file_type}: {count}")
    print()
    
    # 3. Calculate the difference
    print("3. DISCREPANCY ANALYSIS")
    print("-" * 40)
    difference = data_cleaning_count - daily_counts_count
    print(f"Difference: {difference} files")
    print(f"Data Cleaning: {data_cleaning_count}")
    print(f"Daily Counts:  {daily_counts_count}")
    print(f"Missing:       {difference}")
    print()
    
    # 4. Find the missing files (those excluded by file_type filter)
    print("4. MISSING FILES ANALYSIS")
    print("-" * 40)
    
    # Files that meet clean criteria but are NOT JPG/MP3
    missing_files_filter = {
        "School_Year": {"$ne": "N/A"},
        "is_collection_day": True,
        "Outlier_Status": False,
        "file_type": {"$nin": ["JPG", "MP3"]}  # NOT JPG or MP3
    }
    
    missing_files_count = collection.count_documents(missing_files_filter)
    print(f"Files excluded by file_type filter: {missing_files_count}")
    
    # Get breakdown of excluded file types
    missing_files_pipeline = [
        {"$match": missing_files_filter},
        {"$group": {"_id": "$file_type", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}}
    ]
    
    missing_file_types = list(collection.aggregate(missing_files_pipeline))
    print("Excluded file types:")
    for item in missing_file_types:
        file_type = item['_id'] if item['_id'] else 'NULL'
        count = item['count']
        print(f"  {file_type}: {count}")
    print()
    
    # 5. Verification
    print("5. VERIFICATION")
    print("-" * 40)
    expected_difference = missing_files_count
    actual_difference = difference
    
    print(f"Expected difference: {expected_difference}")
    print(f"Actual difference:   {actual_difference}")
    print(f"Match: {'✅ YES' if expected_difference == actual_difference else '❌ NO'}")
    print()
    
    # 6. Sample of excluded files
    print("6. SAMPLE OF EXCLUDED FILES")
    print("-" * 40)
    sample_excluded = list(collection.find(missing_files_filter).limit(10))
    
    if sample_excluded:
        print("Sample excluded files:")
        for i, doc in enumerate(sample_excluded, 1):
            file_path = doc.get('File_Path', 'Unknown')
            file_type = doc.get('file_type', 'Unknown')
            school_year = doc.get('School_Year', 'Unknown')
            print(f"  {i}. {file_path} (Type: {file_type}, Year: {school_year})")
    else:
        print("No excluded files found.")
    print()
    
    # 7. Recommendations
    print("7. RECOMMENDATIONS")
    print("-" * 40)
    print("The discrepancy is caused by inconsistent file_type filtering:")
    print("• Data Cleaning sheet: Includes ALL file types")
    print("• Daily Counts (ACF_PACF): Only includes JPG and MP3 files")
    print()
    print("Options to resolve:")
    print("A. Add file_type filter to Data Cleaning sheet (recommended)")
    print("B. Remove file_type filter from Daily Counts pipeline")
    print("C. Document the difference and explain why it exists")
    print()
    
    return {
        'data_cleaning_count': data_cleaning_count,
        'daily_counts_count': daily_counts_count,
        'difference': difference,
        'excluded_count': missing_files_count,
        'data_cleaning_types': data_cleaning_types,
        'daily_counts_types': daily_counts_types,
        'excluded_types': missing_file_types
    }

if __name__ == "__main__":
    try:
        results = investigate_file_type_discrepancy()
        print("Investigation completed successfully!")
        
    except Exception as e:
        print(f"Error during investigation: {e}")
        import traceback
        traceback.print_exc()
