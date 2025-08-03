#!/usr/bin/env python3
"""
Focused Discrepancy Investigation

Simple, focused script to identify the exact source of the 9,731 vs 9,372 discrepancy.
"""

import sys
import os
from pymongo import MongoClient

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def test_data_cleaning_filter():
    """Test the Data Cleaning sheet filter."""
    print("1. DATA CLEANING FILTER TEST")
    print("-" * 40)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Data Cleaning sheet filter (NO file_type filter)
    filter1 = {
        "School_Year": {"$ne": "N/A"},
        "is_collection_day": True,
        "Outlier_Status": False
    }
    
    count1 = collection.count_documents(filter1)
    print(f"Data Cleaning filter count: {count1}")
    return count1

def test_daily_counts_filter():
    """Test the Daily Counts (ACF_PACF) filter."""
    print("\n2. DAILY COUNTS FILTER TEST")
    print("-" * 40)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Daily Counts filter (WITH file_type filter)
    filter2 = {
        "School_Year": {"$ne": "N/A"},
        "file_type": {"$in": ["JPG", "MP3"]},
        "is_collection_day": True,
        "Outlier_Status": False
    }
    
    count2 = collection.count_documents(filter2)
    print(f"Daily Counts filter count: {count2}")
    return count2

def test_daily_counts_pipeline():
    """Test the actual Daily Counts pipeline."""
    print("\n3. DAILY COUNTS PIPELINE TEST")
    print("-" * 40)
    
    db = get_db_connection()
    collection = db.media_records
    
    try:
        from pipelines import PIPELINES
        
        pipeline_name = "DAILY_COUNTS_COLLECTION_ONLY"
        if pipeline_name in PIPELINES:
            pipeline = PIPELINES[pipeline_name]
            print(f"Found pipeline: {pipeline_name}")
            
            # Run the pipeline
            results = list(collection.aggregate(pipeline))
            print(f"Pipeline returned {len(results)} daily results")
            
            # Sum up the Total_Files
            total_files = 0
            for result in results:
                if 'Total_Files' in result:
                    total_files += result['Total_Files']
            
            print(f"Total files from pipeline aggregation: {total_files}")
            return total_files
        else:
            print(f"❌ Pipeline {pipeline_name} not found")
            return 0
            
    except ImportError as e:
        print(f"❌ Could not import pipelines: {e}")
        return 0

def test_daily_aggregation():
    """Test daily aggregation manually."""
    print("\n4. MANUAL DAILY AGGREGATION TEST")
    print("-" * 40)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Manual daily aggregation with clean filter
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
                "_id": "$ISO_Date",
                "Total_Files": {"$sum": 1}
            }
        }
    ]
    
    results = list(collection.aggregate(pipeline))
    print(f"Days with data: {len(results)}")
    
    total_files = sum(result['Total_Files'] for result in results)
    print(f"Total files from daily aggregation: {total_files}")
    
    return total_files

def main():
    print("FOCUSED DISCREPANCY INVESTIGATION")
    print("=" * 50)
    
    try:
        # Test all approaches
        count1 = test_data_cleaning_filter()
        count2 = test_daily_counts_filter()
        count3 = test_daily_counts_pipeline()
        count4 = test_daily_aggregation()
        
        print("\n" + "=" * 50)
        print("RESULTS SUMMARY")
        print("=" * 50)
        
        results = [
            ("Data Cleaning Filter", count1),
            ("Daily Counts Filter", count2),
            ("Daily Counts Pipeline", count3),
            ("Manual Daily Aggregation", count4)
        ]
        
        for name, count in results:
            if count == 9731:
                status = "✅ Expected"
            elif count == 9372:
                status = "🎯 DISCREPANCY SOURCE!"
            else:
                status = f"❓ Unexpected ({count})"
            
            print(f"{name:<25}: {count:>6} {status}")
        
        # Analysis
        print("\n" + "=" * 50)
        print("ANALYSIS")
        print("=" * 50)
        
        if count3 == 9372:
            print("🎯 FOUND: The Daily Counts Pipeline aggregation returns 9,372!")
            print("   This suggests the issue is in the pipeline aggregation logic,")
            print("   not in the basic filtering.")
        elif all(c == 9731 for c in [count1, count2, count3, count4]):
            print("❓ All database approaches return 9,731.")
            print("   The 9,372 discrepancy must be in Excel sheet generation/display.")
        else:
            print("❓ Mixed results - need further investigation.")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
