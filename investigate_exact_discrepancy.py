#!/usr/bin/env python3
"""
Comprehensive Discrepancy Investigation

This script performs a detailed investigation to identify the exact root cause
of the discrepancy between Data Cleaning sheet (9,731) and Daily Counts (ACF_PACF) (9,372).

We need to trace through the entire data flow to find where the 359 files are lost.
"""

import sys
import os
from pymongo import MongoClient
from collections import defaultdict, Counter
import pandas as pd

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection
from pipelines import PIPELINES

def investigate_data_cleaning_logic():
    """
    Replicate the exact Data Cleaning sheet logic.
    """
    print("=" * 80)
    print("1. DATA CLEANING SHEET LOGIC REPLICATION")
    print("=" * 80)
    
    db = get_db_connection()
    collection = db.media_records
    
    # This is the exact filter used in Data Cleaning sheet
    data_cleaning_filter = {
        "School_Year": {"$ne": "N/A"},
        "is_collection_day": True,
        "Outlier_Status": False
    }
    
    count = collection.count_documents(data_cleaning_filter)
    print(f"Data Cleaning logic count: {count}")
    
    # Get sample documents
    sample_docs = list(collection.find(data_cleaning_filter).limit(5))
    print("Sample documents:")
    for i, doc in enumerate(sample_docs, 1):
        print(f"  {i}. File: {doc.get('File_Path', 'Unknown')}")
        print(f"     Type: {doc.get('file_type', 'Unknown')}")
        print(f"     Date: {doc.get('ISO_Date', 'Unknown')}")
        print(f"     School Year: {doc.get('School_Year', 'Unknown')}")
        print(f"     Collection Day: {doc.get('is_collection_day', 'Unknown')}")
        print(f"     Outlier Status: {doc.get('Outlier_Status', 'Unknown')}")
        print()
    
    return count, data_cleaning_filter

def investigate_daily_counts_pipeline():
    """
    Replicate the exact Daily Counts (ACF_PACF) pipeline logic.
    """
    print("=" * 80)
    print("2. DAILY COUNTS (ACF_PACF) PIPELINE REPLICATION")
    print("=" * 80)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Get the exact pipeline used by Daily Counts (ACF_PACF)
    pipeline_name = "DAILY_COUNTS_COLLECTION_ONLY"
    
    if pipeline_name not in PIPELINES:
        print(f"❌ Pipeline {pipeline_name} not found!")
        return 0, None
    
    pipeline = PIPELINES[pipeline_name]
    print(f"Using pipeline: {pipeline_name}")
    print("Pipeline stages:")
    for i, stage in enumerate(pipeline, 1):
        print(f"  Stage {i}: {list(stage.keys())[0]}")
        if '$match' in stage:
            print(f"    Filter: {stage['$match']}")
    
    # Run the pipeline
    results = list(collection.aggregate(pipeline))
    
    print(f"Pipeline returned {len(results)} aggregated results")
    
    # Calculate total files from aggregation results
    total_files = 0
    for result in results:
        if 'Total_Files' in result:
            total_files += result['Total_Files']
    
    print(f"Total files from pipeline aggregation: {total_files}")
    
    # Also test the raw filter from the pipeline
    if pipeline and '$match' in pipeline[0]:
        raw_filter = pipeline[0]['$match']
        raw_count = collection.count_documents(raw_filter)
        print(f"Raw filter count (before aggregation): {raw_count}")
        
        return total_files, raw_filter, raw_count
    
    return total_files, None, 0

def investigate_pipeline_filter_utils():
    """
    Test the PipelineFilterUtils.get_both_filters() directly.
    """
    print("=" * 80)
    print("3. PIPELINE FILTER UTILS INVESTIGATION")
    print("=" * 80)
    
    try:
        from pipelines.utils import PipelineFilterUtils
        
        # Get the both_filters
        both_filters_stage = PipelineFilterUtils.get_both_filters()
        print("PipelineFilterUtils.get_both_filters():")
        print(f"  {both_filters_stage}")
        
        # Test this filter directly
        db = get_db_connection()
        collection = db.media_records
        
        filter_dict = both_filters_stage['$match']
        count = collection.count_documents(filter_dict)
        print(f"Direct filter count: {count}")
        
        return count, filter_dict
        
    except ImportError as e:
        print(f"❌ Could not import PipelineFilterUtils: {e}")
        return 0, None

def investigate_zero_fill_processing():
    """
    Check if zero-fill processing affects the totals.
    """
    print("=" * 80)
    print("4. ZERO-FILL PROCESSING INVESTIGATION")
    print("=" * 80)
    
    # Check if there's any zero-fill logic that might affect totals
    try:
        from ar_utils import precompute_collection_days
        
        collection_days = precompute_collection_days()
        print(f"Total collection days from ar_utils: {len(collection_days)}")
        
        # Check if zero-fill affects the count
        db = get_db_connection()
        collection = db.media_records
        
        # Count actual days with data
        pipeline = [
            {
                "$match": {
                    "School_Year": {"$ne": "N/A"},
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
        
        daily_results = list(collection.aggregate(pipeline))
        print(f"Days with actual data: {len(daily_results)}")
        
        total_from_daily = sum(result['Total_Files'] for result in daily_results)
        print(f"Total files from daily aggregation: {total_from_daily}")
        
        return total_from_daily, len(daily_results)
        
    except ImportError as e:
        print(f"❌ Could not import ar_utils: {e}")
        return 0, 0

def investigate_excel_sheet_logic():
    """
    Check if there's any Excel-specific processing that affects totals.
    """
    print("=" * 80)
    print("5. EXCEL SHEET GENERATION INVESTIGATION")
    print("=" * 80)
    
    # Look for any Excel-specific processing in the sheet creators
    try:
        # Check if there's any specific logic in the sheet creation that might filter data
        print("Checking for Excel-specific filtering logic...")
        
        # This would require examining the actual sheet creation code
        # For now, let's just note that this needs to be checked
        print("TODO: Examine sheet creation code for additional filtering")
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        return False

def compare_all_approaches():
    """
    Compare all different approaches and identify discrepancies.
    """
    print("=" * 80)
    print("6. COMPREHENSIVE COMPARISON")
    print("=" * 80)
    
    results = {}
    
    # 1. Data Cleaning approach
    data_cleaning_count, data_cleaning_filter = investigate_data_cleaning_logic()
    results['Data Cleaning'] = data_cleaning_count
    
    # 2. Daily Counts pipeline approach
    pipeline_total, pipeline_filter, pipeline_raw = investigate_daily_counts_pipeline()
    results['Daily Counts Pipeline (Aggregated)'] = pipeline_total
    results['Daily Counts Pipeline (Raw Filter)'] = pipeline_raw
    
    # 3. Pipeline Filter Utils approach
    filter_utils_count, filter_utils_filter = investigate_pipeline_filter_utils()
    results['Pipeline Filter Utils'] = filter_utils_count
    
    # 4. Zero-fill approach
    zero_fill_total, zero_fill_days = investigate_zero_fill_processing()
    results['Zero-fill Daily Aggregation'] = zero_fill_total
    
    # 5. Excel processing
    investigate_excel_sheet_logic()
    
    print("\nCOMPREHENSIVE RESULTS COMPARISON:")
    print("-" * 50)
    for approach, count in results.items():
        status = "✅ Match" if count == 9731 else f"⚠️  Diff: {count - 9731:+d}"
        print(f"{approach:<35}: {count:>6} {status}")
    
    # Identify the discrepancy
    print("\nDISCREPANCY ANALYSIS:")
    print("-" * 50)
    
    unique_counts = set(results.values())
    if len(unique_counts) == 1:
        print("✅ All approaches return the same count - no discrepancy found")
    else:
        print(f"⚠️  Found {len(unique_counts)} different counts: {sorted(unique_counts)}")
        
        # Find which approach gives 9372
        for approach, count in results.items():
            if count == 9372:
                print(f"🎯 FOUND: {approach} returns 9372 - this is the source!")
            elif count != 9731:
                print(f"❓ {approach} returns {count} - unexpected")
    
    return results

def main():
    print("COMPREHENSIVE DISCREPANCY INVESTIGATION")
    print("=" * 80)
    print("Investigating the root cause of 9,731 vs 9,372 discrepancy")
    print("=" * 80)
    
    try:
        results = compare_all_approaches()
        
        print("\n" + "=" * 80)
        print("INVESTIGATION COMPLETE")
        print("=" * 80)
        
        # Provide recommendations based on findings
        if 9372 in results.values():
            print("🎯 ROOT CAUSE IDENTIFIED!")
            print("The 9,372 count was found in one of the approaches.")
            print("This indicates where the discrepancy originates.")
        else:
            print("❓ ROOT CAUSE NOT FOUND in database/pipeline logic.")
            print("The discrepancy may be in Excel sheet generation or display logic.")
            print("Recommend examining the actual Excel file contents.")
        
    except Exception as e:
        print(f"Error during investigation: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
