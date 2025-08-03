#!/usr/bin/env python3
"""
Debug Zero-Fill Execution

This script investigates why our zero-fill fix isn't working by:
1. Testing if the _fill_missing_collection_days method is being called
2. Checking if the pipeline name conditions are met
3. Verifying the debug messages are being triggered
4. Testing the actual zero-fill logic directly
"""

import sys
import os
import pandas as pd

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ar_utils import get_school_calendar, get_non_collection_days, precompute_collection_days
from db_utils import get_db_connection
from report_generator.sheet_creators.base import BaseSheetCreator

def test_pipeline_condition():
    """Test if the pipeline name condition is working correctly."""
    print("TESTING PIPELINE NAME CONDITIONS")
    print("=" * 50)
    
    # Test the exact pipeline name used for Daily Counts (ACF_PACF)
    pipeline_name = "DAILY_COUNTS_COLLECTION_ONLY"
    
    print(f"Pipeline name: {pipeline_name}")
    print(f"Upper case: {pipeline_name.upper()}")
    
    # Test the conditions
    has_daily = 'DAILY' in pipeline_name.upper()
    has_with_zeroes = 'WITH_ZEROES' in pipeline_name.upper()
    has_collection_only = 'COLLECTION_ONLY' in pipeline_name.upper()
    
    print(f"Has 'DAILY': {has_daily}")
    print(f"Has 'WITH_ZEROES': {has_with_zeroes}")
    print(f"Has 'COLLECTION_ONLY': {has_collection_only}")
    
    condition_met = has_daily and (has_with_zeroes or has_collection_only)
    print(f"Condition met: {condition_met}")
    
    if condition_met:
        print("✅ Pipeline should trigger zero-fill logic")
    else:
        print("❌ Pipeline will NOT trigger zero-fill logic")
        print("This explains why our fix isn't working!")
    
    return condition_met

def test_zero_fill_method_directly():
    """Test the _fill_missing_collection_days method directly."""
    print("\n" + "=" * 50)
    print("TESTING ZERO-FILL METHOD DIRECTLY")
    print("=" * 50)
    
    try:
        # Create a BaseSheetCreator instance
        sheet_creator = BaseSheetCreator()
        
        # Get sample data from the database
        db = get_db_connection()
        collection = db.media_records
        
        # Get a small sample of data
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
            },
            {
                "$sort": {"_id": 1}
            },
            {
                "$limit": 10  # Just get first 10 days for testing
            }
        ]
        
        sample_data = list(collection.aggregate(pipeline))
        sample_df = pd.DataFrame(sample_data)
        
        print(f"Sample data: {len(sample_df)} rows")
        print("Sample dates:")
        for _, row in sample_df.head().iterrows():
            print(f"  {row['_id']}: {row['Total_Files']} files")
        
        # Test with the exact pipeline name
        pipeline_name = "DAILY_COUNTS_COLLECTION_ONLY"
        print(f"\nTesting with pipeline name: {pipeline_name}")
        
        # Call the method
        result_df = sheet_creator._fill_missing_collection_days(sample_df, pipeline_name)
        
        print(f"Result: {len(result_df)} rows")
        print("First few rows of result:")
        for _, row in result_df.head(15).iterrows():
            print(f"  {row['_id']}: {row['Total_Files']} files")
        
        # Check if early September dates are included
        early_sept_dates = [
            "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
            "2021-09-20", "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
        ]
        
        print("\nEarly September dates in result:")
        for date_str in early_sept_dates:
            if date_str in result_df['_id'].values:
                row = result_df[result_df['_id'] == date_str]
                if not row.empty:
                    count = row.iloc[0]['Total_Files']
                    print(f"  ✅ {date_str}: {count} files")
            else:
                print(f"  ❌ {date_str}: Not found")
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def check_actual_pipeline_names():
    """Check what pipeline names are actually being used in the report generation."""
    print("\n" + "=" * 50)
    print("CHECKING ACTUAL PIPELINE NAMES")
    print("=" * 50)
    
    try:
        import json
        
        # Read the report configuration
        with open('report_config.json', 'r') as f:
            config = json.load(f)
        
        print("ACF/PACF sheet configurations:")
        for sheet_name, sheet_config in config.get('sheets', {}).items():
            if 'ACF_PACF' in sheet_name or 'acf_pacf' in sheet_name.lower():
                pipeline_name = sheet_config.get('pipeline', 'Unknown')
                print(f"  {sheet_name}: {pipeline_name}")
                
                # Test if this pipeline would trigger zero-fill
                has_daily = 'DAILY' in pipeline_name.upper()
                has_with_zeroes = 'WITH_ZEROES' in pipeline_name.upper()
                has_collection_only = 'COLLECTION_ONLY' in pipeline_name.upper()
                condition_met = has_daily and (has_with_zeroes or has_collection_only)
                
                status = "✅ WILL trigger zero-fill" if condition_met else "❌ Will NOT trigger zero-fill"
                print(f"    {status}")
        
        return True
        
    except Exception as e:
        print(f"Error reading config: {e}")
        return False

def main():
    print("DEBUG ZERO-FILL EXECUTION INVESTIGATION")
    print("=" * 60)
    
    try:
        # Test 1: Pipeline conditions
        condition_works = test_pipeline_condition()
        
        # Test 2: Actual pipeline names in config
        check_actual_pipeline_names()
        
        # Test 3: Direct method testing
        if condition_works:
            test_zero_fill_method_directly()
        else:
            print("\n❌ Skipping direct method test - condition not met")
        
        print("\n" + "=" * 60)
        print("INVESTIGATION CONCLUSION")
        print("=" * 60)
        
        if not condition_works:
            print("🎯 FOUND THE ISSUE!")
            print("The pipeline name 'DAILY_COUNTS_COLLECTION_ONLY' does NOT meet")
            print("the condition for zero-fill because it lacks 'WITH_ZEROES'.")
            print()
            print("SOLUTION:")
            print("1. Either modify the condition to include 'COLLECTION_ONLY' pipelines")
            print("2. Or change the pipeline names to include 'WITH_ZEROES'")
            print("3. Or create a separate zero-fill condition for ACF/PACF sheets")
        else:
            print("The pipeline condition is correct. Issue must be elsewhere.")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
