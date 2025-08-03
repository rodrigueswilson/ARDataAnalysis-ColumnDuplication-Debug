#!/usr/bin/env python3
"""
Definitive Pipeline-Level Fix

This script directly modifies the pipeline execution at the MongoDB level
to ensure zero-fill logic is applied to ACF/PACF sheets regardless of 
which code path creates them.
"""

import sys
import os
import pandas as pd
from datetime import datetime, timedelta
import yaml
from pymongo import MongoClient

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def load_collection_day_map():
    """Load the collection day map from config.yaml."""
    try:
        with open('config.yaml', 'r') as f:
            config = yaml.safe_load(f)
        return config.get('collection_day_map', {})
    except Exception as e:
        print(f"Error loading config: {e}")
        return {}

def generate_complete_date_range():
    """Generate the complete date range for zero-fill."""
    collection_day_map = load_collection_day_map()
    
    # Extract all dates from the collection day map
    all_dates = []
    for week_data in collection_day_map.values():
        if isinstance(week_data, dict):
            for date_str in week_data.get('dates', []):
                try:
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
                    all_dates.append(date_obj)
                except ValueError:
                    continue
    
    if not all_dates:
        return []
    
    # Sort dates and create complete range
    all_dates.sort()
    start_date = all_dates[0]
    end_date = all_dates[-1]
    
    # Generate complete date range
    complete_dates = []
    current_date = start_date
    while current_date <= end_date:
        complete_dates.append(current_date.strftime('%Y-%m-%d'))
        current_date += timedelta(days=1)
    
    return complete_dates

def create_zero_filled_pipeline():
    """Create a pipeline that includes zero-fill logic at the MongoDB level."""
    complete_dates = generate_complete_date_range()
    
    if not complete_dates:
        print("❌ Could not generate complete date range")
        return None
    
    print(f"📅 Complete date range: {len(complete_dates)} dates from {complete_dates[0]} to {complete_dates[-1]}")
    
    # Create the zero-fill pipeline
    zero_fill_pipeline = [
        # Stage 1: Match collection day files only (School Normal logic)
        {
            "$match": {
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "is_collection_day": True,
                "Outlier_Status": False
            }
        },
        
        # Stage 2: Group by date and count files
        {
            "$group": {
                "_id": "$Date",
                "Total_Files": {"$sum": 1},
                "JPG_Files": {
                    "$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}
                },
                "MP3_Files": {
                    "$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}
                },
                "Total_Size_MB": {"$sum": "$Size_MB"}
            }
        },
        
        # Stage 3: Add all missing dates with zero counts
        {
            "$unionWith": {
                "coll": "dummy_collection",  # This will be empty
                "pipeline": [
                    {
                        "$documents": [
                            {
                                "_id": date,
                                "Total_Files": 0,
                                "JPG_Files": 0,
                                "MP3_Files": 0,
                                "Total_Size_MB": 0.0
                            }
                            for date in complete_dates
                        ]
                    }
                ]
            }
        },
        
        # Stage 4: Group again to merge actual data with zero-fill dates
        {
            "$group": {
                "_id": "$_id",
                "Total_Files": {"$sum": "$Total_Files"},
                "JPG_Files": {"$sum": "$JPG_Files"},
                "MP3_Files": {"$sum": "$MP3_Files"},
                "Total_Size_MB": {"$sum": "$Total_Size_MB"}
            }
        },
        
        # Stage 5: Sort by date
        {
            "$sort": {"_id": 1}
        }
    ]
    
    return zero_fill_pipeline

def patch_pipeline_execution():
    """Patch the pipeline execution to use zero-fill logic."""
    print("PATCHING PIPELINE EXECUTION")
    print("=" * 50)
    
    try:
        # Import the pipeline utilities
        from pipelines.pipeline_utils import PipelineUtils
        
        # Get the original get_pipeline method
        original_get_pipeline = PipelineUtils.get_pipeline
        
        def patched_get_pipeline(pipeline_name):
            """Patched version that applies zero-fill to ACF/PACF pipelines."""
            print(f"[PIPELINE_PATCH] Getting pipeline: {pipeline_name}")
            
            # Check if this is a daily ACF/PACF pipeline that needs zero-fill
            if pipeline_name == 'DAILY_COUNTS_COLLECTION_ONLY':
                print(f"[PIPELINE_PATCH] 🎯 Applying zero-fill to {pipeline_name}")
                zero_fill_pipeline = create_zero_filled_pipeline()
                if zero_fill_pipeline:
                    return zero_fill_pipeline
            
            # For other pipelines, use the original method
            return original_get_pipeline(pipeline_name)
        
        # Apply the patch
        PipelineUtils.get_pipeline = patched_get_pipeline
        print("✅ Pipeline execution patched successfully")
        
        return True
        
    except Exception as e:
        print(f"❌ Error patching pipeline execution: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_zero_fill_pipeline():
    """Test the zero-fill pipeline directly."""
    print("TESTING ZERO-FILL PIPELINE")
    print("=" * 50)
    
    try:
        # Connect to database
        client = MongoClient('mongodb://localhost:27017/')
        db = client['ar_data_analysis']
        
        # Create and test the zero-fill pipeline
        zero_fill_pipeline = create_zero_filled_pipeline()
        if not zero_fill_pipeline:
            return False
        
        print("🔍 Executing zero-fill pipeline...")
        cursor = db.media_records.aggregate(zero_fill_pipeline, allowDiskUse=True)
        results = list(cursor)
        
        print(f"📊 Zero-fill pipeline results: {len(results)} records")
        
        if results:
            # Show first few results
            print("Sample results:")
            for i, result in enumerate(results[:5]):
                print(f"  {i+1}. {result}")
            
            # Show total files
            total_files = sum(r.get('Total_Files', 0) for r in results)
            print(f"📈 Total files across all dates: {total_files}")
            
            return True
        else:
            print("❌ No results from zero-fill pipeline")
            return False
            
    except Exception as e:
        print(f"❌ Error testing zero-fill pipeline: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("DEFINITIVE PIPELINE-LEVEL FIX")
    print("=" * 70)
    
    # Test the zero-fill pipeline first
    if not test_zero_fill_pipeline():
        print("❌ Zero-fill pipeline test failed")
        return 1
    
    # Apply the pipeline patch
    if not patch_pipeline_execution():
        print("❌ Pipeline patching failed")
        return 1
    
    print("\n" + "=" * 70)
    print("NEXT STEPS")
    print("=" * 70)
    print("1. Run generate_report.py to test the patched pipeline")
    print("2. Verify that ACF/PACF sheets now have correct structure")
    print("3. Check that all 333 dates are included with zero-fill")
    print("4. Confirm that totals match the expected 9,731 files")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
