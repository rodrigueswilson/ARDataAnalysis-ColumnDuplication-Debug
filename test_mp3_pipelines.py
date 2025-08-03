#!/usr/bin/env python3
"""
Test script to verify MP3 duration analysis pipelines are working correctly.
This will help us debug why the MP3 Duration Analysis sheet might still be incomplete.
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection
from pipelines.mp3_analysis import MP3_PIPELINES
import pandas as pd

def test_mp3_pipelines():
    """Test all MP3 duration analysis pipelines and show results."""
    
    print("=== MP3 Duration Analysis Pipeline Test ===")
    
    # Get database connection
    try:
        db = get_db_connection()
        collection = db.ar_data
        print(f"✓ Connected to database: {db.name}")
        print(f"✓ Using collection: {collection.name}")
    except Exception as e:
        print(f"✗ Database connection failed: {e}")
        return
    
    # Test basic data availability
    print("\n--- Basic Data Check ---")
    total_docs = collection.count_documents({})
    mp3_docs = collection.count_documents({"file_type": "MP3"})
    mp3_with_duration = collection.count_documents({
        "file_type": "MP3", 
        "Duration_Seconds": {"$exists": True, "$gt": 0}
    })
    mp3_collection_days = collection.count_documents({
        "file_type": "MP3", 
        "is_collection_day": True,
        "Duration_Seconds": {"$exists": True, "$gt": 0}
    })
    mp3_clean = collection.count_documents({
        "file_type": "MP3",
        "School_Year": {"$ne": "N/A"},
        "Duration_Seconds": {"$exists": True, "$gt": 0},
        "is_collection_day": True,
        "Outlier_Status": False
    })
    
    print(f"Total documents: {total_docs}")
    print(f"MP3 files: {mp3_docs}")
    print(f"MP3 with duration: {mp3_with_duration}")
    print(f"MP3 on collection days: {mp3_collection_days}")
    print(f"Clean MP3 data (final filter): {mp3_clean}")
    
    # Test each pipeline
    for pipeline_name, pipeline in MP3_PIPELINES.items():
        print(f"\n--- Testing {pipeline_name} ---")
        try:
            # Execute pipeline
            results = list(collection.aggregate(pipeline))
            print(f"✓ Pipeline executed successfully")
            print(f"✓ Returned {len(results)} results")
            
            if results:
                print("Sample result:")
                for key, value in results[0].items():
                    print(f"  {key}: {value}")
                    
                # Convert to DataFrame for inspection
                df = pd.DataFrame(results)
                print(f"DataFrame shape: {df.shape}")
                print(f"DataFrame columns: {list(df.columns)}")
            else:
                print("✗ No results returned")
                
        except Exception as e:
            print(f"✗ Pipeline failed: {e}")
    
    print("\n=== Test Complete ===")

if __name__ == "__main__":
    test_mp3_pipelines()
