#!/usr/bin/env python3
"""
Debug script to examine MP3 duration pipeline data.
"""

import pandas as pd
from db_utils import get_db_connection
from pipelines import PIPELINES

def debug_mp3_pipelines():
    print("=== MP3 DURATION PIPELINE DATA DEBUG ===")
    
    try:
        # Connect to database
        db = get_db_connection()
        collection = db['media_records']  # Use correct collection name
        
        # Test each pipeline and examine the data structure
        pipelines_to_test = [
            ('MP3_DURATION_BY_SCHOOL_YEAR', 'School Year Analysis'),
            ('MP3_DURATION_BY_PERIOD', 'Collection Period Analysis'),
            ('MP3_DURATION_BY_MONTH', 'Monthly Analysis')
        ]
        
        for pipeline_name, description in pipelines_to_test:
            print(f"\n{'='*60}")
            print(f"TESTING: {description} ({pipeline_name})")
            print(f"{'='*60}")
            
            pipeline = PIPELINES[pipeline_name]
            print(f"Pipeline stages: {len(pipeline)}")
            
            # Execute pipeline
            cursor = collection.aggregate(pipeline)
            results = list(cursor)
            
            print(f"Results count: {len(results)}")
            
            if results:
                print(f"\nFirst result structure:")
                first_result = results[0]
                for key, value in first_result.items():
                    print(f"  {key}: {value} (type: {type(value)})")
                
                print(f"\nAll results:")
                for i, result in enumerate(results):
                    print(f"  [{i}]: {result}")
            else:
                print("  No results returned!")
        
        # Also check what Collection_Period values exist in the database
        print(f"\n{'='*60}")
        print("CHECKING COLLECTION_PERIOD VALUES IN DATABASE")
        print(f"{'='*60}")
        
        period_values = collection.distinct("Collection_Period")
        print(f"Distinct Collection_Period values: {period_values}")
        
        # Check School_Year values
        school_year_values = collection.distinct("School_Year")
        print(f"Distinct School_Year values: {school_year_values}")
        
        # Check sample MP3 records
        print(f"\n{'='*60}")
        print("SAMPLE MP3 RECORDS")
        print(f"{'='*60}")
        
        sample_mp3 = collection.find({
            "file_type": "MP3",
            "Duration_Seconds": {"$exists": True, "$gt": 0}
        }).limit(3)
        
        for i, record in enumerate(sample_mp3):
            print(f"\nSample MP3 record {i+1}:")
            relevant_fields = ['file_type', 'School_Year', 'Collection_Period', 'Duration_Seconds', 'ISO_Date']
            for field in relevant_fields:
                print(f"  {field}: {record.get(field, 'NOT_FOUND')}")
                
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_mp3_pipelines()
