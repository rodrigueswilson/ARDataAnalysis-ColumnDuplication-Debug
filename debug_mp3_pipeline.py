#!/usr/bin/env python3
"""
Debug script to examine MP3 period pipeline output
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from pymongo import MongoClient
from pipelines.mp3_analysis import MP3_DURATION_BY_PERIOD
import pandas as pd

def debug_mp3_pipeline():
    """Debug the MP3 period pipeline to see what data it returns"""
    try:
        # Connect to MongoDB
        client = MongoClient('mongodb://localhost:27017/')
        db = client['ARDataAnalysis']
        collection = db['AudioRecordings']
        
        print("=== MP3 PERIOD PIPELINE DEBUG ===")
        print(f"Pipeline stages: {len(MP3_DURATION_BY_PERIOD)}")
        
        # Run the pipeline
        print("\nExecuting pipeline...")
        results = list(collection.aggregate(MP3_DURATION_BY_PERIOD))
        
        print(f"Pipeline returned {len(results)} documents")
        
        if results:
            print("\nFirst few results:")
            for i, doc in enumerate(results[:5]):
                print(f"Document {i+1}:")
                for key, value in doc.items():
                    print(f"  {key}: {value} (type: {type(value).__name__})")
                print()
        else:
            print("No results returned from pipeline")
            
        # Convert to DataFrame to see structure
        if results:
            df = pd.DataFrame(results)
            print(f"DataFrame shape: {df.shape}")
            print(f"DataFrame columns: {list(df.columns)}")
            
            # Check for null values in key fields
            if 'Period' in df.columns:
                null_periods = df['Period'].isnull().sum()
                print(f"Null Period values: {null_periods}")
                unique_periods = df['Period'].unique()
                print(f"Unique Period values: {unique_periods}")
                
            if 'School_Year' in df.columns:
                null_years = df['School_Year'].isnull().sum()
                print(f"Null School_Year values: {null_years}")
                unique_years = df['School_Year'].unique()
                print(f"Unique School_Year values: {unique_years}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_mp3_pipeline()
