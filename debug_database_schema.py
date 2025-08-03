#!/usr/bin/env python3
"""
Database schema investigation script to understand field names and data structure
"""

import sys
import os

# Add the project directory to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def investigate_database_schema():
    """Investigate the database schema to understand field names"""
    
    print("=== Database Schema Investigation ===")
    
    try:
        # Get database connection
        print("[1] Getting database connection...")
        db = get_db_connection()
        if db is None:
            print("❌ Failed to get database connection")
            return
        
        # Get collection
        collection = db['media_records']
        
        # Get total count
        total_count = collection.count_documents({})
        print(f"[2] Total documents in media_records: {total_count}")
        
        # Get a sample document to understand the schema
        print("[3] Getting sample document...")
        sample_doc = collection.find_one()
        
        if sample_doc:
            print("\n=== Sample Document Structure ===")
            for key, value in sample_doc.items():
                print(f"  {key}: {value} (type: {type(value).__name__})")
        
        # Check for specific fields we're using in the aggregation
        print("\n=== Field Existence Check ===")
        
        # Check School_Year field
        school_year_count = collection.count_documents({"School_Year": {"$exists": True}})
        print(f"Documents with School_Year field: {school_year_count}")
        
        # Check file_type field
        file_type_count = collection.count_documents({"file_type": {"$exists": True}})
        print(f"Documents with file_type field: {file_type_count}")
        
        # Check is_collection_day field
        collection_day_count = collection.count_documents({"is_collection_day": {"$exists": True}})
        print(f"Documents with is_collection_day field: {collection_day_count}")
        
        # Check Outlier_Status field
        outlier_status_count = collection.count_documents({"Outlier_Status": {"$exists": True}})
        print(f"Documents with Outlier_Status field: {outlier_status_count}")
        
        # Get unique values for key fields
        print("\n=== Unique Values Analysis ===")
        
        # School_Year values
        school_years = collection.distinct("School_Year")
        print(f"Unique School_Year values: {school_years}")
        
        # file_type values
        file_types = collection.distinct("file_type")
        print(f"Unique file_type values: {file_types}")
        
        # is_collection_day values
        collection_day_values = collection.distinct("is_collection_day")
        print(f"Unique is_collection_day values: {collection_day_values}")
        
        # Outlier_Status values
        outlier_status_values = collection.distinct("Outlier_Status")
        print(f"Unique Outlier_Status values: {outlier_status_values}")
        
        # Test the specific aggregation queries from the intersection analysis
        print("\n=== Testing Aggregation Queries ===")
        
        # Test raw data query
        raw_pipeline = [
            {"$match": {
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]}
            }},
            {"$group": {
                "_id": "$file_type",
                "count": {"$sum": 1}
            }}
        ]
        
        raw_result = list(collection.aggregate(raw_pipeline))
        print(f"Raw data aggregation result: {raw_result}")
        
        # Test collection days query
        collection_pipeline = [
            {"$match": {
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "is_collection_day": True
            }},
            {"$group": {
                "_id": "$file_type",
                "count": {"$sum": 1}
            }}
        ]
        
        collection_result = list(collection.aggregate(collection_pipeline))
        print(f"Collection days aggregation result: {collection_result}")
        
        # Test non-outliers query
        non_outlier_pipeline = [
            {"$match": {
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "Outlier_Status": False
            }},
            {"$group": {
                "_id": "$file_type",
                "count": {"$sum": 1}
            }}
        ]
        
        non_outlier_result = list(collection.aggregate(non_outlier_pipeline))
        print(f"Non-outliers aggregation result: {non_outlier_result}")
        
        # Test both criteria query
        both_pipeline = [
            {"$match": {
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "is_collection_day": True,
                "Outlier_Status": False
            }},
            {"$group": {
                "_id": "$file_type",
                "count": {"$sum": 1}
            }}
        ]
        
        both_result = list(collection.aggregate(both_pipeline))
        print(f"Both criteria aggregation result: {both_result}")
        
        # Summary
        print("\n=== Summary ===")
        if not raw_result:
            print("❌ Raw data query returns no results - check field names and data")
        else:
            print("✅ Raw data query works")
            
        if not collection_result:
            print("❌ Collection days query returns no results - check is_collection_day field")
        else:
            print("✅ Collection days query works")
            
        if not non_outlier_result:
            print("❌ Non-outliers query returns no results - check Outlier_Status field")
        else:
            print("✅ Non-outliers query works")
            
        if not both_result:
            print("❌ Both criteria query returns no results")
        else:
            print("✅ Both criteria query works")
        
    except Exception as e:
        print(f"❌ Error during investigation: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    investigate_database_schema()
