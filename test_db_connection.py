#!/usr/bin/env python3
"""
Test script to check MongoDB connection and examine the actual data structure.
"""

from db_utils import get_db_connection
from collections import Counter
import datetime

def test_database_connection():
    """Test MongoDB connection and examine data structure."""
    
    try:
        # Connect to MongoDB using the correct database connection
        db = get_db_connection()
        collection = db["media_records"]
        
        print("=== Database Connection Test ===")
        print(f"Connected to MongoDB successfully")
        
        # Check total record count
        total_count = collection.count_documents({})
        print(f"Total records in collection: {total_count}")
        
        # Check available collection periods
        periods = collection.distinct("Collection_Period")
        print(f"Available Collection_Period values: {periods}")
        
        # Check for SY 22-23 P1 with different queries
        print(f"\n=== SY 22-23 P1 Investigation ===")
        
        # Try exact match
        exact_count = collection.count_documents({"Collection_Period": "SY 22-23 P1"})
        print(f"Exact match 'SY 22-23 P1': {exact_count} records")
        
        # Try case-insensitive search
        regex_count = collection.count_documents({"Collection_Period": {"$regex": "SY 22-23 P1", "$options": "i"}})
        print(f"Case-insensitive match: {regex_count} records")
        
        # Try partial match
        partial_count = collection.count_documents({"Collection_Period": {"$regex": "22-23 P1"}})
        print(f"Partial match '22-23 P1': {partial_count} records")
        
        # Show sample records for any 22-23 data
        sample_records = list(collection.find(
            {"Collection_Period": {"$regex": "22-23"}}, 
            {"ISO_Date": 1, "Collection_Period": 1, "School_Year": 1, "_id": 0}
        ).limit(10))
        
        print(f"\n=== Sample 22-23 Records ===")
        for record in sample_records:
            print(f"  Period: '{record.get('Collection_Period')}', Date: {record.get('ISO_Date')}, Year: {record.get('School_Year')}")
        
        # Check if there are any records with dates in the expected range
        date_range_count = collection.count_documents({
            "ISO_Date": {"$gte": "2022-09-12", "$lte": "2022-12-12"}
        })
        print(f"\nRecords with dates 2022-09-12 to 2022-12-12: {date_range_count}")
        
        # Show sample records in that date range
        date_range_records = list(collection.find(
            {"ISO_Date": {"$gte": "2022-09-12", "$lte": "2022-12-12"}},
            {"ISO_Date": 1, "Collection_Period": 1, "School_Year": 1, "_id": 0}
        ).limit(10))
        
        print(f"Sample records in date range:")
        for record in date_range_records:
            print(f"  Period: '{record.get('Collection_Period')}', Date: {record.get('ISO_Date')}, Year: {record.get('School_Year')}")
        
        # Close connection
        client.close()
        
    except Exception as e:
        print(f"Error connecting to database: {e}")

if __name__ == "__main__":
    test_database_connection()
