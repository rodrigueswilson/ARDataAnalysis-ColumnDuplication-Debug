#!/usr/bin/env python3
"""
Clean Database File Type Check

Simple script to check what file types are actually in the database
and identify any unexpected file types.
"""

import sys
import os
from pymongo import MongoClient

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def check_database_file_types():
    """Check what file types are in the database."""
    print("DATABASE FILE TYPE ANALYSIS")
    print("=" * 50)
    
    try:
        db = get_db_connection()
        collection = db.media_records
        
        print(f"Connected to database: {db.name}")
        print(f"Collection: {collection.name}")
        print(f"Total documents: {collection.count_documents({})}")
        print()
        
        # Get all unique file types
        pipeline = [
            {"$group": {"_id": "$file_type", "count": {"$sum": 1}}},
            {"$sort": {"count": -1}}
        ]
        
        file_types = list(collection.aggregate(pipeline))
        
        print("File types in database:")
        total_files = 0
        expected_types = {'JPG', 'MP3'}
        unexpected_types = []
        
        for item in file_types:
            file_type = item['_id'] if item['_id'] else 'NULL'
            count = item['count']
            total_files += count
            
            status = "✅ Expected" if file_type in expected_types else "⚠️  UNEXPECTED"
            print(f"  {file_type}: {count} files - {status}")
            
            if file_type not in expected_types:
                unexpected_types.append(file_type)
        
        print(f"\nTotal files: {total_files}")
        print(f"Expected file types (JPG, MP3): {len([ft for ft in file_types if ft['_id'] in expected_types])}")
        print(f"Unexpected file types: {len(unexpected_types)}")
        
        if unexpected_types:
            print(f"\n⚠️  UNEXPECTED FILE TYPES FOUND:")
            for file_type in unexpected_types:
                print(f"    - {file_type}")
                
                # Show a few examples
                examples = list(collection.find({"file_type": file_type}).limit(3))
                for i, example in enumerate(examples, 1):
                    file_path = example.get('File_Path', 'Unknown path')
                    print(f"      Example {i}: {file_path}")
            
            print(f"\n🔍 RECOMMENDATION:")
            print(f"   These unexpected file types suggest either:")
            print(f"   1. Data directories contain non-JPG/MP3 files")
            print(f"   2. Database population script is including wrong files")
            print(f"   3. File type detection logic has issues")
        else:
            print(f"\n✅ GOOD: Only expected file types (JPG, MP3) found in database")
        
        return file_types, unexpected_types
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return [], []

def check_clean_data_file_types():
    """Check file types in the clean dataset specifically."""
    print("\n" + "=" * 50)
    print("CLEAN DATA FILE TYPE ANALYSIS")
    print("=" * 50)
    
    try:
        db = get_db_connection()
        collection = db.media_records
        
        # Clean data filter (School Normal)
        clean_filter = {
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": True,
            "Outlier_Status": False
        }
        
        clean_count = collection.count_documents(clean_filter)
        print(f"Total clean files: {clean_count}")
        
        # Get file types in clean dataset
        pipeline = [
            {"$match": clean_filter},
            {"$group": {"_id": "$file_type", "count": {"$sum": 1}}},
            {"$sort": {"count": -1}}
        ]
        
        clean_file_types = list(collection.aggregate(pipeline))
        
        print("File types in clean dataset:")
        clean_total = 0
        for item in clean_file_types:
            file_type = item['_id'] if item['_id'] else 'NULL'
            count = item['count']
            clean_total += count
            
            status = "✅ Expected" if file_type in {'JPG', 'MP3'} else "⚠️  UNEXPECTED"
            print(f"  {file_type}: {count} files - {status}")
        
        print(f"Clean dataset total: {clean_total}")
        
        return clean_file_types
        
    except Exception as e:
        print(f"Error: {e}")
        return []

if __name__ == "__main__":
    print("CHECKING DATABASE FILE TYPES")
    print("=" * 80)
    
    # Check all file types
    all_file_types, unexpected = check_database_file_types()
    
    # Check clean data file types
    clean_file_types = check_clean_data_file_types()
    
    print("\n" + "=" * 50)
    print("SUMMARY")
    print("=" * 50)
    
    if unexpected:
        print(f"❌ ISSUE FOUND: {len(unexpected)} unexpected file types in database")
        print(f"   This explains potential discrepancies in filtering logic")
        print(f"   Unexpected types: {unexpected}")
    else:
        print(f"✅ NO ISSUES: Only JPG and MP3 files found in database")
        print(f"   File type filtering should work consistently")
    
    print("\nInvestigation complete!")
