#!/usr/bin/env python3
"""
Check the actual database schema and field names.
"""

from db_utils import get_db_connection

def check_database_schema():
    print("=== DATABASE SCHEMA ANALYSIS ===")
    
    try:
        # Connect to database
        db = get_db_connection()
        collection = db['media_records']  # Use correct collection name
        
        # Get total count
        total_count = collection.count_documents({})
        print(f"Total documents: {total_count}")
        
        # Get MP3 count
        mp3_count = collection.count_documents({"file_type": "MP3"})
        print(f"MP3 documents: {mp3_count}")
        
        # Get a sample MP3 record to see all fields
        print(f"\n{'='*60}")
        print("SAMPLE MP3 RECORD - ALL FIELDS")
        print(f"{'='*60}")
        
        sample_mp3 = collection.find_one({"file_type": "MP3"})
        if sample_mp3:
            for key, value in sample_mp3.items():
                print(f"  {key}: {value} (type: {type(value)})")
        else:
            print("No MP3 records found!")
            return
        
        # Check for duration-related fields
        print(f"\n{'='*60}")
        print("DURATION-RELATED FIELD ANALYSIS")
        print(f"{'='*60}")
        
        duration_fields = [field for field in sample_mp3.keys() if 'duration' in field.lower() or 'time' in field.lower()]
        print(f"Duration-related fields: {duration_fields}")
        
        # Check for school year related fields
        print(f"\n{'='*60}")
        print("SCHOOL YEAR RELATED FIELD ANALYSIS")
        print(f"{'='*60}")
        
        school_fields = [field for field in sample_mp3.keys() if 'school' in field.lower() or 'year' in field.lower()]
        print(f"School/Year related fields: {school_fields}")
        
        # Check for period related fields
        print(f"\n{'='*60}")
        print("PERIOD RELATED FIELD ANALYSIS")
        print(f"{'='*60}")
        
        period_fields = [field for field in sample_mp3.keys() if 'period' in field.lower() or 'collection' in field.lower()]
        print(f"Period/Collection related fields: {period_fields}")
        
        # Check distinct values for key fields
        print(f"\n{'='*60}")
        print("DISTINCT VALUES FOR KEY FIELDS")
        print(f"{'='*60}")
        
        key_fields_to_check = []
        
        # Add any fields that might contain school year info
        for field in sample_mp3.keys():
            if any(keyword in field.lower() for keyword in ['school', 'year', 'period', 'collection', 'duration']):
                key_fields_to_check.append(field)
        
        for field in key_fields_to_check:
            try:
                distinct_values = collection.distinct(field)
                print(f"  {field}: {distinct_values[:10]}{'...' if len(distinct_values) > 10 else ''} (total: {len(distinct_values)})")
            except Exception as e:
                print(f"  {field}: Error getting distinct values - {e}")
                
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_database_schema()
