#!/usr/bin/env python3
"""
Get the exact field names for MP3 records.
"""

from db_utils import get_db_connection

def get_mp3_field_names():
    print("=== MP3 FIELD NAMES ANALYSIS ===")
    
    try:
        # Connect to database
        db = get_db_connection()
        collection = db['media_records']
        
        # Get one MP3 record to see all field names
        sample_mp3 = collection.find_one({"file_type": "MP3"})
        
        if sample_mp3:
            print("All field names in MP3 record:")
            field_names = list(sample_mp3.keys())
            for i, field in enumerate(field_names):
                print(f"  {i+1:2d}. {field}")
            
            print(f"\nTotal fields: {len(field_names)}")
            
            # Check specific fields we need
            print(f"\n=== KEY FIELD VALUES ===")
            key_fields = ['School_Year', 'Collection_Period', 'Duration_Seconds', 'ISO_Date']
            
            for field in key_fields:
                if field in sample_mp3:
                    print(f"{field}: {sample_mp3[field]} (type: {type(sample_mp3[field])})")
                else:
                    print(f"{field}: FIELD NOT FOUND")
            
            # Check for alternative field names
            print(f"\n=== SEARCHING FOR ALTERNATIVE FIELD NAMES ===")
            alternative_searches = {
                'Collection_Period': ['period', 'collection'],
                'Duration_Seconds': ['duration', 'time', 'length'],
                'School_Year': ['school', 'year', 'sy']
            }
            
            for target_field, search_terms in alternative_searches.items():
                print(f"\nLooking for {target_field} alternatives:")
                found_alternatives = []
                for field in field_names:
                    if any(term.lower() in field.lower() for term in search_terms):
                        found_alternatives.append(field)
                        print(f"  - {field}: {sample_mp3[field]}")
                
                if not found_alternatives:
                    print(f"  No alternatives found for {target_field}")
                    
        else:
            print("No MP3 records found!")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    get_mp3_field_names()
