#!/usr/bin/env python3
"""
Check period data in MongoDB to understand why SY 22-23 P3 might be missing
"""

from db_utils import get_db_connection

def check_period_data():
    try:
        db = get_db_connection()
        
        # Check period data
        pipeline = [
            {'$match': {'file_type': 'MP3', 'is_collection_day': True, 'Outlier_Status': False}},
            {'$group': {
                '_id': {'School_Year': '$School_Year', 'Collection_Period': '$Collection_Period'},
                'Total_MP3_Files': {'$sum': 1},
                'Total_Duration_Hours': {'$sum': {'$divide': ['$Duration_Seconds', 3600]}}
            }},
            {'$sort': {'_id.School_Year': 1, '_id.Collection_Period': 1}}
        ]
        
        result = list(db.media_records.aggregate(pipeline))
        print('=== Period Data from MongoDB ===')
        for r in result:
            school_year = r['_id']['School_Year']
            period = r['_id']['Collection_Period']
            files = r['Total_MP3_Files']
            print(f'{school_year} {period}: {files} files')
        
        print(f'\nTotal periods found: {len(result)}')
        
        # Check specifically for SY 22-23 P3
        sy_22_23_periods = [r for r in result if r['_id']['School_Year'] == '2022-2023']
        print(f'\n2022-2023 periods found: {len(sy_22_23_periods)}')
        for r in sy_22_23_periods:
            period = r['_id']['Collection_Period']
            files = r['Total_MP3_Files']
            print(f'  {period}: {files} files')
            
    except Exception as e:
        print(f"Error checking period data: {e}")

if __name__ == "__main__":
    check_period_data()
