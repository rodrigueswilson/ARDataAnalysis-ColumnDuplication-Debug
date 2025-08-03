#!/usr/bin/env python3
"""
Verify the new 2x2 matrix table structure showing all four mutually exclusive categories
"""

import sys
import os

# Add the project directory to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def show_new_table_structure():
    """Show what the new table structure looks like with actual data"""
    
    print("=== New 2x2 Matrix Table Structure ===")
    
    try:
        # Get database connection
        db = get_db_connection()
        collection = db['media_records']
        
        print("\n📊 Table: Complete Filtering Breakdown (2x2 Matrix)")
        print("=" * 120)
        
        # Headers
        headers = ['File Type', 'Total Files', 'School Outliers', 'Non-School Normal', 'School Normal', 'Non-School Outliers', 'Total Excluded', '% Kept']
        print(f"{'|'.join(f'{h:>15}' for h in headers)}")
        print("-" * 120)
        
        # Calculate data for each file type
        file_types = ['JPG', 'MP3']
        table_data = []
        
        for file_type in file_types:
            # Total files
            total_files = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type
            })
            
            # School Outliers: is_collection_day: TRUE AND Outlier_Status: TRUE
            school_outliers = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": True,
                "Outlier_Status": True
            })
            
            # Non-School Normal: is_collection_day: FALSE AND Outlier_Status: FALSE
            non_school_normal = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": False,
                "Outlier_Status": False
            })
            
            # School Normal: is_collection_day: TRUE AND Outlier_Status: FALSE
            school_normal = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": True,
                "Outlier_Status": False
            })
            
            # Non-School Outliers: is_collection_day: FALSE AND Outlier_Status: TRUE
            non_school_outliers = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": False,
                "Outlier_Status": True
            })
            
            # Total excluded (all categories except School Normal)
            total_excluded = school_outliers + non_school_normal + non_school_outliers
            
            # Retention percentage
            retention_pct = (school_normal / total_files * 100) if total_files > 0 else 0
            
            # Store data
            row_data = [file_type, total_files, school_outliers, non_school_normal, school_normal, non_school_outliers, total_excluded, f"{retention_pct:.1f}%"]
            table_data.append(row_data)
            
            # Print row
            print(f"{'|'.join(f'{str(val):>15}' for val in row_data)}")
        
        # Calculate totals
        totals = ['TOTAL']
        for i in range(1, 7):  # Skip file type column and % kept column
            total_val = sum(row[i] for row in table_data)
            totals.append(total_val)
        
        # Calculate total retention percentage
        total_retention = (totals[4] / totals[1] * 100) if totals[1] > 0 else 0  # school_normal / total_files
        totals.append(f"{total_retention:.1f}%")
        
        print("-" * 120)
        print(f"{'|'.join(f'{str(val):>15}' for val in totals)}")
        
        print("\n🎯 Key Insights:")
        print(f"• Total files: {totals[1]:,}")
        print(f"• School Normal (final dataset): {totals[4]:,} files ({total_retention:.1f}% retention)")
        print(f"• Total excluded: {totals[6]:,} files")
        print(f"  - School Outliers: {totals[2]:,}")
        print(f"  - Non-School Normal: {totals[3]:,}")
        print(f"  - Non-School Outliers: {totals[5]:,}")
        
        print("\n📋 Category Definitions:")
        print("• School Outliers: Unusual files from school collection days")
        print("• Non-School Normal: Regular files from non-school days")
        print("• School Normal: Regular files from school collection days (FINAL DATASET)")
        print("• Non-School Outliers: Unusual files from non-school days")
        
        print("\n✅ Verification:")
        calculated_total = totals[2] + totals[3] + totals[4] + totals[5]
        print(f"Sum of all categories: {calculated_total:,}")
        print(f"Original total: {totals[1]:,}")
        print(f"Match: {'✅ Perfect' if calculated_total == totals[1] else '❌ Error'}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    show_new_table_structure()
