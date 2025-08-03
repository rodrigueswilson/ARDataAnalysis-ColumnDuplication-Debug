#!/usr/bin/env python3
"""
Debug script to investigate the inconsistency between the two Data Cleaning tables
"""

import sys
import os

# Add the project directory to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def debug_table_inconsistency():
    """Debug the inconsistency between intersection table and year-by-year table"""
    
    print("=== Table Inconsistency Debug ===")
    
    try:
        # Get database connection
        db = get_db_connection()
        collection = db['media_records']
        
        print("\n=== Overall Totals ===")
        
        # Total files (excluding N/A school years)
        total_files = collection.count_documents({
            "School_Year": {"$ne": "N/A"}
        })
        print(f"Total files (excluding N/A): {total_files}")
        
        # Files by type
        jpg_total = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "file_type": "JPG"
        })
        mp3_total = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "file_type": "MP3"
        })
        print(f"JPG files: {jpg_total}")
        print(f"MP3 files: {mp3_total}")
        print(f"Sum: {jpg_total + mp3_total}")
        
        print("\n=== Table 1 Logic: Intersection Analysis ===")
        print("Final dataset = is_collection_day: TRUE AND Outlier_Status: FALSE")
        
        # Table 1: Final dataset (both criteria)
        table1_final = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        table1_excluded = total_files - table1_final
        print(f"Table 1 final dataset: {table1_final}")
        print(f"Table 1 excluded: {table1_excluded}")
        
        print("\n=== Table 2 Logic: Outlier Removal Only ===")
        print("Final dataset = Outlier_Status: FALSE (regardless of collection day)")
        
        # Table 2: Final dataset (outlier removal only)
        table2_final = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "Outlier_Status": False
        })
        table2_excluded = total_files - table2_final
        print(f"Table 2 final dataset: {table2_final}")
        print(f"Table 2 excluded (outliers only): {table2_excluded}")
        
        print("\n=== Breakdown of Excluded Files ===")
        
        # Files excluded by collection day criterion only
        collection_excluded = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": False,
            "Outlier_Status": False
        })
        print(f"Excluded by collection day only (non-school normal): {collection_excluded}")
        
        # Files excluded by outlier criterion only
        outlier_excluded = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "Outlier_Status": True
        })
        print(f"Excluded by outlier criterion (all outliers): {outlier_excluded}")
        
        # Files excluded by both criteria
        both_excluded = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": False,
            "Outlier_Status": True
        })
        print(f"Excluded by both criteria (non-school outliers): {both_excluded}")
        
        print("\n=== Year-by-Year Breakdown ===")
        
        years = ["2021-2022", "2022-2023"]
        for year in years:
            print(f"\n--- {year} ---")
            
            for file_type in ["JPG", "MP3"]:
                # Total for this year/type
                year_total = collection.count_documents({
                    "School_Year": year,
                    "file_type": file_type
                })
                
                # Outliers for this year/type
                year_outliers = collection.count_documents({
                    "School_Year": year,
                    "file_type": file_type,
                    "Outlier_Status": True
                })
                
                # After outlier removal
                year_after_outliers = year_total - year_outliers
                
                print(f"{file_type}: {year_total} total, {year_outliers} outliers, {year_after_outliers} after cleaning")
        
        print("\n=== Summary ===")
        print(f"Table 1 approach excludes {table1_excluded} files (intersection logic)")
        print(f"Table 2 approach excludes {table2_excluded} files (outlier-only logic)")
        print(f"Difference: {table1_excluded - table2_excluded} files")
        print(f"This difference should equal 'Non-School Normal' files: {collection_excluded}")
        
    except Exception as e:
        print(f"❌ Error during debug: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_table_inconsistency()
