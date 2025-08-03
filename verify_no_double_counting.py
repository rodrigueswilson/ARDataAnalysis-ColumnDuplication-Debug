#!/usr/bin/env python3
"""
Verify that the intersection logic does not double-count any files
"""

import sys
import os

# Add the project directory to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def verify_no_double_counting():
    """Verify that each file is counted exactly once in the intersection analysis"""
    
    print("=== Double-Counting Verification ===")
    
    try:
        # Get database connection
        db = get_db_connection()
        collection = db['media_records']
        
        print("\n=== Total Files Check ===")
        
        # Total files (excluding N/A school years)
        total_files = collection.count_documents({
            "School_Year": {"$ne": "N/A"}
        })
        print(f"Total files: {total_files}")
        
        print("\n=== Four Mutually Exclusive Categories ===")
        
        # Category 1: School Outliers
        school_outliers = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": True,
            "Outlier_Status": True
        })
        print(f"1. School Outliers: {school_outliers}")
        
        # Category 2: Non-School Normal
        non_school_normal = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": False,
            "Outlier_Status": False
        })
        print(f"2. Non-School Normal: {non_school_normal}")
        
        # Category 3: School Normal (final dataset)
        school_normal = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        print(f"3. School Normal: {school_normal}")
        
        # Category 4: Non-School Outliers
        non_school_outliers = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": False,
            "Outlier_Status": True
        })
        print(f"4. Non-School Outliers: {non_school_outliers}")
        
        print("\n=== Verification ===")
        
        sum_categories = school_outliers + non_school_normal + school_normal + non_school_outliers
        print(f"Sum of all categories: {sum_categories}")
        print(f"Original total: {total_files}")
        print(f"Match: {'✅ Perfect' if sum_categories == total_files else '❌ Mismatch'}")
        
        if sum_categories == total_files:
            print("✅ NO DOUBLE-COUNTING: Each file is counted exactly once")
        else:
            print("❌ POTENTIAL ISSUE: Files may be missing or double-counted")
            print(f"Difference: {sum_categories - total_files}")
        
        print("\n=== Breakdown by File Type ===")
        
        for file_type in ["JPG", "MP3"]:
            print(f"\n--- {file_type} Files ---")
            
            # Total for this file type
            type_total = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type
            })
            
            # Each category for this file type
            type_school_outliers = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": True,
                "Outlier_Status": True
            })
            
            type_non_school_normal = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": False,
                "Outlier_Status": False
            })
            
            type_school_normal = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": True,
                "Outlier_Status": False
            })
            
            type_non_school_outliers = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": False,
                "Outlier_Status": True
            })
            
            type_sum = type_school_outliers + type_non_school_normal + type_school_normal + type_non_school_outliers
            
            print(f"Total {file_type}: {type_total}")
            print(f"School Outliers: {type_school_outliers}")
            print(f"Non-School Normal: {type_non_school_normal}")
            print(f"School Normal: {type_school_normal}")
            print(f"Non-School Outliers: {type_non_school_outliers}")
            print(f"Sum: {type_sum}")
            print(f"Match: {'✅' if type_sum == type_total else '❌'}")
        
        print("\n=== Logic Verification ===")
        print("The intersection logic uses mutually exclusive conditions:")
        print("• Each file has exactly one value for 'is_collection_day' (True or False)")
        print("• Each file has exactly one value for 'Outlier_Status' (True or False)")
        print("• This creates exactly 4 possible combinations (2×2 matrix)")
        print("• No file can belong to multiple categories")
        print("• No file can be excluded twice")
        
    except Exception as e:
        print(f"❌ Error during verification: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    verify_no_double_counting()
