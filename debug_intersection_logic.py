#!/usr/bin/env python3
"""
Debug script to understand the intersection logic and fix the mathematical issue
"""

import sys
import os

# Add the project directory to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def debug_intersection_logic():
    """Debug the intersection logic to understand the mathematical issue"""
    
    print("=== Intersection Logic Debug ===")
    
    try:
        # Get database connection
        db = get_db_connection()
        collection = db['media_records']
        
        print("\n=== Raw Data Analysis ===")
        
        # For JPG files specifically, let's break down the logic
        file_type = "JPG"
        
        # 1. Total raw JPG files (excluding N/A school years)
        total_raw = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "file_type": file_type
        })
        print(f"Total {file_type} files: {total_raw}")
        
        # 2. Collection days (is_collection_day: TRUE)
        collection_days_total = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "file_type": file_type,
            "is_collection_day": True
        })
        print(f"Collection days {file_type} files: {collection_days_total}")
        
        # 3. Non-outliers (Outlier_Status: FALSE)
        non_outliers_total = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "file_type": file_type,
            "Outlier_Status": False
        })
        print(f"Non-outlier {file_type} files: {non_outliers_total}")
        
        # 4. Both criteria (is_collection_day: TRUE AND Outlier_Status: FALSE)
        both_criteria = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "file_type": file_type,
            "is_collection_day": True,
            "Outlier_Status": False
        })
        print(f"Both criteria {file_type} files: {both_criteria}")
        
        print("\n=== Venn Diagram Breakdown ===")
        
        # Calculate the four exclusive categories
        # A = Collection Days, B = Non-Outliers
        
        # A only: Collection days that ARE outliers
        collection_only_outliers = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "file_type": file_type,
            "is_collection_day": True,
            "Outlier_Status": True
        })
        print(f"Collection days only (outliers): {collection_only_outliers}")
        
        # B only: Non-outliers that are NOT collection days  
        non_collection_non_outliers = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "file_type": file_type,
            "is_collection_day": False,
            "Outlier_Status": False
        })
        print(f"Non-collection days, non-outliers: {non_collection_non_outliers}")
        
        # A ∩ B: Both criteria (already calculated)
        print(f"Both criteria (A ∩ B): {both_criteria}")
        
        # Neither: NOT collection days AND ARE outliers
        neither_actual = collection.count_documents({
            "School_Year": {"$ne": "N/A"},
            "file_type": file_type,
            "is_collection_day": False,
            "Outlier_Status": True
        })
        print(f"Neither (non-collection outliers): {neither_actual}")
        
        print("\n=== Verification ===")
        calculated_total = collection_only_outliers + non_collection_non_outliers + both_criteria + neither_actual
        print(f"Sum of all categories: {calculated_total}")
        print(f"Original total: {total_raw}")
        print(f"Match: {'✅' if calculated_total == total_raw else '❌'}")
        
        print("\n=== Current Formula vs. Correct Values ===")
        
        # Current incorrect formula
        current_neither = total_raw - collection_days_total - non_outliers_total + both_criteria
        print(f"Current 'neither' formula result: {current_neither}")
        print(f"Actual 'neither' from database: {neither_actual}")
        print(f"Formula is {'✅ Correct' if current_neither == neither_actual else '❌ Wrong'}")
        
        # Show the issue with current calculation
        print(f"\n=== Current Table Values (Incorrect) ===")
        current_collection_only = collection_days_total - both_criteria
        current_non_outlier_only = non_outliers_total - both_criteria
        
        print(f"School Days Only: {current_collection_only}")
        print(f"Normal Files Only: {current_non_outlier_only}")
        print(f"Clean Research Data: {both_criteria}")
        print(f"Excluded Files: {current_neither}")
        print(f"Total: {current_collection_only + current_non_outlier_only + both_criteria + current_neither}")
        
        print(f"\n=== Correct Table Values ===")
        print(f"School Days Only: {collection_only_outliers}")
        print(f"Normal Files Only: {non_collection_non_outliers}")
        print(f"Clean Research Data: {both_criteria}")
        print(f"Excluded Files: {neither_actual}")
        print(f"Total: {collection_only_outliers + non_collection_non_outliers + both_criteria + neither_actual}")
        
    except Exception as e:
        print(f"❌ Error during debug: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_intersection_logic()
