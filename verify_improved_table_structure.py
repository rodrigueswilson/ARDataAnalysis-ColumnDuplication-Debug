#!/usr/bin/env python3
"""
Verify the improved transparent table structure with reordered columns, clearer names,
% Excluded column, and logic explanation table
"""

import sys
import os

# Add the project directory to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def show_improved_table_structure():
    """Show what the improved transparent table structure looks like"""
    
    print("=== Improved Transparent Data Cleaning Table ===")
    
    try:
        # Get database connection
        db = get_db_connection()
        collection = db['media_records']
        
        print("\n📊 Table 1: Complete Filtering Breakdown (Improved Structure)")
        print("=" * 140)
        
        # New headers with improved names and order
        headers = ['File Type', 'Total Files Before Cleaning', 'Outliers', 'Non-School Days', 'Non-School Days and Outliers', 'Total Excluded', 'School Days', '% Excluded', '% Kept']
        print(f"{'|'.join(f'{h:>15}' for h in headers)}")
        print("-" * 140)
        
        # Calculate data for each file type
        file_types = ['JPG', 'MP3']
        table_data = []
        
        for file_type in file_types:
            # Total files
            total_files = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type
            })
            
            # Outliers (School Outliers): is_collection_day: TRUE AND Outlier_Status: TRUE
            outliers = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": True,
                "Outlier_Status": True
            })
            
            # Non-School Days (Non-School Normal): is_collection_day: FALSE AND Outlier_Status: FALSE
            non_school_days = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": False,
                "Outlier_Status": False
            })
            
            # Non-School Days and Outliers: is_collection_day: FALSE AND Outlier_Status: TRUE
            non_school_outliers = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": False,
                "Outlier_Status": True
            })
            
            # Total excluded (all categories except School Days)
            total_excluded = outliers + non_school_days + non_school_outliers
            
            # School Days (Final Dataset): is_collection_day: TRUE AND Outlier_Status: FALSE
            school_days = collection.count_documents({
                "School_Year": {"$ne": "N/A"},
                "file_type": file_type,
                "is_collection_day": True,
                "Outlier_Status": False
            })
            
            # Percentages
            exclusion_pct = (total_excluded / total_files * 100) if total_files > 0 else 0
            retention_pct = (school_days / total_files * 100) if total_files > 0 else 0
            
            # Store data
            row_data = [file_type, total_files, outliers, non_school_days, non_school_outliers, total_excluded, school_days, f"{exclusion_pct:.1f}%", f"{retention_pct:.1f}%"]
            table_data.append(row_data)
            
            # Print row
            print(f"{'|'.join(f'{str(val):>15}' for val in row_data)}")
        
        # Calculate totals
        totals = ['TOTAL']
        for i in range(1, 7):  # Skip file type column and percentage columns
            total_val = sum(row[i] for row in table_data)
            totals.append(total_val)
        
        # Calculate total percentages
        total_exclusion = (totals[5] / totals[1] * 100) if totals[1] > 0 else 0  # total_excluded / total_files
        total_retention = (totals[6] / totals[1] * 100) if totals[1] > 0 else 0  # school_days / total_files
        totals.extend([f"{total_exclusion:.1f}%", f"{total_retention:.1f}%"])
        
        print("-" * 140)
        print(f"{'|'.join(f'{str(val):>15}' for val in totals)}")
        
        print("\n📋 Logic Explanation: Category Definitions")
        print("=" * 80)
        print(f"{'Category':<30} | {'is_collection_day':<15} | {'Outlier_Status':<13} | {'Count':<8}")
        print("-" * 80)
        
        logic_data = [
            ('Outliers', 'TRUE', 'TRUE', totals[2]),
            ('Non-School Days', 'FALSE', 'FALSE', totals[3]),
            ('Non-School Days and Outliers', 'FALSE', 'TRUE', totals[4]),
            ('School Days (Final Dataset)', 'TRUE', 'FALSE', totals[6])
        ]
        
        for category, collection_day, outlier_status, count in logic_data:
            print(f"{category:<30} | {collection_day:<15} | {outlier_status:<13} | {count:<8}")
        
        print("\n🎯 Key Improvements:")
        print("✅ Columns reordered to group excluded categories together")
        print("✅ Column names clarified for better readability")
        print("✅ Added '% Excluded' column for transparency")
        print("✅ Added logic explanation table showing boolean conditions")
        print("✅ Final dataset (School Days) clearly separated from excluded categories")
        
        print("\n📈 Summary:")
        print(f"• Total files before cleaning: {totals[1]:,}")
        print(f"• Total excluded files: {totals[5]:,} ({total_exclusion:.1f}%)")
        print(f"  - Outliers from school days: {totals[2]:,}")
        print(f"  - Regular files from non-school days: {totals[3]:,}")
        print(f"  - Outliers from non-school days: {totals[4]:,}")
        print(f"• Final research dataset: {totals[6]:,} ({total_retention:.1f}%)")
        
        print("\n✅ Verification:")
        calculated_total = totals[2] + totals[3] + totals[4] + totals[6]
        print(f"Sum of all categories: {calculated_total:,}")
        print(f"Original total: {totals[1]:,}")
        print(f"Match: {'✅ Perfect' if calculated_total == totals[1] else '❌ Error'}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    show_improved_table_structure()
