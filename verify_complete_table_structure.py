#!/usr/bin/env python3
"""
Verify the complete Data Cleaning sheet structure with all three tables:
1. Table 1: Complete Filtering Breakdown (by file type)
2. Table 2: Year-by-Year Breakdown
3. Logic Explanation Table
"""

import sys
import os

# Add the project directory to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def show_complete_table_structure():
    """Show the complete Data Cleaning sheet structure with all tables"""
    
    print("=== Complete Data Cleaning Sheet Structure ===")
    
    try:
        # Get database connection
        db = get_db_connection()
        collection = db['media_records']
        
        # Table 1: Complete Filtering Breakdown (by file type)
        print("\n📊 Table 1: Complete Filtering Breakdown")
        print("=" * 140)
        
        headers = ['File Type', 'Total Files Before Cleaning', 'Outliers', 'Non-School Days', 'Non-School Days and Outliers', 'Total Excluded', 'School Days', '% Excluded', '% Kept']
        print(f"{'|'.join(f'{h:>15}' for h in headers)}")
        print("-" * 140)
        
        # Calculate data for each file type
        file_types = ['JPG', 'MP3']
        table1_data = []
        
        for file_type in file_types:
            # Calculate all categories for this file type
            total_files = collection.count_documents({"School_Year": {"$ne": "N/A"}, "file_type": file_type})
            outliers = collection.count_documents({"School_Year": {"$ne": "N/A"}, "file_type": file_type, "is_collection_day": True, "Outlier_Status": True})
            non_school_days = collection.count_documents({"School_Year": {"$ne": "N/A"}, "file_type": file_type, "is_collection_day": False, "Outlier_Status": False})
            non_school_outliers = collection.count_documents({"School_Year": {"$ne": "N/A"}, "file_type": file_type, "is_collection_day": False, "Outlier_Status": True})
            school_days = collection.count_documents({"School_Year": {"$ne": "N/A"}, "file_type": file_type, "is_collection_day": True, "Outlier_Status": False})
            
            total_excluded = outliers + non_school_days + non_school_outliers
            exclusion_pct = (total_excluded / total_files * 100) if total_files > 0 else 0
            retention_pct = (school_days / total_files * 100) if total_files > 0 else 0
            
            row_data = [file_type, total_files, outliers, non_school_days, non_school_outliers, total_excluded, school_days, f"{exclusion_pct:.1f}%", f"{retention_pct:.1f}%"]
            table1_data.append(row_data)
            print(f"{'|'.join(f'{str(val):>15}' for val in row_data)}")
        
        # Table 1 totals
        totals = ['TOTAL']
        for i in range(1, 7):
            total_val = sum(row[i] for row in table1_data)
            totals.append(total_val)
        total_exclusion = (totals[5] / totals[1] * 100) if totals[1] > 0 else 0
        total_retention = (totals[6] / totals[1] * 100) if totals[1] > 0 else 0
        totals.extend([f"{total_exclusion:.1f}%", f"{total_retention:.1f}%"])
        
        print("-" * 140)
        print(f"{'|'.join(f'{str(val):>15}' for val in totals)}")
        
        # Table 2: Year-by-Year Breakdown
        print("\n\n📊 Table 2: Year-by-Year Breakdown")
        print("=" * 140)
        
        headers2 = ['Category', 'Total Files Before Cleaning', 'Outliers', 'Non-School Days', 'Non-School Days and Outliers', 'Total Excluded', 'School Days', '% Excluded', '% Kept']
        print(f"{'|'.join(f'{h:>15}' for h in headers2)}")
        print("-" * 140)
        
        # Calculate year-by-year data
        years = ["2021-2022", "2022-2023"]
        table2_data = []
        
        for year in years:
            for file_type in ["JPG", "MP3"]:
                category = f"{year} {file_type} Files"
                
                # Calculate all categories for this year/type
                total_files = collection.count_documents({"School_Year": year, "file_type": file_type})
                outliers = collection.count_documents({"School_Year": year, "file_type": file_type, "is_collection_day": True, "Outlier_Status": True})
                non_school_days = collection.count_documents({"School_Year": year, "file_type": file_type, "is_collection_day": False, "Outlier_Status": False})
                non_school_outliers = collection.count_documents({"School_Year": year, "file_type": file_type, "is_collection_day": False, "Outlier_Status": True})
                school_days = collection.count_documents({"School_Year": year, "file_type": file_type, "is_collection_day": True, "Outlier_Status": False})
                
                total_excluded = outliers + non_school_days + non_school_outliers
                exclusion_pct = (total_excluded / total_files * 100) if total_files > 0 else 0
                retention_pct = (school_days / total_files * 100) if total_files > 0 else 0
                
                row_data = [category, total_files, outliers, non_school_days, non_school_outliers, total_excluded, school_days, f"{exclusion_pct:.1f}%", f"{retention_pct:.1f}%"]
                table2_data.append(row_data)
                print(f"{'|'.join(f'{str(val):>15}' for val in row_data)}")
        
        # Table 2 totals
        year_totals = ['TOTAL']
        for i in range(1, 7):
            total_val = sum(row[i] for row in table2_data)
            year_totals.append(total_val)
        year_total_exclusion = (year_totals[5] / year_totals[1] * 100) if year_totals[1] > 0 else 0
        year_total_retention = (year_totals[6] / year_totals[1] * 100) if year_totals[1] > 0 else 0
        year_totals.extend([f"{year_total_exclusion:.1f}%", f"{year_total_retention:.1f}%"])
        
        print("-" * 140)
        print(f"{'|'.join(f'{str(val):>15}' for val in year_totals)}")
        
        # Logic Explanation Table
        print("\n\n📋 Logic Explanation: Category Definitions")
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
        
        print("\n🎯 Key Features of the Complete Structure:")
        print("✅ Table 1: Overall breakdown by file type (JPG/MP3)")
        print("✅ Table 2: Detailed breakdown by school year and file type")
        print("✅ Logic Explanation: Boolean conditions for each category")
        print("✅ Consistent column structure across all tables")
        print("✅ Both exclusion and retention percentages shown")
        print("✅ Mathematical transparency and auditability")
        
        print("\n📈 Cross-Table Verification:")
        print(f"• Table 1 total files: {totals[1]:,}")
        print(f"• Table 2 total files: {year_totals[1]:,}")
        print(f"• Match: {'✅ Perfect' if totals[1] == year_totals[1] else '❌ Mismatch'}")
        
        print(f"• Table 1 final dataset: {totals[6]:,}")
        print(f"• Table 2 final dataset: {year_totals[6]:,}")
        print(f"• Match: {'✅ Perfect' if totals[6] == year_totals[6] else '❌ Mismatch'}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    show_complete_table_structure()
