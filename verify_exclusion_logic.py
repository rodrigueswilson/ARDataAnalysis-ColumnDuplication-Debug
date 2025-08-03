#!/usr/bin/env python3
"""
Verify Data Cleaning Exclusion Logic
====================================

This script verifies whether files are being excluded multiple times or if the 
logic properly creates mutually exclusive categories.
"""

from db_utils import get_db_connection

def analyze_exclusion_logic():
    """
    Analyze the exclusion logic to verify files are not double-counted.
    """
    print("=" * 80)
    print("DATA CLEANING EXCLUSION LOGIC ANALYSIS")
    print("=" * 80)
    
    # Get database connection
    db = get_db_connection()
    collection = db['media_records']
    
    # Base filter (same as Data Cleaning sheet)
    base_filter = {
        "School_Year": {"$ne": "N/A"},
        "file_type": {"$in": ["JPG", "MP3"]}
    }
    
    print("\n1. TOTAL FILES BY CATEGORY (Raw Counts)")
    print("-" * 50)
    
    # Get raw counts for each combination
    total_files = collection.count_documents(base_filter)
    print(f"Total Files (base filter): {total_files}")
    
    # Collection day files
    collection_day_files = collection.count_documents({
        **base_filter,
        "is_collection_day": True
    })
    print(f"Collection Day Files: {collection_day_files}")
    
    # Non-collection day files  
    non_collection_day_files = collection.count_documents({
        **base_filter,
        "is_collection_day": False
    })
    print(f"Non-Collection Day Files: {non_collection_day_files}")
    
    # Outlier files
    outlier_files = collection.count_documents({
        **base_filter,
        "Outlier_Status": True
    })
    print(f"Outlier Files: {outlier_files}")
    
    # Non-outlier files
    non_outlier_files = collection.count_documents({
        **base_filter,
        "Outlier_Status": False
    })
    print(f"Non-Outlier Files: {non_outlier_files}")
    
    print(f"\nVerification: Collection + Non-Collection = {collection_day_files + non_collection_day_files} (should equal {total_files})")
    print(f"Verification: Outlier + Non-Outlier = {outlier_files + non_outlier_files} (should equal {total_files})")
    
    print("\n2. FOUR MUTUALLY EXCLUSIVE CATEGORIES")
    print("-" * 50)
    
    # Category 1: School Normal (Collection Day + Non-Outlier)
    school_normal = collection.count_documents({
        **base_filter,
        "is_collection_day": True,
        "Outlier_Status": False
    })
    print(f"School Normal (Collection Day + Non-Outlier): {school_normal}")
    
    # Category 2: School Outliers (Collection Day + Outlier)
    school_outliers = collection.count_documents({
        **base_filter,
        "is_collection_day": True,
        "Outlier_Status": True
    })
    print(f"School Outliers (Collection Day + Outlier): {school_outliers}")
    
    # Category 3: Non-School Normal (Non-Collection Day + Non-Outlier)
    non_school_normal = collection.count_documents({
        **base_filter,
        "is_collection_day": False,
        "Outlier_Status": False
    })
    print(f"Non-School Normal (Non-Collection Day + Non-Outlier): {non_school_normal}")
    
    # Category 4: Non-School Outliers (Non-Collection Day + Outlier)
    non_school_outliers = collection.count_documents({
        **base_filter,
        "is_collection_day": False,
        "Outlier_Status": True
    })
    print(f"Non-School Outliers (Non-Collection Day + Outlier): {non_school_outliers}")
    
    # Verify mutually exclusive
    category_total = school_normal + school_outliers + non_school_normal + non_school_outliers
    print(f"\nSum of all categories: {category_total}")
    print(f"Total files: {total_files}")
    print(f"Categories are mutually exclusive: {category_total == total_files}")
    
    print("\n3. EXCLUSION ANALYSIS")
    print("-" * 50)
    
    total_excluded = school_outliers + non_school_normal + non_school_outliers
    retention_rate = (school_normal / total_files * 100) if total_files > 0 else 0
    exclusion_rate = (total_excluded / total_files * 100) if total_files > 0 else 0
    
    print(f"Files Kept (School Normal): {school_normal} ({retention_rate:.1f}%)")
    print(f"Files Excluded: {total_excluded} ({exclusion_rate:.1f}%)")
    print(f"  - School Outliers: {school_outliers}")
    print(f"  - Non-School Normal: {non_school_normal}")
    print(f"  - Non-School Outliers: {non_school_outliers}")
    
    print("\n4. FILE TYPE BREAKDOWN")
    print("-" * 50)
    
    for file_type in ["JPG", "MP3"]:
        print(f"\n{file_type} Files:")
        
        type_filter = {**base_filter, "file_type": file_type}
        
        type_total = collection.count_documents(type_filter)
        type_school_normal = collection.count_documents({**type_filter, "is_collection_day": True, "Outlier_Status": False})
        type_school_outliers = collection.count_documents({**type_filter, "is_collection_day": True, "Outlier_Status": True})
        type_non_school_normal = collection.count_documents({**type_filter, "is_collection_day": False, "Outlier_Status": False})
        type_non_school_outliers = collection.count_documents({**type_filter, "is_collection_day": False, "Outlier_Status": True})
        
        type_category_total = type_school_normal + type_school_outliers + type_non_school_normal + type_non_school_outliers
        
        print(f"  Total: {type_total}")
        print(f"  School Normal: {type_school_normal}")
        print(f"  School Outliers: {type_school_outliers}")
        print(f"  Non-School Normal: {type_non_school_normal}")
        print(f"  Non-School Outliers: {type_non_school_outliers}")
        print(f"  Category Sum: {type_category_total} (matches total: {type_category_total == type_total})")
    
    print("\n5. CONCLUSION")
    print("-" * 50)
    
    if category_total == total_files:
        print("✅ FILES ARE NOT EXCLUDED MULTIPLE TIMES")
        print("✅ Each file appears in exactly ONE category")
        print("✅ Categories are mutually exclusive")
        print("✅ No double-counting occurs")
    else:
        print("❌ LOGIC ERROR DETECTED")
        print("❌ Files may be double-counted or missing")
        print(f"❌ Discrepancy: {abs(category_total - total_files)} files")
    
    print("\n" + "=" * 80)

if __name__ == "__main__":
    analyze_exclusion_logic()
