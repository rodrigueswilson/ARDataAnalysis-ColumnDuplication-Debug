#!/usr/bin/env python3
"""
Test to verify that the three excluded categories are mutually exclusive
and that no file appears in multiple lists.
"""

import sys
import os

# Add the project directory to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def test_mutual_exclusivity():
    """Test that excluded categories are mutually exclusive - no file appears in multiple lists"""
    
    print("=== Mutual Exclusivity Test ===")
    print("Testing if any file appears in multiple excluded categories")
    
    try:
        # Get database connection
        db = get_db_connection()
        collection = db['media_records']
        
        print("\n🔍 Step 1: Get all files in each excluded category")
        
        # Get actual file IDs for each category
        print("\n📋 School Outliers (is_collection_day: TRUE AND Outlier_Status: TRUE):")
        school_outliers_cursor = collection.find({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": True,
            "Outlier_Status": True
        }, {"_id": 1, "file_type": 1, "filename": 1})
        
        school_outliers_files = list(school_outliers_cursor)
        school_outliers_ids = set(str(doc["_id"]) for doc in school_outliers_files)
        print(f"   Count: {len(school_outliers_files)}")
        
        if school_outliers_files:
            print("   Sample files:")
            for i, doc in enumerate(school_outliers_files[:3]):
                print(f"     {i+1}. {doc.get('filename', 'N/A')} ({doc.get('file_type', 'N/A')})")
        
        print("\n📋 Non-School Normal (is_collection_day: FALSE AND Outlier_Status: FALSE):")
        non_school_normal_cursor = collection.find({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": False,
            "Outlier_Status": False
        }, {"_id": 1, "file_type": 1, "filename": 1})
        
        non_school_normal_files = list(non_school_normal_cursor)
        non_school_normal_ids = set(str(doc["_id"]) for doc in non_school_normal_files)
        print(f"   Count: {len(non_school_normal_files)}")
        
        if non_school_normal_files:
            print("   Sample files:")
            for i, doc in enumerate(non_school_normal_files[:3]):
                print(f"     {i+1}. {doc.get('filename', 'N/A')} ({doc.get('file_type', 'N/A')})")
        
        print("\n📋 Non-School Outliers (is_collection_day: FALSE AND Outlier_Status: TRUE):")
        non_school_outliers_cursor = collection.find({
            "School_Year": {"$ne": "N/A"},
            "is_collection_day": False,
            "Outlier_Status": True
        }, {"_id": 1, "file_type": 1, "filename": 1})
        
        non_school_outliers_files = list(non_school_outliers_cursor)
        non_school_outliers_ids = set(str(doc["_id"]) for doc in non_school_outliers_files)
        print(f"   Count: {len(non_school_outliers_files)}")
        
        if non_school_outliers_files:
            print("   Sample files:")
            for i, doc in enumerate(non_school_outliers_files[:3]):
                print(f"     {i+1}. {doc.get('filename', 'N/A')} ({doc.get('file_type', 'N/A')})")
        
        print(f"\n🔍 Step 2: Test for overlaps between categories")
        
        # Test overlap between School Outliers and Non-School Normal
        overlap_1_2 = school_outliers_ids.intersection(non_school_normal_ids)
        print(f"\n❓ School Outliers ∩ Non-School Normal: {len(overlap_1_2)} files")
        if overlap_1_2:
            print(f"   ❌ OVERLAP FOUND! Files: {list(overlap_1_2)[:5]}")
        else:
            print(f"   ✅ No overlap - mutually exclusive")
        
        # Test overlap between School Outliers and Non-School Outliers
        overlap_1_3 = school_outliers_ids.intersection(non_school_outliers_ids)
        print(f"\n❓ School Outliers ∩ Non-School Outliers: {len(overlap_1_3)} files")
        if overlap_1_3:
            print(f"   ❌ OVERLAP FOUND! Files: {list(overlap_1_3)[:5]}")
        else:
            print(f"   ✅ No overlap - mutually exclusive")
        
        # Test overlap between Non-School Normal and Non-School Outliers
        overlap_2_3 = non_school_normal_ids.intersection(non_school_outliers_ids)
        print(f"\n❓ Non-School Normal ∩ Non-School Outliers: {len(overlap_2_3)} files")
        if overlap_2_3:
            print(f"   ❌ OVERLAP FOUND! Files: {list(overlap_2_3)[:5]}")
        else:
            print(f"   ✅ No overlap - mutually exclusive")
        
        # Test three-way overlap
        overlap_all = school_outliers_ids.intersection(non_school_normal_ids).intersection(non_school_outliers_ids)
        print(f"\n❓ Three-way overlap: {len(overlap_all)} files")
        if overlap_all:
            print(f"   ❌ THREE-WAY OVERLAP FOUND! Files: {list(overlap_all)}")
        else:
            print(f"   ✅ No three-way overlap")
        
        print(f"\n🎯 Step 3: Verify logical impossibility")
        print("Checking if the conditions make overlaps logically impossible:")
        
        print("\n📊 Condition Analysis:")
        print("• School Outliers:     is_collection_day=TRUE  AND Outlier_Status=TRUE")
        print("• Non-School Normal:   is_collection_day=FALSE AND Outlier_Status=FALSE")
        print("• Non-School Outliers: is_collection_day=FALSE AND Outlier_Status=TRUE")
        
        print("\n🧮 Logical Analysis:")
        print("• School Outliers vs Non-School Normal:")
        print("  - Different is_collection_day values (TRUE vs FALSE) → Impossible overlap ✅")
        print("• School Outliers vs Non-School Outliers:")
        print("  - Different is_collection_day values (TRUE vs FALSE) → Impossible overlap ✅")
        print("• Non-School Normal vs Non-School Outliers:")
        print("  - Different Outlier_Status values (FALSE vs TRUE) → Impossible overlap ✅")
        
        print(f"\n📈 Step 4: Summary Statistics")
        total_excluded = len(school_outliers_ids) + len(non_school_normal_ids) + len(non_school_outliers_ids)
        all_excluded_ids = school_outliers_ids.union(non_school_normal_ids).union(non_school_outliers_ids)
        
        print(f"• School Outliers: {len(school_outliers_ids):,} files")
        print(f"• Non-School Normal: {len(non_school_normal_ids):,} files")
        print(f"• Non-School Outliers: {len(non_school_outliers_ids):,} files")
        print(f"• Sum of categories: {total_excluded:,} files")
        print(f"• Union of all sets: {len(all_excluded_ids):,} files")
        print(f"• Match: {'✅ Perfect' if total_excluded == len(all_excluded_ids) else '❌ Overlap detected'}")
        
        print(f"\n🏆 FINAL RESULT:")
        total_overlaps = len(overlap_1_2) + len(overlap_1_3) + len(overlap_2_3) + len(overlap_all)
        if total_overlaps == 0:
            print("✅ SUCCESS: All three excluded categories are perfectly mutually exclusive!")
            print("✅ No file appears in multiple categories!")
            print("✅ The intersection logic is mathematically sound!")
        else:
            print("❌ FAILURE: Overlaps detected between categories!")
            print(f"❌ Total overlapping files: {total_overlaps}")
        
    except Exception as e:
        print(f"❌ Error during test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_mutual_exclusivity()
