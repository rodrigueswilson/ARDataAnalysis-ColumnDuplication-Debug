#!/usr/bin/env python3
"""
Investigate Directory File Types

This script investigates what file types actually exist in the data directories
and compares them with what's in the database. This will help identify if there
are unexpected file types causing data discrepancies.
"""

import sys
import os
from pathlib import Path
from collections import Counter, defaultdict
from pymongo import MongoClient

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def scan_directory_file_types(base_path):
    """
    Scan directories for all file types present.
    
    Args:
        base_path (str): Base path to scan
        
    Returns:
        dict: File type statistics
    """
    print(f"Scanning directory: {base_path}")
    
    if not os.path.exists(base_path):
        print(f"❌ Directory not found: {base_path}")
        return {}
    
    file_types = Counter()
    file_details = defaultdict(list)
    total_files = 0
    
    # Scan all subdirectories
    for root, dirs, files in os.walk(base_path):
        for file in files:
            total_files += 1
            file_path = os.path.join(root, file)
            
            # Get file extension
            _, ext = os.path.splitext(file)
            ext = ext.upper().lstrip('.')  # Remove dot and convert to uppercase
            
            if not ext:
                ext = 'NO_EXTENSION'
            
            file_types[ext] += 1
            file_details[ext].append(file_path)
    
    print(f"Total files found: {total_files}")
    print(f"File types found: {len(file_types)}")
    
    return {
        'total_files': total_files,
        'file_types': dict(file_types),
        'file_details': dict(file_details)
    }

def investigate_database_file_types():
    """
    Investigate what file types are in the database.
    """
    print("\n" + "="*60)
    print("DATABASE FILE TYPE ANALYSIS")
    print("="*60)
    
    db = get_db_connection()
    collection = db.media_records
    
    print(f"Connected to database: {db.name}")
    print(f"Collection: {collection.name}")
    
    # Get all file types in database
    pipeline = [
        {"$group": {"_id": "$file_type", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}}
    ]
    
    db_file_types = list(collection.aggregate(pipeline))
    
    print("File types in database:")
    total_db_files = 0
    for item in db_file_types:
        file_type = item['_id'] if item['_id'] else 'NULL'
        count = item['count']
        total_db_files += count
        print(f"  {file_type}: {count}")
    
    print(f"Total files in database: {total_db_files}")
    
    return db_file_types, total_db_files

def investigate_specific_directories():
    """
    Investigate the specific data directories mentioned in the project.
    """
    print("\n" + "="*60)
    print("DIRECTORY FILE TYPE ANALYSIS")
    print("="*60)
    
    # Common directory patterns based on project structure
    possible_directories = [
        "D:/ARDataAnalysis",
        "D:/ARDataAnalysis/21_22 Photos",
        "D:/ARDataAnalysis/22_23 Photos", 
        "D:/ARDataAnalysis/21_22 Audio",
        "D:/ARDataAnalysis/22_23 Audio",
        "D:/ARDataAnalysis/TestData",
        # Alternative paths
        "../ARDataAnalysis",
        "../../ARDataAnalysis",
        "./data",
        "../data"
    ]
    
    directory_results = {}
    
    for dir_path in possible_directories:
        if os.path.exists(dir_path):
            print(f"\n📁 Found directory: {dir_path}")
            results = scan_directory_file_types(dir_path)
            directory_results[dir_path] = results
            
            # Show file type breakdown
            if results['file_types']:
                print("File type breakdown:")
                for file_type, count in sorted(results['file_types'].items(), key=lambda x: x[1], reverse=True):
                    print(f"  {file_type}: {count}")
                    
                    # Show unexpected file types
                    if file_type not in ['JPG', 'MP3']:
                        print(f"    ⚠️  UNEXPECTED FILE TYPE: {file_type}")
                        # Show first few examples
                        examples = results['file_details'][file_type][:3]
                        for example in examples:
                            print(f"      Example: {example}")
        else:
            print(f"❌ Directory not found: {dir_path}")
    
    return directory_results

def compare_directory_vs_database(directory_results, db_file_types):
    """
    Compare directory scan results with database contents.
    """
    print("\n" + "="*60)
    print("DIRECTORY vs DATABASE COMPARISON")
    print("="*60)
    
    # Aggregate all directory file types
    total_dir_files = 0
    dir_file_types = Counter()
    
    for dir_path, results in directory_results.items():
        total_dir_files += results['total_files']
        for file_type, count in results['file_types'].items():
            dir_file_types[file_type] += count
    
    # Convert database results to Counter
    db_file_types_counter = Counter()
    total_db_files = 0
    for item in db_file_types:
        file_type = item['_id'] if item['_id'] else 'NULL'
        count = item['count']
        db_file_types_counter[file_type] = count
        total_db_files += count
    
    print(f"Total files in directories: {total_dir_files}")
    print(f"Total files in database: {total_db_files}")
    print(f"Difference: {total_dir_files - total_db_files}")
    print()
    
    # Compare file types
    all_file_types = set(dir_file_types.keys()) | set(db_file_types_counter.keys())
    
    print("File type comparison:")
    print(f"{'File Type':<15} {'Directory':<10} {'Database':<10} {'Difference':<10} {'Status'}")
    print("-" * 60)
    
    for file_type in sorted(all_file_types):
        dir_count = dir_file_types.get(file_type, 0)
        db_count = db_file_types_counter.get(file_type, 0)
        difference = dir_count - db_count
        
        if difference == 0:
            status = "✅ Match"
        elif difference > 0:
            status = "⚠️  More in dir"
        else:
            status = "⚠️  More in DB"
        
        print(f"{file_type:<15} {dir_count:<10} {db_count:<10} {difference:<10} {status}")
    
    # Identify unexpected file types
    unexpected_types = [ft for ft in all_file_types if ft not in ['JPG', 'MP3']]
    
    if unexpected_types:
        print(f"\n⚠️  UNEXPECTED FILE TYPES FOUND: {unexpected_types}")
        print("These file types should not be present if directories contain only JPG and MP3 files.")
    else:
        print(f"\n✅ Only expected file types found: JPG and MP3")

def main():
    print("DIRECTORY FILE TYPE INVESTIGATION")
    print("="*80)
    print("This script investigates what file types exist in directories vs database")
    print("to identify potential data inconsistencies.")
    print()
    
    try:
        # 1. Investigate database file types
        db_file_types, total_db_files = investigate_database_file_types()
        
        # 2. Investigate directory file types
        directory_results = investigate_specific_directories()
        
        # 3. Compare results
        if directory_results:
            compare_directory_vs_database(directory_results, db_file_types)
        else:
            print("\n❌ No data directories found for comparison")
            print("Please verify the correct path to your data directories.")
        
        print("\n" + "="*60)
        print("INVESTIGATION COMPLETE")
        print("="*60)
        
    except Exception as e:
        print(f"Error during investigation: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
