"""
Test script for DataCleaningUtils class.

This script validates that the DataCleaningUtils class correctly implements
the 2×2 matrix filtering logic and produces the expected results.
"""

import pandas as pd
from utils.data_cleaning import DataCleaningUtils
from db_utils import get_db_connection

def test_data_cleaning_utils():
    """Test the DataCleaningUtils class with real data."""
    print("=" * 80)
    print("Testing DataCleaningUtils with real data...")
    print("=" * 80)
    
    # Get database connection
    try:
        print("\nConnecting to database...")
        db = get_db_connection()
        print("✓ Database connection successful")
    except Exception as e:
        print(f"✗ Failed to connect to database: {e}")
        return
    
    # Initialize DataCleaningUtils
    try:
        print("\nInitializing DataCleaningUtils...")
        utils = DataCleaningUtils(db)
        print("✓ DataCleaningUtils initialized")
    except Exception as e:
        print(f"✗ Failed to initialize DataCleaningUtils: {e}")
        return
    
    # Test complete analysis
    try:
        print("\nRunning complete data cleaning analysis...")
        result = utils.get_complete_cleaning_data()
        print("✓ Complete analysis successful")
    except Exception as e:
        print(f"✗ Failed to run complete analysis: {e}")
        return
    
    # Verify results
    intersection_data = result['intersection_data']
    totals = result['totals']
    
    print("\n=== Complete Filtering Breakdown ===")
    for item in intersection_data[:-1]:  # Skip the last item which is totals
        print(f"{item['file_type']}: {item['total_files']} total, {item['school_normal']} in final dataset ({item['retention_pct']:.1f}% retention)")
    
    print("\n=== Totals ===")
    print(f"Total files: {totals['total_files']}")
    print(f"School normal (final dataset): {totals['school_normal']}")
    print(f"School outliers: {totals['school_outliers']}")
    print(f"Non-school normal: {totals['non_school_normal']}")
    print(f"Non-school outliers: {totals['non_school_outliers']}")
    print(f"Retention rate: {totals['retention_pct']:.1f}%")
    print(f"Exclusion rate: {totals['exclusion_pct']:.1f}%")
    
    # Validate math: retention + exclusion should equal 100%
    total_percentage = totals['retention_pct'] + totals['exclusion_pct']
    print(f"Total percentage: {total_percentage:.1f}% (should be 100.0%)")
    
    # Test year breakdown
    try:
        print("\nRunning year-by-year breakdown...")
        year_data = utils.get_year_breakdown_data()
        print("✓ Year breakdown successful")
    except Exception as e:
        print(f"✗ Failed to run year breakdown: {e}")
        return
    
    print("\n=== Year Breakdown ===")
    for item in year_data[:-1]:  # Skip the last item which is totals
        print(f"{item['category']}: {item['total_files']} total, {item['school_days']} in final dataset ({item['retention_pct']:.1f}% retention)")
        
    # Get the totals from year breakdown
    year_totals = year_data[-1]
    print("\n=== Year Totals ===")
    print(f"Total files: {year_totals['total_files']}")
    print(f"School days (final dataset): {year_totals['school_days']}")
    print(f"Retention rate: {year_totals['retention_pct']:.1f}%")
    
    # Compare the totals from both methods to ensure consistency
    print("\n=== Consistency Check ===")
    print(f"Complete analysis total files: {totals['total_files']}, Year breakdown total files: {year_totals['total_files']}")
    print(f"Complete analysis final dataset: {totals['school_normal']}, Year breakdown final dataset: {year_totals['school_days']}")
    print(f"Complete analysis retention rate: {totals['retention_pct']:.1f}%, Year breakdown retention rate: {year_totals['retention_pct']:.1f}%")
    
    is_consistent = (totals['total_files'] == year_totals['total_files'] and 
                    totals['school_normal'] == year_totals['school_days'] and 
                    abs(totals['retention_pct'] - year_totals['retention_pct']) < 0.1)
    
    print(f"Consistency check: {'✓ Passed' if is_consistent else '✗ Failed'}")
    
    print("\n=== Test Complete ===")
    return result


if __name__ == "__main__":
    test_data_cleaning_utils()
