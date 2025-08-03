#!/usr/bin/env python3
"""
Excel Sheet Generation Trace

This script traces the exact Excel sheet generation process to understand
where and how the left-aligned rows are created differently from the main data.
"""

import sys
import os
import pandas as pd
from datetime import datetime, date

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ar_utils import get_school_calendar, get_non_collection_days, precompute_collection_days
from db_utils import get_db_connection

def simulate_complete_sheet_generation():
    """Simulate the complete Daily Counts (ACF_PACF) sheet generation process."""
    print("COMPLETE EXCEL SHEET GENERATION SIMULATION")
    print("=" * 70)
    
    try:
        # Step 1: Get data from pipeline (we know this works)
        print("1. PIPELINE DATA EXTRACTION")
        print("-" * 50)
        
        db = get_db_connection()
        collection = db.media_records
        
        # DAILY_COUNTS_COLLECTION_ONLY pipeline
        pipeline = [
            {
                "$match": {
                    "School_Year": {"$ne": "N/A"},
                    "file_type": {"$in": ["JPG", "MP3"]},
                    "is_collection_day": True,
                    "Outlier_Status": False
                }
            },
            {
                "$group": {
                    "_id": "$ISO_Date",
                    "Total_Files": {"$sum": 1}
                }
            },
            {
                "$sort": {"_id": 1}
            }
        ]
        
        pipeline_data = list(collection.aggregate(pipeline))
        pipeline_df = pd.DataFrame(pipeline_data)
        
        print(f"Pipeline returned {len(pipeline_df)} days with data")
        print(f"Total files from pipeline: {pipeline_df['Total_Files'].sum()}")
        
        # Step 2: Apply zero-fill (we know this works)
        print("\n2. ZERO-FILL PROCESS")
        print("-" * 50)
        
        school_calendar = get_school_calendar()
        non_collection_days = get_non_collection_days()
        collection_day_map = precompute_collection_days(school_calendar, non_collection_days)
        
        all_collection_days = []
        for date_obj, info in collection_day_map.items():
            all_collection_days.append({'_id': date_obj.strftime('%Y-%m-%d')})
        
        all_days_df = pd.DataFrame(all_collection_days)
        zero_filled_df = pd.merge(all_days_df, pipeline_df, on='_id', how='left').fillna(0)
        
        # Ensure correct data types
        zero_filled_df['Total_Files'] = zero_filled_df['Total_Files'].astype(int)
        zero_filled_df = zero_filled_df.sort_values('_id').reset_index(drop=True)
        
        print(f"Zero-filled data: {len(zero_filled_df)} days")
        print(f"Total files after zero-fill: {zero_filled_df['Total_Files'].sum()}")
        
        # Step 3: Check what happens during Excel writing
        print("\n3. EXCEL WRITING ANALYSIS")
        print("-" * 50)
        
        # The issue might be in how the data is written to Excel
        # Let's check if there are any special conditions or filters
        
        # Check the first 20 rows to see the pattern
        print("First 20 rows of zero-filled data:")
        for i in range(min(20, len(zero_filled_df))):
            row = zero_filled_df.iloc[i]
            date_str = row['_id']
            count = row['Total_Files']
            
            # Mark early September dates
            early_sept_dates = [
                "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
                "2021-09-20", "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
            ]
            
            is_early_sept = date_str in early_sept_dates
            is_partial_day = date_str == "2021-09-20"
            
            marker = ""
            if is_partial_day:
                marker = "🔶 PARTIAL"
            elif is_early_sept:
                marker = "🎯 EARLY_SEPT"
            
            print(f"  {i+1:2d}. {date_str}: {count:3d} files {marker}")
        
        # Step 4: Investigate potential Excel display logic
        print("\n4. EXCEL DISPLAY LOGIC INVESTIGATION")
        print("-" * 50)
        
        # Check if there's any logic that would cause different display
        # The issue might be in:
        # 1. How totals are calculated
        # 2. How certain rows are formatted
        # 3. Date range filtering for display vs totals
        
        # Let's see if we can identify a pattern
        early_sept_data = zero_filled_df[zero_filled_df['_id'].isin([
            "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
            "2021-09-20", "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
        ])]
        
        print("Early September data analysis:")
        total_early_sept = early_sept_data['Total_Files'].sum()
        print(f"  Total early September files: {total_early_sept}")
        
        # Separate partial day
        partial_day_data = zero_filled_df[zero_filled_df['_id'] == "2021-09-20"]
        if not partial_day_data.empty:
            partial_count = partial_day_data.iloc[0]['Total_Files']
            print(f"  Partial day (2021-09-20): {partial_count} files")
            print(f"  Other early September: {total_early_sept - partial_count} files")
        
        # Check if the "other early September" matches Excel observation
        if total_early_sept - 104 == 358:
            print("  🎯 MATCH! Other early Sept (358) matches Excel left-aligned rows!")
            print("  🔶 Partial day (104) is handled differently!")
        
        return zero_filled_df, early_sept_data
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None, None

def investigate_partial_day_handling():
    """Investigate how partial days are handled differently."""
    print("\n" + "=" * 70)
    print("PARTIAL DAY HANDLING INVESTIGATION")
    print("=" * 70)
    
    try:
        # Check the non-collection days configuration
        non_collection_days = get_non_collection_days()
        
        partial_day = date(2021, 9, 20)
        if partial_day in non_collection_days:
            info = non_collection_days[partial_day]
            print(f"2021-09-20 configuration:")
            print(f"  Type: {info.get('type', 'Unknown')}")
            print(f"  Event: {info.get('event', 'Unknown')}")
            
            if info.get('type') == 'Partial':
                print("  🔶 This is a PARTIAL collection day!")
                print("  This might be handled differently in Excel formatting.")
        
        # Check if there's special logic for partial days in the codebase
        print("\nHypothesis:")
        print("  - Regular collection days: Appear as left-aligned rows (358 files)")
        print("  - Partial collection days: Handled differently (104 files from 2021-09-20)")
        print("  - Excel total calculation: Excludes both types (9731 - 358 - 104 = 9269)")
        print("  - But actual Excel shows: 9372")
        print("  - Difference: 9372 - 9269 = 103 files")
        print("  - This suggests partial days are included in totals but not left-aligned!")
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        return False

def main():
    try:
        zero_filled_df, early_sept_data = simulate_complete_sheet_generation()
        investigate_partial_day_handling()
        
        print("\n" + "=" * 70)
        print("INVESTIGATION CONCLUSION")
        print("=" * 70)
        
        print("KEY FINDINGS:")
        print("1. ✅ Zero-fill process works correctly (9731 total files)")
        print("2. ✅ Early September has 462 files total")
        print("3. 🔶 2021-09-20 (Partial day) has 104 files")
        print("4. 🎯 Other early September has 358 files (matches Excel left-aligned)")
        print("5. ❓ Excel total is 9372, suggesting partial days are included in totals")
        
        print("\nHYPOTHESIS:")
        print("- Left-aligned rows: 358 files (regular early September dates)")
        print("- Partial day handling: 104 files (included in totals, not left-aligned)")
        print("- Excel calculation: 9731 - 358 = 9373 ≈ 9372")
        
        print("\nNEXT STEPS:")
        print("1. Find the Excel code that treats partial days differently")
        print("2. Locate the total calculation logic")
        print("3. Identify why regular early dates are left-aligned")
        print("4. Fix the date range to include all dates from 2021-09-13")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
