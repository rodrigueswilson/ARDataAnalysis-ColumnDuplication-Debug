#!/usr/bin/env python3
"""
Zero-Fill Process Investigation

Since all early September dates ARE in the collection day map, the issue must be
elsewhere in the zero-fill process. Let's trace the complete _fill_missing_collection_days
process step by step.
"""

import sys
import os
import pandas as pd
from datetime import datetime, date

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ar_utils import get_school_calendar, get_non_collection_days, precompute_collection_days
from db_utils import get_db_connection

def simulate_zero_fill_process():
    """Simulate the exact zero-fill process used in _fill_missing_collection_days."""
    print("ZERO-FILL PROCESS SIMULATION")
    print("=" * 60)
    
    try:
        # Step 1: Get the collection day map (we know this works)
        print("1. GETTING COLLECTION DAY MAP")
        print("-" * 40)
        
        school_calendar = get_school_calendar()
        non_collection_days = get_non_collection_days()
        collection_day_map = precompute_collection_days(school_calendar, non_collection_days)
        
        print(f"Collection day map contains {len(collection_day_map)} days")
        
        # Step 2: Create the all_collection_days DataFrame (this is where issues might occur)
        print("\n2. CREATING ALL_COLLECTION_DAYS DATAFRAME")
        print("-" * 40)
        
        all_collection_days = []
        for date_obj, info in collection_day_map.items():
            all_collection_days.append({'_id': date_obj.strftime('%Y-%m-%d')})
        
        all_days_df = pd.DataFrame(all_collection_days)
        print(f"Created DataFrame with {len(all_days_df)} rows")
        
        # Check if early September dates are in the DataFrame
        early_sept_dates = [
            "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
            "2021-09-20", "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
        ]
        
        print("Early September dates in all_days_df:")
        for date_str in early_sept_dates:
            is_present = date_str in all_days_df['_id'].values
            print(f"  {date_str}: {'✅ Present' if is_present else '❌ Missing'}")
        
        # Step 3: Get actual data from database (simulate Daily Counts pipeline)
        print("\n3. GETTING ACTUAL DATA FROM DATABASE")
        print("-" * 40)
        
        db = get_db_connection()
        collection = db.media_records
        
        # Simulate the DAILY_COUNTS_COLLECTION_ONLY pipeline
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
        
        actual_data = list(collection.aggregate(pipeline))
        actual_df = pd.DataFrame(actual_data)
        
        print(f"Actual data contains {len(actual_df)} days")
        
        # Check early September in actual data
        print("Early September dates in actual data:")
        for date_str in early_sept_dates:
            matching_rows = actual_df[actual_df['_id'] == date_str]
            if not matching_rows.empty:
                count = matching_rows.iloc[0]['Total_Files']
                print(f"  {date_str}: ✅ {count} files")
            else:
                print(f"  {date_str}: ❌ No data")
        
        # Step 4: Perform the merge (this is where the magic happens)
        print("\n4. PERFORMING THE MERGE")
        print("-" * 40)
        
        print("Before merge:")
        print(f"  all_days_df: {len(all_days_df)} rows")
        print(f"  actual_df: {len(actual_df)} rows")
        
        # This is the exact merge from _fill_missing_collection_days
        merged_df = pd.merge(all_days_df, actual_df, on='_id', how='left').fillna(0)
        
        print(f"After merge: {len(merged_df)} rows")
        
        # Check early September in merged data
        print("Early September dates in merged data:")
        total_early_sept = 0
        for date_str in early_sept_dates:
            matching_rows = merged_df[merged_df['_id'] == date_str]
            if not matching_rows.empty:
                count = matching_rows.iloc[0]['Total_Files']
                total_early_sept += count
                print(f"  {date_str}: ✅ {count} files")
            else:
                print(f"  {date_str}: ❌ Missing from merge!")
        
        print(f"Total early September files in merged data: {total_early_sept}")
        
        # Step 5: Check the final sorted result
        print("\n5. FINAL SORTED RESULT")
        print("-" * 40)
        
        final_df = merged_df.sort_values('_id').reset_index(drop=True)
        
        print(f"Final DataFrame: {len(final_df)} rows")
        print("First 15 rows:")
        for i in range(min(15, len(final_df))):
            row = final_df.iloc[i]
            date_str = row['_id']
            count = row['Total_Files'] if 'Total_Files' in row else 0
            is_early_sept = date_str in early_sept_dates
            marker = "🎯" if is_early_sept else "  "
            print(f"  {i+1:2d}. {date_str}: {count:3.0f} files {marker}")
        
        # Calculate totals
        total_files = final_df['Total_Files'].sum() if 'Total_Files' in final_df.columns else 0
        print(f"\nTotal files in zero-filled data: {total_files}")
        
        return final_df, total_early_sept
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None, 0

def compare_with_excel_behavior():
    """Compare our simulation with what actually happens in Excel generation."""
    print("\n" + "=" * 60)
    print("COMPARISON WITH EXCEL BEHAVIOR")
    print("=" * 60)
    
    try:
        # Get the zero-filled data
        zero_filled_df, early_sept_total = simulate_zero_fill_process()
        
        if zero_filled_df is not None:
            print(f"Zero-fill simulation total: {zero_filled_df['Total_Files'].sum()}")
            print(f"Early September total: {early_sept_total}")
            
            # Expected Excel behavior
            expected_excel_total = 9372
            expected_early_sept = 358
            
            print(f"\nExpected Excel total: {expected_excel_total}")
            print(f"Expected early Sept excluded: {expected_early_sept}")
            
            # Analysis
            if zero_filled_df['Total_Files'].sum() == 9731:
                print("\n✅ Zero-fill includes ALL data (9731 files)")
                if early_sept_total == expected_early_sept:
                    print("✅ Early September count matches Excel observation")
                    print("🎯 CONCLUSION: Zero-fill is working correctly!")
                    print("   The issue must be in Excel DISPLAY/FORMATTING logic")
                else:
                    print(f"❓ Early September count mismatch: {early_sept_total} vs {expected_early_sept}")
            else:
                print(f"❌ Zero-fill total mismatch: {zero_filled_df['Total_Files'].sum()} vs 9731")
        
    except Exception as e:
        print(f"Error: {e}")

def main():
    try:
        simulate_zero_fill_process()
        compare_with_excel_behavior()
        
        print("\n" + "=" * 60)
        print("INVESTIGATION CONCLUSION")
        print("=" * 60)
        
        print("KEY FINDINGS:")
        print("1. ✅ Collection day map includes all early September dates")
        print("2. ✅ Zero-fill process should include all dates")
        print("3. ❓ Issue must be in Excel generation/display logic")
        
        print("\nNEXT STEPS:")
        print("1. Check if there's additional filtering AFTER zero-fill")
        print("2. Examine Excel cell formatting/alignment logic")
        print("3. Look for date range cutoffs in sheet creation")
        print("4. Check if totals are calculated differently than data display")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
