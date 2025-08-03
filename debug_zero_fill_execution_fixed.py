#!/usr/bin/env python3
"""
Debug Zero-Fill Execution - Fixed Version

This script investigates why our zero-fill fix isn't working by checking
the actual pipeline names and conditions used in the report generation.
"""

import sys
import os
import pandas as pd

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ar_utils import get_school_calendar, get_non_collection_days, precompute_collection_days
from db_utils import get_db_connection

def check_actual_pipeline_names():
    """Check what pipeline names are actually being used in the report generation."""
    print("CHECKING ACTUAL PIPELINE NAMES")
    print("=" * 50)
    
    try:
        import json
        
        # Read the report configuration
        with open('report_config.json', 'r') as f:
            config = json.load(f)
        
        print("ACF/PACF sheet configurations:")
        acf_pacf_sheets = []
        
        for sheet_name, sheet_config in config.get('sheets', {}).items():
            if 'ACF_PACF' in sheet_name or 'acf_pacf' in sheet_name.lower():
                pipeline_name = sheet_config.get('pipeline', 'Unknown')
                print(f"  {sheet_name}: {pipeline_name}")
                
                # Test if this pipeline would trigger zero-fill
                has_daily = 'DAILY' in pipeline_name.upper()
                has_with_zeroes = 'WITH_ZEROES' in pipeline_name.upper()
                has_collection_only = 'COLLECTION_ONLY' in pipeline_name.upper()
                condition_met = has_daily and (has_with_zeroes or has_collection_only)
                
                status = "✅ WILL trigger zero-fill" if condition_met else "❌ Will NOT trigger zero-fill"
                print(f"    {status}")
                
                acf_pacf_sheets.append({
                    'name': sheet_name,
                    'pipeline': pipeline_name,
                    'triggers_zero_fill': condition_met
                })
        
        return acf_pacf_sheets
        
    except Exception as e:
        print(f"Error reading config: {e}")
        return []

def test_zero_fill_logic_directly():
    """Test the zero-fill logic directly without the BaseSheetCreator class."""
    print("\n" + "=" * 50)
    print("TESTING ZERO-FILL LOGIC DIRECTLY")
    print("=" * 50)
    
    try:
        # Get sample data
        db = get_db_connection()
        collection = db.media_records
        
        # Simulate the Daily Counts (ACF_PACF) pipeline
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
        
        # Get the actual data
        actual_data = list(collection.aggregate(pipeline))
        df = pd.DataFrame(actual_data)
        
        print(f"Original data: {len(df)} days")
        print(f"Total files in original data: {df['Total_Files'].sum()}")
        
        # Apply the zero-fill logic manually
        school_calendar = get_school_calendar()
        non_collection_days = get_non_collection_days()
        collection_day_map = precompute_collection_days(school_calendar, non_collection_days)
        
        # Create all collection days DataFrame
        all_collection_days = []
        for date_obj, info in collection_day_map.items():
            all_collection_days.append({'_id': date_obj.strftime('%Y-%m-%d')})
        
        all_days_df = pd.DataFrame(all_collection_days)
        print(f"Collection day map: {len(all_days_df)} days")
        
        # Merge with existing data
        merged_df = pd.merge(all_days_df, df, on='_id', how='left').fillna(0)
        merged_df['Total_Files'] = merged_df['Total_Files'].astype(int)
        final_df = merged_df.sort_values('_id').reset_index(drop=True)
        
        print(f"Zero-filled data: {len(final_df)} days")
        print(f"Total files after zero-fill: {final_df['Total_Files'].sum()}")
        
        # Check early September dates
        early_sept_dates = [
            "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
            "2021-09-20", "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
        ]
        
        print("\nFirst 15 rows of zero-filled data:")
        for i in range(min(15, len(final_df))):
            row = final_df.iloc[i]
            date_str = row['_id']
            count = row['Total_Files']
            
            is_early_sept = date_str in early_sept_dates
            marker = "🎯" if is_early_sept else ""
            
            print(f"  {i+1:2d}. {date_str}: {count:3d} files {marker}")
        
        # Count early September files
        early_sept_count = 0
        for date_str in early_sept_dates:
            if date_str in final_df['_id'].values:
                row = final_df[final_df['_id'] == date_str]
                if not row.empty:
                    count = row.iloc[0]['Total_Files']
                    early_sept_count += count
        
        print(f"\nEarly September files: {early_sept_count}")
        print(f"Expected total if zero-fill works: {final_df['Total_Files'].sum()}")
        
        return final_df
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

def investigate_why_fix_not_working():
    """Investigate why our fix isn't working in the actual report generation."""
    print("\n" + "=" * 50)
    print("INVESTIGATING WHY FIX ISN'T WORKING")
    print("=" * 50)
    
    # Check if our modified method is actually being used
    try:
        from report_generator.sheet_creators.base import BaseSheetCreator
        import inspect
        
        # Get the source code of the method
        method = BaseSheetCreator._fill_missing_collection_days
        source_lines = inspect.getsourcelines(method)[0]
        
        # Check if our fix is present
        source_text = ''.join(source_lines)
        
        has_critical_fix = "CRITICAL FIX" in source_text
        has_debug_logging = "ZERO_FILL" in source_text
        has_early_sept_check = "early_sept_dates" in source_text
        
        print("Method analysis:")
        print(f"  Has 'CRITICAL FIX' comment: {has_critical_fix}")
        print(f"  Has debug logging: {has_debug_logging}")
        print(f"  Has early September check: {has_early_sept_check}")
        
        if has_critical_fix and has_debug_logging:
            print("  ✅ Our fix is present in the method")
            
            # The issue might be that the method isn't being called at all
            print("\n🤔 HYPOTHESIS:")
            print("The method contains our fix, but it might not be called")
            print("for the Daily Counts (ACF_PACF) sheet. Let's check why...")
            
        else:
            print("  ❌ Our fix is NOT present in the method")
            print("  The method might have been reverted or not saved properly")
        
        return True
        
    except Exception as e:
        print(f"Error analyzing method: {e}")
        return False

def main():
    print("DEBUG ZERO-FILL EXECUTION INVESTIGATION - FIXED")
    print("=" * 60)
    
    try:
        # Check actual pipeline names
        acf_pacf_sheets = check_actual_pipeline_names()
        
        # Test zero-fill logic directly
        zero_filled_df = test_zero_fill_logic_directly()
        
        # Investigate why our fix isn't working
        investigate_why_fix_not_working()
        
        print("\n" + "=" * 60)
        print("INVESTIGATION CONCLUSION")
        print("=" * 60)
        
        # Analyze the findings
        daily_sheet = None
        for sheet in acf_pacf_sheets:
            if 'Daily' in sheet['name']:
                daily_sheet = sheet
                break
        
        if daily_sheet:
            if daily_sheet['triggers_zero_fill']:
                print("🎯 PIPELINE CONDITION IS CORRECT")
                print(f"'{daily_sheet['name']}' uses '{daily_sheet['pipeline']}'")
                print("This SHOULD trigger zero-fill logic.")
                print()
                print("❓ POSSIBLE CAUSES:")
                print("1. The method isn't being called for this sheet")
                print("2. The method is being called but debug output is suppressed")
                print("3. There's an exception in the zero-fill logic")
                print("4. The zero-fill is working but Excel formatting is still wrong")
            else:
                print("🎯 FOUND THE ISSUE!")
                print(f"'{daily_sheet['name']}' uses '{daily_sheet['pipeline']}'")
                print("This does NOT trigger zero-fill logic.")
                print()
                print("SOLUTION: Update the pipeline name or condition")
        else:
            print("❌ Could not find Daily Counts (ACF_PACF) sheet in config")
        
        if zero_filled_df is not None:
            total_files = zero_filled_df['Total_Files'].sum()
            print(f"\n📊 ZERO-FILL TEST RESULT: {total_files} total files")
            if total_files == 9731:
                print("✅ Zero-fill logic produces correct total (9731)")
                print("The issue is that this logic isn't being applied to the Excel sheet")
            else:
                print(f"❌ Zero-fill logic produces incorrect total (expected 9731)")
                
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
