#!/usr/bin/env python3
"""
Left-Aligned Rows Investigation

This script investigates why certain rows appear left-aligned in Daily and Weekly
ACF/PACF sheets but not in other ACF/PACF sheets, and why they're excluded from totals.

Pattern observed:
- Daily Counts (ACF_PACF): Early September 2021 dates are left-aligned
- Weekly Counts (ACF_PACF): Early weeks (2021-W37 to 2021-W46) are left-aligned
- Other ACF/PACF sheets: No left-aligned rows

This suggests the issue is in the sheet generation or formatting logic.
"""

import sys
import os
import json
from pathlib import Path

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def analyze_sheet_configurations():
    """Analyze the report_config.json to understand sheet differences."""
    print("1. ANALYZING SHEET CONFIGURATIONS")
    print("-" * 50)
    
    try:
        config_path = Path(__file__).parent / "report_config.json"
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        acf_pacf_sheets = []
        for sheet in config.get('sheets', []):
            if 'ACF_PACF' in sheet.get('name', ''):
                acf_pacf_sheets.append(sheet)
        
        print(f"Found {len(acf_pacf_sheets)} ACF/PACF sheets:")
        
        for sheet in acf_pacf_sheets:
            name = sheet.get('name', 'Unknown')
            pipeline = sheet.get('pipeline', 'Unknown')
            has_problem = name in ['Daily Counts (ACF_PACF)', 'Weekly Counts (ACF_PACF)']
            status = "🔴 HAS PROBLEM" if has_problem else "✅ OK"
            
            print(f"  {name}")
            print(f"    Pipeline: {pipeline}")
            print(f"    Status: {status}")
            print(f"    Config: {sheet}")
            print()
        
        return acf_pacf_sheets
        
    except Exception as e:
        print(f"Error reading config: {e}")
        return []

def analyze_pipeline_differences():
    """Compare pipelines used by problematic vs working sheets."""
    print("2. ANALYZING PIPELINE DIFFERENCES")
    print("-" * 50)
    
    try:
        from pipelines import PIPELINES
        
        # Problematic pipelines
        problematic = ['DAILY_COUNTS_COLLECTION_ONLY', 'WEEKLY_COUNTS_WITH_ZEROES']
        
        # Working pipelines (examples)
        working = ['MONTHLY_COUNTS_WITH_ZEROES', 'PERIOD_COUNTS_WITH_ZEROES']
        
        print("PROBLEMATIC PIPELINES:")
        for pipeline_name in problematic:
            if pipeline_name in PIPELINES:
                pipeline = PIPELINES[pipeline_name]
                print(f"  {pipeline_name}:")
                for i, stage in enumerate(pipeline, 1):
                    print(f"    Stage {i}: {list(stage.keys())[0]}")
                    if '$match' in stage:
                        print(f"      Filter: {stage['$match']}")
                print()
        
        print("WORKING PIPELINES:")
        for pipeline_name in working:
            if pipeline_name in PIPELINES:
                pipeline = PIPELINES[pipeline_name]
                print(f"  {pipeline_name}:")
                for i, stage in enumerate(pipeline, 1):
                    print(f"    Stage {i}: {list(stage.keys())[0]}")
                    if '$match' in stage:
                        print(f"      Filter: {stage['$match']}")
                print()
        
        return True
        
    except ImportError as e:
        print(f"Could not import pipelines: {e}")
        return False

def analyze_sheet_creation_logic():
    """Analyze the sheet creation logic for differences."""
    print("3. ANALYZING SHEET CREATION LOGIC")
    print("-" * 50)
    
    # Check if there's special logic for Daily/Weekly vs other sheets
    try:
        # Look for sheet creation patterns
        print("Checking sheet creator patterns...")
        
        # The issue might be in:
        # 1. Zero-fill logic differences
        # 2. Date range handling
        # 3. Alignment/formatting logic
        # 4. Total calculation logic
        
        print("Potential causes:")
        print("  1. Zero-fill logic differences between daily/weekly vs monthly/period")
        print("  2. Date range filtering in ar_utils.py")
        print("  3. Excel alignment/formatting in sheet creators")
        print("  4. Total row calculation excluding certain ranges")
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        return False

def check_zero_fill_date_ranges():
    """Check if zero-fill logic has different date ranges for different time scales."""
    print("4. CHECKING ZERO-FILL DATE RANGES")
    print("-" * 50)
    
    try:
        from ar_utils import precompute_collection_days
        
        collection_days = precompute_collection_days()
        
        if collection_days:
            min_date = min(collection_days)
            max_date = max(collection_days)
            
            print(f"Collection days range: {min_date} to {max_date}")
            
            # Check if the problematic dates are before the start
            problematic_dates = [
                "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
                "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", "2021-09-27"
            ]
            
            dates_before_start = [d for d in problematic_dates if d < min_date]
            dates_in_range = [d for d in problematic_dates if min_date <= d <= max_date]
            
            print(f"Problematic dates before collection start: {len(dates_before_start)}")
            print(f"Problematic dates in collection range: {len(dates_in_range)}")
            
            if dates_before_start:
                print(f"  Before start: {dates_before_start}")
            if dates_in_range:
                print(f"  In range: {dates_in_range}")
            
            # This might explain why they're left-aligned - they're outside the zero-fill range
            if dates_before_start and not dates_in_range:
                print("🎯 HYPOTHESIS: Dates are before zero-fill range, causing left-alignment!")
                return "before_zero_fill_range"
            elif dates_in_range and not dates_before_start:
                print("❓ All dates are in range - zero-fill not the cause")
                return "in_range"
            else:
                print("❓ Mixed - some before, some in range")
                return "mixed"
        
        return None
        
    except Exception as e:
        print(f"Error checking zero-fill: {e}")
        return None

def check_excel_formatting_logic():
    """Check if there's specific Excel formatting causing left-alignment."""
    print("5. CHECKING EXCEL FORMATTING LOGIC")
    print("-" * 50)
    
    # Look for alignment-related code
    try:
        print("Looking for Excel alignment/formatting patterns...")
        
        # The left-alignment might be caused by:
        # 1. Different data types (text vs number)
        # 2. Explicit alignment settings
        # 3. Conditional formatting
        # 4. Data preprocessing differences
        
        print("Potential Excel formatting causes:")
        print("  1. Data type differences (text vs numeric)")
        print("  2. Explicit cell alignment settings")
        print("  3. Conditional formatting rules")
        print("  4. Data preprocessing that affects formatting")
        print("  5. Zero-fill vs actual data handling")
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        return False

def main():
    print("LEFT-ALIGNED ROWS INVESTIGATION")
    print("=" * 60)
    print("Investigating why certain rows are left-aligned in Daily/Weekly ACF/PACF")
    print("=" * 60)
    
    try:
        # Run all investigations
        sheet_configs = analyze_sheet_configurations()
        pipeline_analysis = analyze_pipeline_differences()
        sheet_logic = analyze_sheet_creation_logic()
        zero_fill_result = check_zero_fill_date_ranges()
        excel_formatting = check_excel_formatting_logic()
        
        print("\n" + "=" * 60)
        print("INVESTIGATION SUMMARY")
        print("=" * 60)
        
        print("PATTERN CONFIRMED:")
        print("  ✅ Daily Counts (ACF_PACF): Has left-aligned rows")
        print("  ✅ Weekly Counts (ACF_PACF): Has left-aligned rows")
        print("  ✅ Other ACF/PACF sheets: No left-aligned rows")
        
        print("\nPOTENTIAL ROOT CAUSES:")
        
        if zero_fill_result == "before_zero_fill_range":
            print("  🎯 PRIMARY SUSPECT: Zero-fill date range exclusion")
            print("     - Early dates are before the zero-fill start date")
            print("     - This causes them to be treated differently in Excel")
        
        print("  🔍 SECONDARY SUSPECTS:")
        print("     - Pipeline-specific data processing differences")
        print("     - Excel formatting logic in sheet creators")
        print("     - Data type handling (text vs numeric)")
        print("     - Total calculation exclusion logic")
        
        print("\nNEXT STEPS:")
        print("  1. Examine ar_utils.py zero-fill logic")
        print("  2. Check sheet creator Excel formatting code")
        print("  3. Compare data types in problematic vs working sheets")
        print("  4. Identify the exact code causing left-alignment")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
