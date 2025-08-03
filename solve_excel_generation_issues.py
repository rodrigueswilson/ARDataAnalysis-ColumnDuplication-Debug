#!/usr/bin/env python3
"""
Comprehensive Excel Generation Issues Investigation

This script systematically investigates and documents the cluster of problems:
1. Left-aligned numbers in Excel sheets
2. The 359/358/1 discrepancy in file counts
3. Period start date ambiguity (2021-09-13 vs 2021-09-20)
4. Why only Daily and Weekly ACF/PACF sheets are affected

Goal: Identify the exact code path and logic causing these issues.
"""

import sys
import os
import json
from pathlib import Path

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection

def analyze_excel_generation_pipeline():
    """Trace the exact Excel generation pipeline for Daily Counts (ACF_PACF)."""
    print("1. EXCEL GENERATION PIPELINE ANALYSIS")
    print("=" * 60)
    
    try:
        # Load the report configuration
        config_path = Path(__file__).parent / "report_config.json"
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        # Find Daily Counts (ACF_PACF) configuration
        daily_counts_config = None
        for sheet in config.get('sheets', []):
            if sheet.get('name') == 'Daily Counts (ACF_PACF)':
                daily_counts_config = sheet
                break
        
        if daily_counts_config:
            print("Daily Counts (ACF_PACF) Configuration:")
            print(f"  Pipeline: {daily_counts_config.get('pipeline')}")
            print(f"  Sheet Name: {daily_counts_config.get('sheet_name')}")
            print(f"  Enabled: {daily_counts_config.get('enabled')}")
            print(f"  Full Config: {daily_counts_config}")
            
            # Get the pipeline details
            pipeline_name = daily_counts_config.get('pipeline')
            if pipeline_name:
                try:
                    from pipelines import PIPELINES
                    if pipeline_name in PIPELINES:
                        pipeline = PIPELINES[pipeline_name]
                        print(f"\nPipeline '{pipeline_name}' stages:")
                        for i, stage in enumerate(pipeline, 1):
                            stage_type = list(stage.keys())[0]
                            print(f"  Stage {i}: {stage_type}")
                            if stage_type == '$match':
                                print(f"    Filter: {stage[stage_type]}")
                            elif stage_type == '$group':
                                group_by = stage[stage_type].get('_id', 'Unknown')
                                print(f"    Group by: {group_by}")
                except ImportError as e:
                    print(f"  Could not load pipeline: {e}")
        else:
            print("❌ Daily Counts (ACF_PACF) configuration not found!")
        
        return daily_counts_config
        
    except Exception as e:
        print(f"Error analyzing configuration: {e}")
        return None

def trace_sheet_creation_logic():
    """Trace the sheet creation logic to find where left-alignment occurs."""
    print("\n2. SHEET CREATION LOGIC ANALYSIS")
    print("=" * 60)
    
    try:
        # The sheet creation happens in PipelineSheetCreator
        print("Sheet creation flow:")
        print("  1. report_config.json → sheet configuration")
        print("  2. PipelineSheetCreator._create_pipeline_sheet()")
        print("  3. Pipeline execution → raw data")
        print("  4. Zero-fill logic → time series completion")
        print("  5. Excel formatting → final output")
        
        print("\nKey questions to investigate:")
        print("  - Where does zero-fill logic start/end dates?")
        print("  - How are left-aligned rows created?")
        print("  - Why are they excluded from totals?")
        print("  - What's different about Daily/Weekly vs other sheets?")
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        return False

def investigate_zero_fill_date_logic():
    """Investigate the zero-fill date range logic."""
    print("\n3. ZERO-FILL DATE RANGE INVESTIGATION")
    print("=" * 60)
    
    try:
        # Check if there's a specific zero-fill start date
        from ar_utils import get_school_calendar, get_non_collection_days
        
        calendar = get_school_calendar()
        non_collection = get_non_collection_days()
        
        print("School Calendar Analysis:")
        for year, config in calendar.items():
            print(f"  {year}:")
            print(f"    Start: {config.get('start_date')}")
            print(f"    End: {config.get('end_date')}")
            
            periods = config.get('periods', {})
            for period, dates in periods.items():
                if isinstance(dates, list) and len(dates) >= 2:
                    print(f"    {period}: {dates[0]} to {dates[1]}")
        
        print(f"\nNon-collection days count: {len(non_collection)}")
        
        # Check specific dates around the problematic period
        problem_dates = [
            "2021-09-13", "2021-09-14", "2021-09-15", "2021-09-16", "2021-09-17",
            "2021-09-20", "2021-09-21", "2021-09-22", "2021-09-23", "2021-09-24", 
            "2021-09-27", "2021-09-28", "2021-09-29", "2021-09-30"
        ]
        
        print("\nProblem dates analysis:")
        for date_str in problem_dates:
            from datetime import datetime
            date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
            
            is_non_collection = date_obj in non_collection
            if is_non_collection:
                event_info = non_collection[date_obj]
                print(f"  {date_str}: {event_info.get('type', 'Unknown')} - {event_info.get('event', 'Unknown')}")
            else:
                print(f"  {date_str}: Regular collection day")
        
        return True
        
    except Exception as e:
        print(f"Error investigating zero-fill logic: {e}")
        return False

def find_left_alignment_code():
    """Try to find the specific code that creates left-aligned rows."""
    print("\n4. LEFT-ALIGNMENT CODE SEARCH")
    print("=" * 60)
    
    # The left-alignment is likely in one of these places:
    potential_locations = [
        "report_generator/sheet_creators/pipeline.py",
        "ar_utils.py",
        "utils/formatting.py",
        "report_generator/formatters.py"
    ]
    
    print("Potential code locations for left-alignment logic:")
    for location in potential_locations:
        file_path = Path(__file__).parent / location
        if file_path.exists():
            print(f"  ✅ {location} (exists)")
        else:
            print(f"  ❌ {location} (not found)")
    
    print("\nLikely causes of left-alignment:")
    print("  1. Data type mismatch (text vs numeric)")
    print("  2. Excel cell formatting differences")
    print("  3. Zero-fill vs actual data handling")
    print("  4. Date range filtering in sheet creation")
    print("  5. Conditional formatting based on date ranges")
    
    return True

def investigate_total_calculation_logic():
    """Investigate how Excel totals are calculated."""
    print("\n5. TOTAL CALCULATION LOGIC INVESTIGATION")
    print("=" * 60)
    
    try:
        # The total calculation likely happens in one of the sheet creators
        print("Total calculation analysis:")
        print("  - Excel totals are calculated during sheet creation")
        print("  - Left-aligned rows are excluded from totals")
        print("  - This suggests conditional logic based on date ranges")
        
        # Check if there's a specific total calculation method
        print("\nPotential total calculation methods:")
        print("  1. Excel SUM() formula with range")
        print("  2. Python calculation before Excel export")
        print("  3. Conditional totals based on date filters")
        print("  4. Zero-fill range exclusion logic")
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        return False

def test_date_range_hypothesis():
    """Test the hypothesis about date range cutoffs."""
    print("\n6. DATE RANGE CUTOFF HYPOTHESIS TEST")
    print("=" * 60)
    
    db = get_db_connection()
    collection = db.media_records
    
    # Test different potential cutoff dates
    cutoff_candidates = [
        "2021-09-20",  # After the partial day
        "2021-09-28",  # End of September
        "2021-10-01",  # Start of October
        "2021-10-15",  # Mid October
        "2021-10-31",  # End of October
        "2021-11-01"   # Start of November
    ]
    
    print("Testing potential cutoff dates:")
    
    for cutoff in cutoff_candidates:
        count_after = collection.count_documents({
            "ISO_Date": {"$gte": cutoff},
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        
        count_before = collection.count_documents({
            "ISO_Date": {"$lt": cutoff},
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        })
        
        total = count_before + count_after
        
        print(f"  {cutoff}:")
        print(f"    Before: {count_before:4d} files")
        print(f"    After:  {count_after:4d} files")
        print(f"    Total:  {total:4d} files")
        
        # Check if this matches our Excel total
        if count_after == 9372:
            print(f"    🎯 MATCH! Excel total = files after {cutoff}")
        elif count_before == 358:
            print(f"    🎯 MATCH! Left-aligned = files before {cutoff}")
    
    return True

def main():
    print("COMPREHENSIVE EXCEL GENERATION ISSUES INVESTIGATION")
    print("=" * 70)
    print("Solving the cluster of problems before ACF/PACF harmonization")
    print("=" * 70)
    
    try:
        # Run all investigations
        config = analyze_excel_generation_pipeline()
        sheet_logic = trace_sheet_creation_logic()
        zero_fill = investigate_zero_fill_date_logic()
        alignment_code = find_left_alignment_code()
        total_logic = investigate_total_calculation_logic()
        date_hypothesis = test_date_range_hypothesis()
        
        print("\n" + "=" * 70)
        print("INVESTIGATION SUMMARY")
        print("=" * 70)
        
        print("PROBLEMS TO SOLVE:")
        print("  1. ❓ Left-aligned numbers in Daily/Weekly ACF/PACF sheets")
        print("  2. ❓ 359/358/1 file count discrepancy")
        print("  3. ❓ Period start ambiguity (2021-09-13 vs 2021-09-20)")
        print("  4. ❓ Why only certain sheets are affected")
        
        print("\nNEXT STEPS:")
        print("  1. Examine PipelineSheetCreator code for left-alignment logic")
        print("  2. Trace zero-fill implementation in ar_utils.py")
        print("  3. Find the exact date cutoff causing the 9372 total")
        print("  4. Document the complete code path")
        print("  5. Propose and implement fixes")
        
        print("\nFILES TO EXAMINE:")
        print("  - report_generator/sheet_creators/pipeline.py")
        print("  - ar_utils.py (zero-fill logic)")
        print("  - config.yaml (date range configuration)")
        print("  - utils/formatting.py (Excel formatting)")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
