#!/usr/bin/env python3
"""
Trace ACF/PACF Sheet Creation

This script traces exactly how the ACF/PACF sheets are being created
to understand why our zero-fill fix isn't being applied.
"""

import sys
import os
import json

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def analyze_report_config():
    """Analyze the report configuration to understand sheet creation flow."""
    print("ANALYZING REPORT CONFIGURATION")
    print("=" * 50)
    
    try:
        with open('report_config.json', 'r') as f:
            config = json.load(f)
        
        # Find ACF/PACF sheets
        acf_pacf_sheets = []
        if 'sheets' in config:
            for sheet in config['sheets']:
                if isinstance(sheet, dict) and 'ACF_PACF' in sheet.get('name', ''):
                    acf_pacf_sheets.append(sheet)
        
        print(f"Found {len(acf_pacf_sheets)} ACF/PACF sheets:")
        for sheet in acf_pacf_sheets:
            print(f"  - {sheet.get('name', 'Unknown')}")
            print(f"    Pipeline: {sheet.get('pipeline', 'Unknown')}")
            print(f"    Module: {sheet.get('module', 'Unknown')}")
            print(f"    Category: {sheet.get('category', 'Unknown')}")
            print(f"    Enabled: {sheet.get('enabled', 'Unknown')}")
            print()
        
        return acf_pacf_sheets
        
    except Exception as e:
        print(f"Error reading config: {e}")
        return []

def trace_sheet_creation_methods():
    """Trace all possible sheet creation methods."""
    print("TRACING SHEET CREATION METHODS")
    print("=" * 50)
    
    # Check if sheets are created via SheetCreator
    try:
        from report_generator.sheet_creators import SheetCreator
        print("✅ SheetCreator class found")
        
        # Check methods
        methods = [method for method in dir(SheetCreator) if not method.startswith('_')]
        print(f"Public methods: {len(methods)}")
        
        # Look for ACF/PACF specific methods
        acf_methods = [method for method in methods if 'acf' in method.lower() or 'pacf' in method.lower()]
        if acf_methods:
            print(f"ACF/PACF methods: {acf_methods}")
        else:
            print("No ACF/PACF specific methods found")
            
    except Exception as e:
        print(f"Error importing SheetCreator: {e}")
    
    # Check if sheets are created via core ReportGenerator
    try:
        from report_generator.core import ReportGenerator
        print("✅ ReportGenerator class found")
        
        # Check methods
        methods = [method for method in dir(ReportGenerator) if not method.startswith('_')]
        print(f"Public methods: {len(methods)}")
        
    except Exception as e:
        print(f"Error importing ReportGenerator: {e}")

def check_pipeline_execution():
    """Check if pipelines are actually being executed."""
    print("\nCHECKING PIPELINE EXECUTION")
    print("=" * 50)
    
    try:
        from db_utils import get_db_connection
        
        db = get_db_connection()
        collection = db.media_records
        
        # Test the DAILY_COUNTS_COLLECTION_ONLY pipeline directly
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
        
        result = list(collection.aggregate(pipeline))
        print(f"DAILY_COUNTS_COLLECTION_ONLY pipeline returns: {len(result)} days")
        print(f"Total files: {sum(row['Total_Files'] for row in result)}")
        
        # Check first few dates
        print("First 10 dates:")
        for i, row in enumerate(result[:10]):
            print(f"  {i+1}. {row['_id']}: {row['Total_Files']} files")
        
        return len(result), sum(row['Total_Files'] for row in result)
        
    except Exception as e:
        print(f"Error executing pipeline: {e}")
        return 0, 0

def investigate_sheet_creation_bypass():
    """Investigate why sheet creation bypasses our zero-fill logic."""
    print("\nINVESTIGATING SHEET CREATION BYPASS")
    print("=" * 50)
    
    print("HYPOTHESIS:")
    print("1. ACF/PACF sheets are created via a different code path")
    print("2. The zero-fill method is never called for these sheets")
    print("3. The sheets use raw pipeline data without zero-fill processing")
    print()
    
    print("EVIDENCE:")
    print("- Our debug messages don't appear in report generation output")
    print("- Excel verification shows incorrect totals and structure")
    print("- Left-aligned numbers still present")
    print()
    
    print("POSSIBLE CAUSES:")
    print("1. Sheets created via core.py ReportGenerator instead of PipelineSheetCreator")
    print("2. Sheets created via specialized sheet creator that bypasses zero-fill")
    print("3. Configuration issue causing sheets to use wrong creation path")
    print("4. Method inheritance issue preventing our fix from being used")

def main():
    print("TRACE ACF/PACF SHEET CREATION")
    print("=" * 60)
    
    try:
        # Analyze configuration
        acf_pacf_sheets = analyze_report_config()
        
        # Trace creation methods
        trace_sheet_creation_methods()
        
        # Check pipeline execution
        days, total_files = check_pipeline_execution()
        
        # Investigate bypass
        investigate_sheet_creation_bypass()
        
        print("\n" + "=" * 60)
        print("INVESTIGATION SUMMARY")
        print("=" * 60)
        
        print("KEY FINDINGS:")
        print(f"1. Found {len(acf_pacf_sheets)} ACF/PACF sheets in configuration")
        print(f"2. Pipeline execution returns {days} days with {total_files} total files")
        print("3. Debug messages not appearing suggests bypass of PipelineSheetCreator")
        print("4. Zero-fill logic not being applied to ACF/PACF sheets")
        
        print("\nNEXT STEPS:")
        print("1. Find the actual code path used for ACF/PACF sheet creation")
        print("2. Apply zero-fill fix to the correct creation method")
        print("3. Ensure all ACF/PACF sheets use consistent zero-fill logic")
        print("4. Test fix with new report generation")
        
        if total_files == 9731:
            print(f"\n✅ Pipeline data is correct ({total_files} files)")
            print("The issue is in sheet creation, not data retrieval")
        else:
            print(f"\n❌ Pipeline data is incorrect (expected 9731, got {total_files})")
            print("The issue may be in both data retrieval and sheet creation")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
