#!/usr/bin/env python3
"""
Debug Totals Implementation Script
=================================

This script debugs why some high-priority sheets have totals while others don't.
It traces the totals configuration and application process.
"""

import json
import pandas as pd
from pathlib import Path
from report_generator.sheet_creators.pipeline import PipelineSheetCreator
from report_generator.totals_manager import TotalsManager
from report_generator.formatters import ExcelFormatter
from db_utils import get_db_connection
from pipelines import PIPELINES

def debug_totals_config():
    """
    Debug the totals configuration for high-priority sheets.
    """
    print("🔍 DEBUGGING TOTALS CONFIGURATION")
    print("="*50)
    
    # Initialize components
    db = get_db_connection()
    formatter = ExcelFormatter()
    pipeline_creator = PipelineSheetCreator(db, formatter)
    
    # High-priority sheets to debug
    debug_sheets = {
        'Monthly Capture Volume': 'CAPTURE_VOLUME_PER_MONTH',
        'File Size Stats': 'FILE_SIZE_STATISTICS',
        'Time of Day': 'TIME_OF_DAY_DISTRIBUTION',
        'Weekday by Period': 'WEEKDAY_BY_PERIOD'
    }
    
    for sheet_name, pipeline_name in debug_sheets.items():
        print(f"\n🔍 Debugging: {sheet_name}")
        print(f"   Pipeline: {pipeline_name}")
        
        # Check if pipeline exists
        if pipeline_name not in PIPELINES:
            print(f"   ❌ ERROR: Pipeline '{pipeline_name}' not found in PIPELINES")
            continue
        
        # Get pipeline data
        try:
            pipeline = PIPELINES[pipeline_name]
            result = list(db.media_records.aggregate(pipeline))
            df = pd.DataFrame(result)
            
            print(f"   📊 Data: {len(df)} rows, {len(df.columns)} columns")
            print(f"   📋 Columns: {list(df.columns)}")
            
            # Test totals configuration
            sheet_type = 'unknown'  # Most of these aren't time series
            config = pipeline_creator._get_totals_config_for_sheet(sheet_name, df, sheet_type)
            
            if config:
                print(f"   ✅ Config generated:")
                print(f"      - add_totals: {config.get('add_totals', False)}")
                print(f"      - add_row_totals: {config.get('add_row_totals', False)}")
                print(f"      - add_column_totals: {config.get('add_column_totals', False)}")
                print(f"      - add_grand_total: {config.get('add_grand_total', False)}")
                if 'rationale' in config:
                    print(f"      - rationale: {config['rationale']}")
            else:
                print(f"   ❌ No configuration generated")
                
        except Exception as e:
            print(f"   ❌ ERROR processing {sheet_name}: {e}")

def test_individual_totals_application():
    """
    Test applying totals to individual sheets manually.
    """
    print("\n🧪 TESTING INDIVIDUAL TOTALS APPLICATION")
    print("="*50)
    
    # Test with Monthly Capture Volume (the one that's failing)
    db = get_db_connection()
    
    try:
        pipeline = PIPELINES['CAPTURE_VOLUME_PER_MONTH']
        result = list(db.media_records.aggregate(pipeline))
        df = pd.DataFrame(result)
        
        print(f"📊 Monthly Capture Volume data:")
        print(f"   Rows: {len(df)}")
        print(f"   Columns: {list(df.columns)}")
        print(f"   Sample data:")
        print(df.head(3).to_string())
        
        # Test totals manager directly
        totals_manager = TotalsManager()
        
        # Test configuration
        config = {
            'add_row_totals': False,
            'add_column_totals': True,
            'add_grand_total': True,
            'totals_label': 'TOTAL',
            'add_totals': True
        }
        
        print(f"\n🔧 Testing totals configuration:")
        print(f"   Config: {config}")
        
        # Identify numeric columns
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        print(f"   Numeric columns: {numeric_cols}")
        
        if numeric_cols:
            # Calculate column totals manually
            column_totals = totals_manager.calculate_column_totals(df, numeric_columns=numeric_cols)
            print(f"   ✅ Column totals calculated: {dict(column_totals)}")
        else:
            print(f"   ⚠️  No numeric columns found for totals")
            
    except Exception as e:
        print(f"❌ Error testing Monthly Capture Volume: {e}")
        import traceback
        traceback.print_exc()

def check_pipeline_sheet_integration():
    """
    Check if the pipeline sheet creator is properly calling totals methods.
    """
    print("\n🔍 CHECKING PIPELINE SHEET INTEGRATION")
    print("="*50)
    
    # Load report configuration
    config_path = Path("report_config.json")
    with open(config_path, 'r') as f:
        config = json.load(f)
    
    # Find Monthly Capture Volume configuration
    monthly_config = None
    for sheet_config in config.get('sheets', []):
        if sheet_config.get('name') == 'Monthly Capture Volume':
            monthly_config = sheet_config
            break
    
    if monthly_config:
        print(f"✅ Found Monthly Capture Volume in report config:")
        print(f"   Pipeline: {monthly_config.get('pipeline')}")
        print(f"   Enabled: {monthly_config.get('enabled')}")
        print(f"   Category: {monthly_config.get('category')}")
    else:
        print(f"❌ Monthly Capture Volume not found in report config")
    
    # Check if totals are being applied in the pipeline creation process
    print(f"\n🔍 Checking totals application process...")
    
    # Look for the specific code path in pipeline sheet creator
    try:
        db = get_db_connection()
        formatter = ExcelFormatter()
        creator = PipelineSheetCreator(db, formatter)
        
        # Check if totals manager is initialized
        if hasattr(creator, 'totals_manager'):
            print(f"   ✅ TotalsManager initialized in PipelineSheetCreator")
        else:
            print(f"   ❌ TotalsManager not found in PipelineSheetCreator")
            
        # Check if the _get_totals_config_for_sheet method exists
        if hasattr(creator, '_get_totals_config_for_sheet'):
            print(f"   ✅ _get_totals_config_for_sheet method exists")
        else:
            print(f"   ❌ _get_totals_config_for_sheet method missing")
            
    except Exception as e:
        print(f"   ❌ Error checking integration: {e}")

def main():
    """
    Main debugging function.
    """
    print("🚀 STARTING TOTALS IMPLEMENTATION DEBUGGING")
    print("="*60)
    
    try:
        # Debug totals configuration
        debug_totals_config()
        
        # Test individual totals application
        test_individual_totals_application()
        
        # Check pipeline sheet integration
        check_pipeline_sheet_integration()
        
        print("\n" + "="*60)
        print("🎯 DEBUGGING SUMMARY")
        print("="*60)
        print("1. Check if HIGH-PRIORITY messages appear in the output above")
        print("2. Verify that totals configuration is being generated correctly")
        print("3. Ensure numeric columns are being detected properly")
        print("4. Confirm TotalsManager integration in PipelineSheetCreator")
        
        print("\n🔧 POTENTIAL FIXES:")
        print("- If no HIGH-PRIORITY messages: Configuration not being triggered")
        print("- If no numeric columns: Data type detection issue")
        print("- If config generated but no totals: Integration issue in sheet creation")
        
    except Exception as e:
        print(f"❌ Debugging error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
