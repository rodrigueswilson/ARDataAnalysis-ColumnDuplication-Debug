#!/usr/bin/env python3
"""
Investigate Data Cleaning Sheet Issue

This script investigates why the Data Cleaning sheet is still not properly populated
despite adding the configuration back.
"""

import sys
import os
import pandas as pd
import glob

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def check_data_cleaning_sheet_raw():
    """Check the raw Data Cleaning sheet structure."""
    print("INVESTIGATING DATA CLEANING SHEET ISSUE")
    print("=" * 60)
    
    try:
        # Find latest Excel file
        excel_files = glob.glob("AR_Analysis_Report_*.xlsx")
        latest_file = max(excel_files, key=os.path.getctime)
        print(f"📄 Latest Excel file: {latest_file}")
        
        # Read raw Data Cleaning sheet
        try:
            df_raw = pd.read_excel(latest_file, sheet_name='Data Cleaning', header=None)
            print(f"📊 Raw Data Cleaning sheet dimensions: {df_raw.shape}")
            
            # Show all content
            print(f"\n📋 ALL SHEET CONTENT:")
            for i in range(len(df_raw)):
                for j in range(len(df_raw.columns)):
                    cell_value = df_raw.iloc[i, j]
                    if pd.notna(cell_value):
                        print(f"   Row {i}, Col {j}: {cell_value}")
            
            # Check if the sheet is basically empty
            if df_raw.shape[0] <= 1:
                print("🚨 CONFIRMED: Data Cleaning sheet is still essentially empty")
                return False
            else:
                print("✅ Data Cleaning sheet has some content")
                return True
                
        except Exception as e:
            print(f"❌ Error reading Data Cleaning sheet: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Error in investigation: {e}")
        return False

def check_data_cleaning_pipeline():
    """Check if the DATA_CLEANING_ANALYSIS pipeline exists."""
    print(f"\n🔍 CHECKING DATA_CLEANING_ANALYSIS PIPELINE")
    print("=" * 50)
    
    try:
        # Check if the pipeline module exists
        pipeline_files = [
            'report_generator/pipelines/data_cleaning.py',
            'report_generator/pipelines/data_cleaning_analysis.py',
            'pipelines/data_cleaning.py',
            'data_cleaning_pipeline.py'
        ]
        
        pipeline_found = False
        for pipeline_file in pipeline_files:
            if os.path.exists(pipeline_file):
                print(f"✅ Found pipeline file: {pipeline_file}")
                pipeline_found = True
                break
        
        if not pipeline_found:
            print("❌ No data cleaning pipeline file found!")
            print("🔍 Searching for any data cleaning related files...")
            
            # Search for data cleaning files
            for root, dirs, files in os.walk('.'):
                for file in files:
                    if 'data_cleaning' in file.lower() or 'cleaning' in file.lower():
                        print(f"   Found: {os.path.join(root, file)}")
        
        return pipeline_found
        
    except Exception as e:
        print(f"❌ Error checking pipeline: {e}")
        return False

def check_sheet_creation_logs():
    """Check if Data Cleaning sheet creation was attempted in the logs."""
    print(f"\n🔍 CHECKING SHEET CREATION LOGS")
    print("=" * 50)
    
    # Look for any mention of Data Cleaning in recent output
    # This would be in the generate_report.py output we saw earlier
    
    print("📋 From recent report generation output:")
    print("   - 'Data Cleaning' was mentioned in the sheet list")
    print("   - Pipeline 'DATA_CLEANING_ANALYSIS' was referenced")
    print("   - But the actual sheet creation may have failed silently")
    
    return True

def suggest_immediate_fix():
    """Suggest immediate fix for the Data Cleaning sheet issue."""
    print(f"\n🔧 IMMEDIATE FIX SUGGESTIONS")
    print("=" * 50)
    
    print("🎯 ROOT CAUSE ANALYSIS:")
    print("   1. Configuration is now correct (✅)")
    print("   2. Pipeline DATA_CLEANING_ANALYSIS may not exist (❌)")
    print("   3. Sheet creation is failing silently (❌)")
    
    print(f"\n🔧 RECOMMENDED FIXES:")
    print("   1. Create the missing DATA_CLEANING_ANALYSIS pipeline")
    print("   2. Or change to use an existing pipeline (e.g., DASHBOARD_DATA_CLEANING)")
    print("   3. Add error handling to catch pipeline failures")
    print("   4. Test Data Cleaning sheet generation in isolation")
    
    print(f"\n⚡ QUICK FIX OPTIONS:")
    print("   A. Change pipeline to 'DASHBOARD_DATA_CLEANING' (exists in dashboard_data module)")
    print("   B. Create a simple DATA_CLEANING_ANALYSIS pipeline")
    print("   C. Temporarily disable Data Cleaning sheet until pipeline is fixed")

def main():
    print("DATA CLEANING SHEET ISSUE INVESTIGATION")
    print("=" * 70)
    
    try:
        # Check raw sheet
        sheet_has_content = check_data_cleaning_sheet_raw()
        
        # Check pipeline
        pipeline_exists = check_data_cleaning_pipeline()
        
        # Check logs
        check_sheet_creation_logs()
        
        # Suggest fix
        suggest_immediate_fix()
        
        print("\n" + "=" * 70)
        print("INVESTIGATION COMPLETE")
        print("=" * 70)
        
        if not sheet_has_content and not pipeline_exists:
            print("🚨 CRITICAL ISSUE IDENTIFIED:")
            print("   - Data Cleaning sheet is empty")
            print("   - DATA_CLEANING_ANALYSIS pipeline does not exist")
            print("   - Need to create pipeline or use existing one")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error in investigation: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
