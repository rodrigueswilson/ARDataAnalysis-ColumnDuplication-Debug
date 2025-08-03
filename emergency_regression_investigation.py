#!/usr/bin/env python3
"""
Emergency Regression Investigation for Data Cleaning Sheet

This script investigates the critical regression where the Data Cleaning sheet
is now empty except for cell A1. This is likely caused by recent changes.
"""

import sys
import os
import pandas as pd
import glob
import json

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def analyze_data_cleaning_sheet():
    """Analyze the current state of the Data Cleaning sheet."""
    print("EMERGENCY REGRESSION INVESTIGATION")
    print("=" * 60)
    print("🚨 ANALYZING DATA CLEANING SHEET REGRESSION")
    print("=" * 60)
    
    try:
        # Find latest Excel file
        excel_files = glob.glob("AR_Analysis_Report_*.xlsx")
        if not excel_files:
            print("❌ No Excel files found")
            return False
        
        latest_file = max(excel_files, key=os.path.getctime)
        print(f"📄 Latest Excel file: {latest_file}")
        
        # Read Data Cleaning sheet
        try:
            df = pd.read_excel(latest_file, sheet_name='Data Cleaning', header=None)
            print(f"📊 Data Cleaning sheet dimensions: {df.shape}")
            
            # Check what's actually in the sheet
            print(f"\n📋 SHEET CONTENTS:")
            non_empty_cells = []
            for i in range(min(20, len(df))):
                for j in range(min(10, len(df.columns))):
                    cell_value = df.iloc[i, j]
                    if pd.notna(cell_value) and str(cell_value).strip():
                        non_empty_cells.append((i, j, str(cell_value)))
            
            print(f"📊 Non-empty cells found: {len(non_empty_cells)}")
            
            if non_empty_cells:
                print(f"📋 Non-empty cells:")
                for row, col, value in non_empty_cells[:10]:  # Show first 10
                    print(f"   Row {row}, Col {col}: {value}")
            else:
                print("❌ No non-empty cells found!")
            
            # Check if only A1 is populated
            a1_value = df.iloc[0, 0] if len(df) > 0 and len(df.columns) > 0 else None
            if a1_value and pd.notna(a1_value):
                print(f"📋 A1 value: {a1_value}")
                
                # Check if this is the only populated cell
                if len(non_empty_cells) == 1:
                    print("🚨 CONFIRMED: Only A1 is populated - CRITICAL REGRESSION")
                elif len(non_empty_cells) <= 3:
                    print("⚠️  SEVERE: Very few cells populated - MAJOR REGRESSION")
            
            return True
            
        except Exception as e:
            print(f"❌ Error reading Data Cleaning sheet: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Error in analysis: {e}")
        import traceback
        traceback.print_exc()
        return False

def check_report_config_for_data_cleaning():
    """Check the report configuration for Data Cleaning sheet."""
    print(f"\n🔍 CHECKING REPORT CONFIGURATION")
    print("=" * 50)
    
    try:
        with open('report_config.json', 'r') as f:
            config = json.load(f)
        
        # Find Data Cleaning sheet config
        data_cleaning_config = None
        for sheet in config.get('sheets', []):
            if 'Data Cleaning' in sheet.get('name', ''):
                data_cleaning_config = sheet
                break
        
        if data_cleaning_config:
            print(f"✅ Found Data Cleaning config:")
            print(f"   Name: {data_cleaning_config.get('name')}")
            print(f"   Module: {data_cleaning_config.get('module')}")
            print(f"   Pipeline: {data_cleaning_config.get('pipeline')}")
            print(f"   Enabled: {data_cleaning_config.get('enabled', True)}")
            
            # Check if it's enabled
            if not data_cleaning_config.get('enabled', True):
                print("🚨 FOUND ISSUE: Data Cleaning sheet is DISABLED!")
                return False
            
            return True
        else:
            print("❌ Data Cleaning sheet config not found!")
            return False
            
    except Exception as e:
        print(f"❌ Error checking config: {e}")
        return False

def identify_recent_changes():
    """Identify what changes might have caused this regression."""
    print(f"\n🔍 IDENTIFYING POTENTIAL CAUSES")
    print("=" * 50)
    
    potential_causes = [
        "Changes to base.py _run_aggregation_cached method",
        "Pipeline configuration changes in report_config.json", 
        "Debug logging additions affecting sheet creation",
        "Zero-fill logic changes affecting other sheets",
        "Import or module loading issues",
        "Cache key or pipeline name matching issues"
    ]
    
    print("🚨 POTENTIAL REGRESSION CAUSES:")
    for i, cause in enumerate(potential_causes, 1):
        print(f"   {i}. {cause}")
    
    # Check if our recent changes might have affected Data Cleaning
    print(f"\n🔍 RECENT CHANGES ANALYSIS:")
    
    # Check if we modified base.py recently
    try:
        with open('report_generator/sheet_creators/base.py', 'r') as f:
            base_content = f.read()
        
        # Look for our debug additions
        if '[AGGREGATION_DEBUG]' in base_content:
            print("⚠️  Found debug logging in base.py - potential cause")
        
        # Look for our zero-fill modifications
        if '_should_apply_zero_fill' in base_content:
            print("⚠️  Found zero-fill modifications in base.py - potential cause")
            
    except Exception as e:
        print(f"❌ Error checking base.py: {e}")
    
    return True

def suggest_immediate_fixes():
    """Suggest immediate fixes for the regression."""
    print(f"\n🔧 IMMEDIATE FIX SUGGESTIONS")
    print("=" * 50)
    
    fixes = [
        "1. Check if Data Cleaning sheet is disabled in report_config.json",
        "2. Remove debug logging from base.py _run_aggregation_cached method",
        "3. Verify Data Cleaning pipeline is still working",
        "4. Check for import errors in data cleaning module",
        "5. Test Data Cleaning sheet generation in isolation",
        "6. Roll back recent changes to base.py if necessary"
    ]
    
    print("🔧 RECOMMENDED IMMEDIATE ACTIONS:")
    for fix in fixes:
        print(f"   {fix}")
    
    print(f"\n⚡ PRIORITY ORDER:")
    print("   1. Check report_config.json (quickest fix)")
    print("   2. Remove debug logging (likely cause)")
    print("   3. Test Data Cleaning in isolation")
    print("   4. Roll back changes if needed")

def main():
    print("EMERGENCY REGRESSION INVESTIGATION")
    print("=" * 70)
    
    try:
        # Analyze current state
        sheet_analyzed = analyze_data_cleaning_sheet()
        
        # Check configuration
        config_checked = check_report_config_for_data_cleaning()
        
        # Identify causes
        causes_identified = identify_recent_changes()
        
        # Suggest fixes
        suggest_immediate_fixes()
        
        print("\n" + "=" * 70)
        print("EMERGENCY INVESTIGATION COMPLETE")
        print("=" * 70)
        
        if not sheet_analyzed or not config_checked:
            print("🚨 CRITICAL REGRESSION CONFIRMED")
            print("   Data Cleaning sheet is severely damaged")
            print("   Immediate action required")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error in investigation: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
