#!/usr/bin/env python3
"""
Debug Sheet Data Types
=====================

This script investigates why the Phase 2 sheets are not showing numeric columns
in the validation, despite having totals.
"""

import pandas as pd
from pathlib import Path

def find_latest_report():
    """Find the most recent Excel report file."""
    report_files = list(Path('.').glob('AR_Analysis_Report_*.xlsx'))
    if not report_files:
        raise FileNotFoundError("No Excel report files found")
    
    latest_file = max(report_files, key=lambda f: f.stat().st_mtime)
    return latest_file

def debug_phase2_sheets():
    """
    Debug the Phase 2 sheets to understand data type issues.
    """
    print("🔍 DEBUGGING PHASE 2 SHEET DATA TYPES")
    print("="*50)
    
    phase2_sheets = [
        'Weekly Counts',
        'Day of Week Counts', 
        'Activity Counts',
        'File Size by Day',
        'SY Days',
        'Data Quality'
    ]
    
    try:
        report_file = find_latest_report()
        print(f"📊 Analyzing report: {report_file}")
        
        for sheet_name in phase2_sheets:
            print(f"\n📋 DEBUGGING: {sheet_name}")
            print("-" * 30)
            
            try:
                # Read the sheet
                df = pd.read_excel(report_file, sheet_name=sheet_name)
                
                print(f"   📏 Shape: {df.shape}")
                print(f"   📊 Columns: {list(df.columns)}")
                
                # Check data types
                print(f"   🔢 Data types:")
                for col in df.columns:
                    dtype = df[col].dtype
                    print(f"      {col}: {dtype}")
                
                # Check for numeric columns using different methods
                numeric_cols_default = df.select_dtypes(include=['number']).columns.tolist()
                numeric_cols_int_float = df.select_dtypes(include=['int64', 'float64', 'int32', 'float32']).columns.tolist()
                
                print(f"   🔢 Numeric columns (default): {numeric_cols_default}")
                print(f"   🔢 Numeric columns (int/float): {numeric_cols_int_float}")
                
                # Check last row for totals
                if len(df) > 0:
                    last_row = df.iloc[-1]
                    print(f"   📝 Last row values:")
                    for col in df.columns:
                        value = last_row[col]
                        print(f"      {col}: {value} (type: {type(value)})")
                
                # Check if we can convert text columns to numeric
                convertible_cols = []
                for col in df.columns:
                    if col not in numeric_cols_default:
                        try:
                            # Try to convert to numeric, excluding the last row which might have "TOTAL"
                            data_rows = df[col].iloc[:-1] if len(df) > 1 else df[col]
                            pd.to_numeric(data_rows, errors='raise')
                            convertible_cols.append(col)
                        except:
                            pass
                
                print(f"   🔄 Convertible to numeric: {convertible_cols}")
                
                # Sample a few data values
                if len(df) > 1:
                    print(f"   📋 Sample data (first 3 rows):")
                    for i in range(min(3, len(df)-1)):  # Exclude last row (totals)
                        print(f"      Row {i}: {dict(df.iloc[i])}")
                
            except Exception as e:
                print(f"   🚨 ERROR reading {sheet_name}: {e}")
    
    except Exception as e:
        print(f"❌ Error: {e}")

def test_manual_validation():
    """
    Test manual validation of one sheet to understand the issue.
    """
    print(f"\n🧪 MANUAL VALIDATION TEST")
    print("="*30)
    
    try:
        report_file = find_latest_report()
        
        # Test with Weekly Counts
        sheet_name = 'Weekly Counts'
        print(f"📋 Testing: {sheet_name}")
        
        # Try different read methods
        print(f"\n📖 Method 1: Default read_excel")
        df1 = pd.read_excel(report_file, sheet_name=sheet_name)
        print(f"   Shape: {df1.shape}")
        print(f"   Numeric columns: {df1.select_dtypes(include=['number']).columns.tolist()}")
        
        print(f"\n📖 Method 2: read_excel with dtype inference")
        df2 = pd.read_excel(report_file, sheet_name=sheet_name, dtype=None)
        print(f"   Shape: {df2.shape}")
        print(f"   Numeric columns: {df2.select_dtypes(include=['number']).columns.tolist()}")
        
        print(f"\n📖 Method 3: Manual conversion")
        df3 = pd.read_excel(report_file, sheet_name=sheet_name)
        for col in df3.columns:
            # Try to convert each column to numeric
            try:
                # Convert all but last row (which might have TOTAL)
                if len(df3) > 1:
                    data_part = df3[col].iloc[:-1]
                    totals_part = df3[col].iloc[-1:]
                    
                    # Try to convert data part
                    converted_data = pd.to_numeric(data_part, errors='coerce')
                    if not converted_data.isna().all():
                        print(f"   ✅ {col}: Can be converted to numeric")
                        # Replace the column with converted values
                        df3[col] = pd.concat([converted_data, totals_part])
                    else:
                        print(f"   ❌ {col}: Cannot be converted to numeric")
                else:
                    converted = pd.to_numeric(df3[col], errors='coerce')
                    if not converted.isna().all():
                        print(f"   ✅ {col}: Can be converted to numeric")
                        df3[col] = converted
                    else:
                        print(f"   ❌ {col}: Cannot be converted to numeric")
            except Exception as e:
                print(f"   🚨 {col}: Error converting - {e}")
        
        print(f"   Final numeric columns: {df3.select_dtypes(include=['number']).columns.tolist()}")
        
    except Exception as e:
        print(f"❌ Manual test error: {e}")

def main():
    """
    Main debugging function.
    """
    debug_phase2_sheets()
    test_manual_validation()

if __name__ == "__main__":
    main()
