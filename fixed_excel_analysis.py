#!/usr/bin/env python3
"""
Fixed Excel Analysis for Daily Counts (ACF_PACF) Sheet

This script provides a robust analysis of the Excel output.
"""

import sys
import os
import pandas as pd
import glob

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def analyze_excel_output():
    """Analyze the Excel output with proper error handling."""
    print("FIXED EXCEL ANALYSIS FOR Daily Counts (ACF_PACF)")
    print("=" * 60)
    
    try:
        # Find latest Excel file
        excel_files = glob.glob("AR_Analysis_Report_*.xlsx")
        if not excel_files:
            print("❌ No Excel files found")
            return False
        
        latest_file = max(excel_files, key=os.path.getctime)
        print(f"📄 Latest file: {latest_file}")
        
        # Read Daily Counts (ACF_PACF) sheet with proper error handling
        try:
            df = pd.read_excel(latest_file, sheet_name='Daily Counts (ACF_PACF)', header=2)
            print(f"📊 Sheet dimensions: {df.shape}")
            print(f"📊 Columns: {list(df.columns)}")
            
            # Find the Total_Files column (might have different name)
            total_files_col = None
            for col in df.columns:
                if 'Total_Files' in str(col) or 'total' in str(col).lower():
                    total_files_col = col
                    break
            
            if total_files_col is not None:
                print(f"📈 Found total files column: {total_files_col}")
                
                # Convert to numeric, handling any string values
                df[total_files_col] = pd.to_numeric(df[total_files_col], errors='coerce')
                
                # Calculate totals
                total = df[total_files_col].sum()
                print(f"📈 Total files: {total}")
                
                # Check for zero values
                zero_count = (df[total_files_col] == 0).sum()
                non_zero_count = (df[total_files_col] > 0).sum()
                null_count = df[total_files_col].isnull().sum()
                
                print(f"📊 Days with files: {non_zero_count}")
                print(f"📊 Days with zero files: {zero_count}")
                print(f"📊 Days with null values: {null_count}")
                
                # Show some sample data
                print(f"\n📋 SAMPLE DATA (first 10 rows):")
                sample_cols = [col for col in df.columns if any(x in str(col) for x in ['Date', 'Total_Files', 'date', 'total'])][:3]
                if sample_cols:
                    print(df[sample_cols].head(10).to_string())
                
                # Date analysis if date column exists
                date_col = None
                for col in df.columns:
                    if 'date' in str(col).lower() or 'Date' in str(col):
                        date_col = col
                        break
                
                if date_col is not None:
                    print(f"\n📅 DATE ANALYSIS:")
                    print(f"📅 Date column: {date_col}")
                    
                    # Convert to datetime
                    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                    valid_dates = df[date_col].dropna()
                    
                    if not valid_dates.empty:
                        print(f"📅 Date range: {valid_dates.min()} to {valid_dates.max()}")
                        print(f"📅 Total days: {len(valid_dates)}")
                        
                        # Expected vs actual days
                        date_range = pd.date_range(valid_dates.min(), valid_dates.max(), freq='D')
                        expected_days = len(date_range)
                        actual_days = len(valid_dates)
                        print(f"📅 Expected days in range: {expected_days}")
                        print(f"📅 Actual days present: {actual_days}")
                        print(f"📅 Missing days: {expected_days - actual_days}")
                        
                        # Check if we have the expected 333 days
                        print(f"📅 Zero-fill target (333 days): {'✅ ACHIEVED' if actual_days >= 333 else '❌ MISSING'}")
            else:
                print("❌ No total files column found")
                print(f"Available columns: {list(df.columns)}")
        
        except Exception as e:
            print(f"❌ Error reading sheet: {e}")
            
            # Try reading without header
            try:
                df_raw = pd.read_excel(latest_file, sheet_name='Daily Counts (ACF_PACF)', header=None)
                print(f"📊 Raw sheet dimensions: {df_raw.shape}")
                print(f"📊 First few rows:")
                print(df_raw.head().to_string())
            except Exception as e2:
                print(f"❌ Error reading raw sheet: {e2}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error in analysis: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("COMPREHENSIVE EXCEL ANALYSIS")
    print("=" * 70)
    
    try:
        # Analyze Excel output
        if not analyze_excel_output():
            print("❌ Analysis failed")
            return 1
        
        print("\n" + "=" * 70)
        print("ANALYSIS COMPLETE")
        print("=" * 70)
        
        print("\n💡 SUMMARY:")
        print("- Debug logging shows DAILY_COUNTS_ALL_WITH_ZEROES pipeline is used")
        print("- Zero-fill method (_fill_missing_collection_days) is being called")
        print("- Sheet dimensions increased to 334 rows (from ~260)")
        print("- This suggests zero-fill logic is partially working")
        
        print("\n🎯 NEXT STEPS:")
        print("- Verify if we now have 333+ days (full zero-fill)")
        print("- Check if total files matches expected 9,731")
        print("- Confirm zero-fill logic is complete")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
