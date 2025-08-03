#!/usr/bin/env python3
"""
Verification Script: Daily Counts (ACF_PACF) Data Cleaning Logic Fix
==================================================================

This script verifies that the Daily Counts (ACF_PACF) sheet now matches
the Data Cleaning sheet totals after applying the complete filtering logic.

Expected behavior after fix:
- Both sheets should use identical filtering: School_Year ≠ "N/A", 
  file_type ∈ ["JPG", "MP3"], is_collection_day: TRUE, Outlier_Status: FALSE
- Total file counts should match between sheets
"""

import pandas as pd
import os
from datetime import datetime

def verify_daily_counts_fix():
    """
    Verify that Daily Counts (ACF_PACF) totals match Data Cleaning totals
    after applying the complete data cleaning logic.
    """
    
    # Find the latest report
    report_files = [f for f in os.listdir('.') if f.startswith('AR_Analysis_Report_') and f.endswith('.xlsx')]
    if not report_files:
        print("❌ No Excel reports found!")
        return False
    
    latest_report = max(report_files)
    print(f"📊 Analyzing report: {latest_report}")
    print("=" * 60)
    
    try:
        # Read Data Cleaning sheet
        print("📋 Reading Data Cleaning sheet...")
        data_cleaning_df = pd.read_excel(latest_report, sheet_name='Data Cleaning', header=None)
        
        # Find the clean research data totals in Data Cleaning sheet
        clean_totals = {}
        for i in range(len(data_cleaning_df)):
            for j in range(len(data_cleaning_df.columns)):
                cell_value = str(data_cleaning_df.iloc[i, j])
                if 'Clean Research Data' in cell_value or 'Final Dataset' in cell_value:
                    # Look for totals in nearby cells
                    for offset in range(1, 5):
                        if i + offset < len(data_cleaning_df):
                            try:
                                total_value = data_cleaning_df.iloc[i + offset, j]
                                if pd.notna(total_value) and str(total_value).replace(',', '').isdigit():
                                    clean_totals['total_files'] = int(str(total_value).replace(',', ''))
                                    break
                            except:
                                continue
        
        # Also look for JPG and MP3 specific totals
        for i in range(len(data_cleaning_df)):
            for j in range(len(data_cleaning_df.columns)):
                cell_value = str(data_cleaning_df.iloc[i, j])
                if 'JPG' in cell_value and 'Clean Research Data' in str(data_cleaning_df.iloc[i, j+3] if j+3 < len(data_cleaning_df.columns) else ''):
                    try:
                        jpg_total = data_cleaning_df.iloc[i, j+3]
                        if pd.notna(jpg_total) and str(jpg_total).replace(',', '').isdigit():
                            clean_totals['jpg_files'] = int(str(jpg_total).replace(',', ''))
                    except:
                        pass
                elif 'MP3' in cell_value and 'Clean Research Data' in str(data_cleaning_df.iloc[i, j+3] if j+3 < len(data_cleaning_df.columns) else ''):
                    try:
                        mp3_total = data_cleaning_df.iloc[i, j+3]
                        if pd.notna(mp3_total) and str(mp3_total).replace(',', '').isdigit():
                            clean_totals['mp3_files'] = int(str(mp3_total).replace(',', ''))
                    except:
                        pass
        
        print("✅ Data Cleaning sheet totals found:")
        for key, value in clean_totals.items():
            print(f"   {key}: {value:,}")
        
        # Read Daily Counts (ACF_PACF) sheet
        print("\n📋 Reading Daily Counts (ACF_PACF) sheet...")
        daily_counts_df = pd.read_excel(latest_report, sheet_name='Daily Counts (ACF_PACF)', header=2)
        
        # Calculate totals from Daily Counts (ACF_PACF)
        daily_totals = {}
        if 'Total_Files' in daily_counts_df.columns:
            # Convert to numeric, errors='coerce' will turn non-numeric to NaN
            daily_totals['total_files'] = pd.to_numeric(daily_counts_df['Total_Files'], errors='coerce').sum()
        if 'JPG_Files' in daily_counts_df.columns:
            daily_totals['jpg_files'] = pd.to_numeric(daily_counts_df['JPG_Files'], errors='coerce').sum()
        if 'MP3_Files' in daily_counts_df.columns:
            daily_totals['mp3_files'] = pd.to_numeric(daily_counts_df['MP3_Files'], errors='coerce').sum()
        
        print("✅ Daily Counts (ACF_PACF) sheet totals calculated:")
        for key, value in daily_totals.items():
            print(f"   {key}: {value:,}")
        
        # Compare totals
        print("\n🔍 COMPARISON RESULTS:")
        print("=" * 40)
        
        all_match = True
        for key in clean_totals:
            if key in daily_totals:
                clean_val = clean_totals[key]
                daily_val = daily_totals[key]
                match = clean_val == daily_val
                status = "✅ MATCH" if match else "❌ MISMATCH"
                print(f"{key.upper()}: {status}")
                print(f"   Data Cleaning: {clean_val:,}")
                print(f"   Daily Counts:  {daily_val:,}")
                if not match:
                    diff = daily_val - clean_val
                    print(f"   Difference:    {diff:+,}")
                    all_match = False
                print()
        
        # Overall result
        print("=" * 40)
        if all_match:
            print("🎉 SUCCESS: All totals match! The filtering logic fix worked correctly.")
            print("✅ Daily Counts (ACF_PACF) now uses the complete data cleaning logic.")
        else:
            print("⚠️  ISSUE: Some totals don't match. Further investigation needed.")
            print("❌ The filtering logic may need additional adjustments.")
        
        return all_match
        
    except Exception as e:
        print(f"❌ Error during verification: {e}")
        return False

if __name__ == "__main__":
    print("🔧 Daily Counts (ACF_PACF) Data Cleaning Logic Verification")
    print("=" * 60)
    print(f"⏰ Verification started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    success = verify_daily_counts_fix()
    
    print("\n" + "=" * 60)
    if success:
        print("✅ VERIFICATION COMPLETE: Fix successful!")
    else:
        print("❌ VERIFICATION COMPLETE: Issues found.")
