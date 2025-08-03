#!/usr/bin/env python3
"""
Final Totals Verification: Data Cleaning vs Daily Counts (ACF_PACF)
==================================================================

Based on the sheet structure exploration, this script extracts the correct
totals from both sheets and compares them.
"""

import pandas as pd
import os

def extract_data_cleaning_totals():
    """Extract clean dataset totals from Data Cleaning sheet using known structure."""
    
    # Use the specific latest report we know exists
    latest_report = "AR_Analysis_Report_20250802_131741.xlsx"
    
    try:
        df = pd.read_excel(latest_report, sheet_name='Data Cleaning', header=None)
        
        print("🔍 Extracting Data Cleaning totals using known structure...")
        
        totals = {}
        
        # Based on exploration, JPG is at row 5, MP3 at row 6, TOTAL at row 7
        # Clean dataset values appear to be in column 4 or 6
        
        # Extract JPG clean total (row 5)
        for col in [4, 6]:  # Try both potential columns
            if col < len(df.columns):
                jpg_value = df.iloc[5, col]
                if pd.notna(jpg_value) and str(jpg_value).replace(',', '').isdigit():
                    totals['jpg_files'] = int(str(jpg_value).replace(',', ''))
                    print(f"   JPG clean total: {totals['jpg_files']:,} (from row 5, col {col})")
                    break
        
        # Extract MP3 clean total (row 6)
        for col in [4, 6]:  # Try both potential columns
            if col < len(df.columns):
                mp3_value = df.iloc[6, col]
                if pd.notna(mp3_value) and str(mp3_value).replace(',', '').isdigit():
                    totals['mp3_files'] = int(str(mp3_value).replace(',', ''))
                    print(f"   MP3 clean total: {totals['mp3_files']:,} (from row 6, col {col})")
                    break
        
        # Extract overall total (row 7)
        for col in [4, 6]:  # Try both potential columns
            if col < len(df.columns):
                total_value = df.iloc[7, col]
                if pd.notna(total_value) and str(total_value).replace(',', '').isdigit():
                    totals['total_files'] = int(str(total_value).replace(',', ''))
                    print(f"   Total clean files: {totals['total_files']:,} (from row 7, col {col})")
                    break
        
        # Verify that JPG + MP3 = Total (if we have all values)
        if 'jpg_files' in totals and 'mp3_files' in totals and 'total_files' in totals:
            calculated_total = totals['jpg_files'] + totals['mp3_files']
            if calculated_total == totals['total_files']:
                print(f"   ✅ Verification: JPG + MP3 = Total ({calculated_total:,})")
            else:
                print(f"   ⚠️  Note: JPG + MP3 = {calculated_total:,}, but Total = {totals['total_files']:,}")
        
        return totals
        
    except Exception as e:
        print(f"❌ Error extracting Data Cleaning totals: {e}")
        return {}

def extract_daily_counts_totals():
    """Extract totals from Daily Counts (ACF_PACF) sheet."""
    
    filename = "AR_Analysis_Report_20250802_131741.xlsx"
    
    try:
        df = pd.read_excel(filename, sheet_name='Daily Counts (ACF_PACF)', header=2)
        
        print("🔍 Calculating Daily Counts (ACF_PACF) totals...")
        
        totals = {}
        
        # Sum the columns, converting to numeric first
        if 'Total_Files' in df.columns:
            total_files = pd.to_numeric(df['Total_Files'], errors='coerce').sum()
            totals['total_files'] = int(total_files)
            print(f"   Total Files: {totals['total_files']:,}")
        
        if 'JPG_Files' in df.columns:
            jpg_files = pd.to_numeric(df['JPG_Files'], errors='coerce').sum()
            totals['jpg_files'] = int(jpg_files)
            print(f"   JPG Files: {totals['jpg_files']:,}")
        
        if 'MP3_Files' in df.columns:
            mp3_files = pd.to_numeric(df['MP3_Files'], errors='coerce').sum()
            totals['mp3_files'] = int(mp3_files)
            print(f"   MP3 Files: {totals['mp3_files']:,}")
        
        # Verify that JPG + MP3 = Total
        if 'jpg_files' in totals and 'mp3_files' in totals and 'total_files' in totals:
            calculated_total = totals['jpg_files'] + totals['mp3_files']
            if calculated_total == totals['total_files']:
                print(f"   ✅ Verification: JPG + MP3 = Total ({calculated_total:,})")
            else:
                print(f"   ⚠️  Note: JPG + MP3 = {calculated_total:,}, but Total = {totals['total_files']:,}")
        
        return totals
        
    except Exception as e:
        print(f"❌ Error extracting Daily Counts totals: {e}")
        return {}

def main():
    print("🔧 FINAL VERIFICATION: Data Cleaning vs Daily Counts (ACF_PACF)")
    print("=" * 70)
    print("📊 Analyzing: AR_Analysis_Report_20250802_131741.xlsx")
    print()
    
    # Extract totals from both sheets
    print("📋 EXTRACTING DATA CLEANING TOTALS:")
    data_cleaning_totals = extract_data_cleaning_totals()
    
    print("\n📋 EXTRACTING DAILY COUNTS (ACF_PACF) TOTALS:")
    daily_counts_totals = extract_daily_counts_totals()
    
    # Compare results
    print("\n🔍 FINAL COMPARISON RESULTS:")
    print("=" * 50)
    
    if not data_cleaning_totals:
        print("❌ Could not extract Data Cleaning totals")
        return
    
    if not daily_counts_totals:
        print("❌ Could not extract Daily Counts totals")
        return
    
    all_match = True
    for key in ['total_files', 'jpg_files', 'mp3_files']:
        if key in data_cleaning_totals and key in daily_counts_totals:
            clean_val = data_cleaning_totals[key]
            daily_val = daily_counts_totals[key]
            match = clean_val == daily_val
            
            status = "✅ MATCH" if match else "❌ MISMATCH"
            print(f"\n{key.upper().replace('_', ' ')}: {status}")
            print(f"   Data Cleaning:    {clean_val:,}")
            print(f"   Daily Counts:     {daily_val:,}")
            
            if not match:
                diff = daily_val - clean_val
                percentage_diff = (diff / clean_val) * 100 if clean_val > 0 else 0
                print(f"   Difference:       {diff:+,} ({percentage_diff:+.2f}%)")
                all_match = False
    
    print("\n" + "=" * 50)
    if all_match:
        print("🎉 SUCCESS: All totals match perfectly!")
        print("✅ The filtering logic fix worked correctly.")
        print("✅ Daily Counts (ACF_PACF) now uses complete data cleaning logic.")
        print("✅ Both sheets apply: School_Year ≠ 'N/A', file_type ∈ ['JPG', 'MP3'],")
        print("   is_collection_day: TRUE, Outlier_Status: FALSE")
    else:
        print("⚠️  ISSUE: Totals don't match.")
        print("❌ The filtering logic may need further adjustment.")
        print("🔍 This suggests there may be additional filtering differences.")

if __name__ == "__main__":
    main()
