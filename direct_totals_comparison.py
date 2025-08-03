#!/usr/bin/env python3
"""
Direct Totals Comparison: Data Cleaning vs Daily Counts (ACF_PACF)
================================================================

This script directly compares totals between the two sheets to verify
that the filtering logic fix worked correctly.
"""

import pandas as pd
import os

def find_latest_report():
    """Find the most recent Excel report."""
    report_files = [f for f in os.listdir('.') if f.startswith('AR_Analysis_Report_') and f.endswith('.xlsx')]
    if not report_files:
        return None
    return max(report_files)

def extract_data_cleaning_totals(filename):
    """Extract clean dataset totals from Data Cleaning sheet."""
    try:
        # Read the entire Data Cleaning sheet
        df = pd.read_excel(filename, sheet_name='Data Cleaning', header=None)
        
        print("🔍 Searching Data Cleaning sheet for clean dataset totals...")
        
        # Look for the intersection analysis table
        totals = {}
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                
                # Look for JPG row in intersection table
                if 'JPG' in cell_value and i < len(df) - 1:
                    # Check if this is the intersection analysis table by looking for "Clean Research Data" column
                    for col_offset in range(1, 8):
                        if j + col_offset < len(df.columns):
                            header_cell = str(df.iloc[i-1, j + col_offset]) if i > 0 else ""
                            if 'Clean Research Data' in header_cell or 'Final Dataset' in header_cell:
                                try:
                                    jpg_clean = df.iloc[i, j + col_offset]
                                    if pd.notna(jpg_clean):
                                        totals['jpg_files'] = int(str(jpg_clean).replace(',', ''))
                                        print(f"   Found JPG clean total: {totals['jpg_files']:,}")
                                except:
                                    pass
                
                # Look for MP3 row in intersection table
                elif 'MP3' in cell_value and i < len(df) - 1:
                    for col_offset in range(1, 8):
                        if j + col_offset < len(df.columns):
                            header_cell = str(df.iloc[i-1, j + col_offset]) if i > 0 else ""
                            if 'Clean Research Data' in header_cell or 'Final Dataset' in header_cell:
                                try:
                                    mp3_clean = df.iloc[i, j + col_offset]
                                    if pd.notna(mp3_clean):
                                        totals['mp3_files'] = int(str(mp3_clean).replace(',', ''))
                                        print(f"   Found MP3 clean total: {totals['mp3_files']:,}")
                                except:
                                    pass
        
        # Calculate total files
        if 'jpg_files' in totals and 'mp3_files' in totals:
            totals['total_files'] = totals['jpg_files'] + totals['mp3_files']
            print(f"   Calculated total files: {totals['total_files']:,}")
        
        return totals
        
    except Exception as e:
        print(f"❌ Error reading Data Cleaning sheet: {e}")
        return {}

def extract_daily_counts_totals(filename):
    """Extract totals from Daily Counts (ACF_PACF) sheet."""
    try:
        # Read Daily Counts sheet with proper headers
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
        
        return totals
        
    except Exception as e:
        print(f"❌ Error reading Daily Counts (ACF_PACF) sheet: {e}")
        return {}

def main():
    print("🔧 Direct Totals Comparison: Data Cleaning vs Daily Counts (ACF_PACF)")
    print("=" * 70)
    
    # Find latest report
    # Use the specific latest report we know exists
    latest_report = "AR_Analysis_Report_20250802_130914.xlsx"
    if not os.path.exists(latest_report):
        print(f"❌ Report file not found: {latest_report}")
        return
    
    print(f"📊 Analyzing: {latest_report}")
    print()
    
    # Extract totals from both sheets
    print("📋 EXTRACTING DATA CLEANING TOTALS:")
    data_cleaning_totals = extract_data_cleaning_totals(latest_report)
    
    print("\n📋 EXTRACTING DAILY COUNTS (ACF_PACF) TOTALS:")
    daily_counts_totals = extract_daily_counts_totals(latest_report)
    
    # Compare results
    print("\n🔍 COMPARISON RESULTS:")
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
                print(f"   Difference:       {diff:+,}")
                all_match = False
    
    print("\n" + "=" * 50)
    if all_match:
        print("🎉 SUCCESS: All totals match perfectly!")
        print("✅ The filtering logic fix worked correctly.")
        print("✅ Daily Counts (ACF_PACF) now uses complete data cleaning logic.")
    else:
        print("⚠️  ISSUE: Totals don't match.")
        print("❌ Further investigation needed.")

if __name__ == "__main__":
    main()
