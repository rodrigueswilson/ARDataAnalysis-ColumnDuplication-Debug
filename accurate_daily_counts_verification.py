#!/usr/bin/env python3
"""
Accurate Daily Counts Verification Script

This script properly reads the Daily Counts (ACF_PACF) sheet to get accurate totals
and compares them with the Data Cleaning sheet totals.
"""

import pandas as pd
import os

def get_latest_report():
    """Find the latest AR Analysis report file."""
    files = [f for f in os.listdir('.') if f.startswith('AR_Analysis_Report_') and f.endswith('.xlsx')]
    if not files:
        return None
    return max(files)

def extract_data_cleaning_totals(filename):
    """Extract clean dataset totals from Data Cleaning sheet."""
    try:
        df = pd.read_excel(filename, sheet_name='Data Cleaning', header=None)
        print(f"📋 Data Cleaning sheet dimensions: {df.shape}")
        
        # Based on known structure - clean dataset totals are in column 4 (index 3)
        jpg_clean = df.iloc[4, 3]  # Row 5, Column 4 (0-indexed: 4,3)
        mp3_clean = df.iloc[5, 3]  # Row 6, Column 4 (0-indexed: 5,3)
        total_clean = df.iloc[6, 3]  # Row 7, Column 4 (0-indexed: 6,3)
        
        print(f"   JPG clean files: {jpg_clean}")
        print(f"   MP3 clean files: {mp3_clean}")
        print(f"   Total clean files: {total_clean}")
        
        return {
            'jpg': jpg_clean,
            'mp3': mp3_clean,
            'total': total_clean
        }
        
    except Exception as e:
        print(f"❌ Error reading Data Cleaning sheet: {e}")
        return None

def extract_daily_counts_totals(filename):
    """Extract totals from Daily Counts (ACF_PACF) sheet more accurately."""
    try:
        # Read the sheet with proper header
        df = pd.read_excel(filename, sheet_name='Daily Counts (ACF_PACF)', header=2)
        print(f"📋 Daily Counts sheet dimensions: {df.shape}")
        print(f"📋 Daily Counts columns: {list(df.columns)}")
        
        # Look for the totals row (usually at the bottom)
        print("\n🔍 Looking for totals row...")
        
        # Check the last few rows for totals
        for i in range(len(df)-1, max(len(df)-10, -1), -1):
            row_data = df.iloc[i]
            print(f"   Row {i}: {row_data.iloc[0] if len(row_data) > 0 else 'N/A'}")
            
            # Look for a row that might contain totals
            if pd.notna(row_data.iloc[0]) and str(row_data.iloc[0]).lower().strip() in ['total', 'totals']:
                print(f"   ✅ Found totals row at index {i}")
                
                # Extract totals from this row
                total_files = row_data.get('Total_Files', 0)
                jpg_files = row_data.get('JPG_Files', 0)
                mp3_files = row_data.get('MP3_Files', 0)
                
                print(f"   Total Files: {total_files}")
                print(f"   JPG Files: {jpg_files}")
                print(f"   MP3 Files: {mp3_files}")
                
                return {
                    'total': total_files,
                    'jpg': jpg_files,
                    'mp3': mp3_files
                }
        
        # If no totals row found, sum all the data rows
        print("   ⚠️  No explicit totals row found, summing data rows...")
        
        # Sum numeric columns
        total_files = df['Total_Files'].sum() if 'Total_Files' in df.columns else 0
        jpg_files = df['JPG_Files'].sum() if 'JPG_Files' in df.columns else 0
        mp3_files = df['MP3_Files'].sum() if 'MP3_Files' in df.columns else 0
        
        print(f"   Calculated Total Files: {total_files}")
        print(f"   Calculated JPG Files: {jpg_files}")
        print(f"   Calculated MP3 Files: {mp3_files}")
        
        return {
            'total': total_files,
            'jpg': jpg_files,
            'mp3': mp3_files
        }
        
    except Exception as e:
        print(f"❌ Error reading Daily Counts sheet: {e}")
        return None

def main():
    print("🔧 ACCURATE DAILY COUNTS VERIFICATION")
    print("=" * 60)
    
    # Use the latest report
    filename = "AR_Analysis_Report_20250802_131741.xlsx"
    if not os.path.exists(filename):
        filename = get_latest_report()
        if not filename:
            print("❌ No report file found!")
            return
    
    print(f"📊 Analyzing: {filename}")
    print()
    
    # Extract totals from both sheets
    print("📋 EXTRACTING DATA CLEANING TOTALS:")
    data_cleaning_totals = extract_data_cleaning_totals(filename)
    
    print("\n📋 EXTRACTING DAILY COUNTS TOTALS:")
    daily_counts_totals = extract_daily_counts_totals(filename)
    
    if not data_cleaning_totals or not daily_counts_totals:
        print("❌ Failed to extract totals from one or both sheets.")
        return
    
    print("\n🔍 COMPARISON RESULTS:")
    print("=" * 50)
    
    # Compare totals
    dc_total = data_cleaning_totals['total']
    daily_total = daily_counts_totals['total']
    
    print(f"TOTAL FILES:")
    print(f"   Data Cleaning:    {dc_total:,}")
    print(f"   Daily Counts:     {daily_total:,}")
    
    if dc_total == daily_total:
        print("   ✅ MATCH!")
    else:
        diff = daily_total - dc_total
        pct_diff = (diff / dc_total * 100) if dc_total != 0 else 0
        print(f"   ❌ MISMATCH: {diff:+,} ({pct_diff:+.2f}%)")
    
    print()
    
    # Compare JPG files
    dc_jpg = data_cleaning_totals['jpg']
    daily_jpg = daily_counts_totals['jpg']
    
    print(f"JPG FILES:")
    print(f"   Data Cleaning:    {dc_jpg:,}")
    print(f"   Daily Counts:     {daily_jpg:,}")
    
    if dc_jpg == daily_jpg:
        print("   ✅ MATCH!")
    else:
        diff = daily_jpg - dc_jpg
        pct_diff = (diff / dc_jpg * 100) if dc_jpg != 0 else 0
        print(f"   ❌ MISMATCH: {diff:+,} ({pct_diff:+.2f}%)")
    
    print()
    
    # Compare MP3 files
    dc_mp3 = data_cleaning_totals['mp3']
    daily_mp3 = daily_counts_totals['mp3']
    
    print(f"MP3 FILES:")
    print(f"   Data Cleaning:    {dc_mp3:,}")
    print(f"   Daily Counts:     {daily_mp3:,}")
    
    if dc_mp3 == daily_mp3:
        print("   ✅ MATCH!")
    else:
        diff = daily_mp3 - dc_mp3
        pct_diff = (diff / dc_mp3 * 100) if dc_mp3 != 0 else 0
        print(f"   ❌ MISMATCH: {diff:+,} ({pct_diff:+.2f}%)")
    
    print("\n" + "=" * 50)
    
    # Overall assessment
    total_match = dc_total == daily_total
    jpg_match = dc_jpg == daily_jpg
    mp3_match = dc_mp3 == daily_mp3
    
    if total_match and jpg_match and mp3_match:
        print("🎉 SUCCESS: All totals match perfectly!")
        print("✅ Daily Counts (ACF_PACF) uses complete data cleaning logic.")
    else:
        print("⚠️  ISSUE: Some totals don't match.")
        print("❌ The filtering logic may need further adjustment.")

if __name__ == "__main__":
    main()
