#!/usr/bin/env python3
"""
Corrected Verification Script

This script properly extracts totals from both sheets to verify the data cleaning logic.
"""

import pandas as pd
import numpy as np

def extract_data_cleaning_totals():
    """Extract clean dataset totals from Data Cleaning sheet."""
    filename = "AR_Analysis_Report_20250802_131741.xlsx"
    
    try:
        # Read the Data Cleaning sheet
        df = pd.read_excel(filename, sheet_name='Data Cleaning', header=None)
        print(f"📋 Data Cleaning sheet dimensions: {df.shape}")
        
        # Print first 10 rows to understand structure
        print("\n🔍 First 10 rows of Data Cleaning sheet:")
        for i in range(min(10, len(df))):
            print(f"   Row {i}: {df.iloc[i, 0]} | {df.iloc[i, 1] if len(df.columns) > 1 else 'N/A'} | {df.iloc[i, 2] if len(df.columns) > 2 else 'N/A'} | {df.iloc[i, 3] if len(df.columns) > 3 else 'N/A'}")
        
        # Look for the intersection analysis table
        # Find rows containing "JPG", "MP3", and "Total"
        jpg_row = None
        mp3_row = None
        total_row = None
        
        for i in range(len(df)):
            cell_value = str(df.iloc[i, 0]).strip().upper()
            if 'JPG' in cell_value:
                jpg_row = i
                print(f"   Found JPG row at index {i}")
            elif 'MP3' in cell_value:
                mp3_row = i
                print(f"   Found MP3 row at index {i}")
            elif cell_value == 'TOTAL':
                total_row = i
                print(f"   Found TOTAL row at index {i}")
        
        if jpg_row is not None and mp3_row is not None and total_row is not None:
            # Extract clean dataset values (should be in the "Clean Research Data" column)
            # Look for the column with clean data - usually column 4 or 5
            for col in range(1, min(8, len(df.columns))):
                header = str(df.iloc[0, col]).strip() if pd.notna(df.iloc[0, col]) else ""
                print(f"   Column {col} header: '{header}'")
                
                if 'CLEAN' in header.upper() or 'RESEARCH' in header.upper() or 'FINAL' in header.upper():
                    jpg_clean = df.iloc[jpg_row, col]
                    mp3_clean = df.iloc[mp3_row, col]
                    total_clean = df.iloc[total_row, col]
                    
                    print(f"   Using column {col} for clean data:")
                    print(f"   JPG clean: {jpg_clean}")
                    print(f"   MP3 clean: {mp3_clean}")
                    print(f"   Total clean: {total_clean}")
                    
                    return {
                        'jpg': jpg_clean,
                        'mp3': mp3_clean,
                        'total': total_clean
                    }
        
        print("❌ Could not find clean dataset values in Data Cleaning sheet")
        return None
        
    except Exception as e:
        print(f"❌ Error reading Data Cleaning sheet: {e}")
        return None

def extract_daily_counts_totals():
    """Extract totals from Daily Counts (ACF_PACF) sheet."""
    filename = "AR_Analysis_Report_20250802_131741.xlsx"
    
    try:
        # Read the Daily Counts sheet
        df = pd.read_excel(filename, sheet_name='Daily Counts (ACF_PACF)', header=2)
        print(f"📋 Daily Counts sheet dimensions: {df.shape}")
        print(f"📋 Daily Counts columns: {list(df.columns)}")
        
        # Convert numeric columns to proper numeric types
        numeric_columns = ['Total_Files', 'JPG_Files', 'MP3_Files']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Look for totals row first
        totals_found = False
        for i in range(len(df)):
            if pd.notna(df.iloc[i, 0]):
                cell_value = str(df.iloc[i, 0]).strip().upper()
                if 'TOTAL' in cell_value:
                    print(f"   Found totals row at index {i}")
                    total_files = df.iloc[i, df.columns.get_loc('Total_Files')] if 'Total_Files' in df.columns else 0
                    jpg_files = df.iloc[i, df.columns.get_loc('JPG_Files')] if 'JPG_Files' in df.columns else 0
                    mp3_files = df.iloc[i, df.columns.get_loc('MP3_Files')] if 'MP3_Files' in df.columns else 0
                    
                    print(f"   Total Files: {total_files}")
                    print(f"   JPG Files: {jpg_files}")
                    print(f"   MP3 Files: {mp3_files}")
                    
                    totals_found = True
                    return {
                        'total': total_files,
                        'jpg': jpg_files,
                        'mp3': mp3_files
                    }
        
        if not totals_found:
            print("   No totals row found, summing data rows...")
            
            # Filter out non-data rows and sum
            data_df = df.dropna(subset=['Total_Files'])
            
            total_files = data_df['Total_Files'].sum() if 'Total_Files' in data_df.columns else 0
            jpg_files = data_df['JPG_Files'].sum() if 'JPG_Files' in data_df.columns else 0
            mp3_files = data_df['MP3_Files'].sum() if 'MP3_Files' in data_df.columns else 0
            
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
        import traceback
        traceback.print_exc()
        return None

def main():
    print("🔧 CORRECTED VERIFICATION SCRIPT")
    print("=" * 50)
    
    print("\n📋 EXTRACTING DATA CLEANING TOTALS:")
    data_cleaning_totals = extract_data_cleaning_totals()
    
    print("\n📋 EXTRACTING DAILY COUNTS TOTALS:")
    daily_counts_totals = extract_daily_counts_totals()
    
    if not data_cleaning_totals or not daily_counts_totals:
        print("❌ Failed to extract totals from one or both sheets.")
        return
    
    print("\n🔍 COMPARISON RESULTS:")
    print("=" * 50)
    
    # Compare totals
    dc_total = data_cleaning_totals['total']
    daily_total = daily_counts_totals['total']
    
    print(f"TOTAL FILES:")
    print(f"   Data Cleaning:    {dc_total}")
    print(f"   Daily Counts:     {daily_total}")
    
    if dc_total == daily_total:
        print("   ✅ MATCH!")
    else:
        diff = daily_total - dc_total
        print(f"   ❌ MISMATCH: {diff:+} files")
    
    print(f"\nUser reported Daily Counts total: 9,372")
    print(f"Data Cleaning total: {dc_total}")
    print(f"Difference: {9372 - dc_total} files")

if __name__ == "__main__":
    main()
