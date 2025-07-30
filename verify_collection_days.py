#!/usr/bin/env python3
"""
Verification script to check if the Collection Days column was correctly added
to the Period Counts (ACF_PACF) sheet in the latest generated report.
"""

import pandas as pd
import glob
import os
from pathlib import Path

def find_latest_report():
    """Find the most recent AR Analysis Report Excel file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        print("No report files found!")
        return None
    
    # Sort by modification time, get the latest
    latest_file = max(files, key=os.path.getmtime)
    print(f"Latest report: {latest_file}")
    return latest_file

def verify_collection_days_column():
    """Verify that the Collection Days column exists and is correctly positioned."""
    latest_report = find_latest_report()
    if not latest_report:
        return
    
    try:
        # Read the Period Counts (ACF_PACF) sheet
        df = pd.read_excel(latest_report, sheet_name='Period Counts (ACF_PACF)')
        
        print(f"\n=== Period Counts (ACF_PACF) Sheet Analysis ===")
        print(f"Total rows: {len(df)}")
        print(f"Total columns: {len(df.columns)}")
        
        # Check if Collection_Days column exists
        if 'Collection_Days' in df.columns:
            print(f"✓ Collection_Days column found!")
            
            # Check column position
            columns_list = list(df.columns)
            collection_days_index = columns_list.index('Collection_Days')
            
            print(f"Collection_Days column position: {collection_days_index + 1}")
            
            # Check if it's positioned correctly (between Total_Size_MB and Days_With_Data)
            if 'Total_Size_MB' in columns_list and 'Days_With_Data' in columns_list:
                total_size_index = columns_list.index('Total_Size_MB')
                days_with_data_index = columns_list.index('Days_With_Data')
                
                print(f"Total_Size_MB position: {total_size_index + 1}")
                print(f"Days_With_Data position: {days_with_data_index + 1}")
                
                if total_size_index < collection_days_index < days_with_data_index:
                    print("✓ Collection_Days is correctly positioned between Total_Size_MB and Days_With_Data")
                else:
                    print("✗ Collection_Days is NOT correctly positioned")
            
            # Show sample data
            print(f"\n=== Sample Collection Days Data ===")
            relevant_cols = ['Period', 'Total_Size_MB', 'Collection_Days', 'Days_With_Data']
            available_cols = [col for col in relevant_cols if col in df.columns]
            print(df[available_cols].head())
            
            # Show statistics
            print(f"\n=== Collection Days Statistics ===")
            print(f"Min Collection Days: {df['Collection_Days'].min()}")
            print(f"Max Collection Days: {df['Collection_Days'].max()}")
            print(f"Mean Collection Days: {df['Collection_Days'].mean():.1f}")
            
        else:
            print("✗ Collection_Days column NOT found!")
            print("Available columns:")
            for i, col in enumerate(df.columns, 1):
                print(f"  {i:2d}. {col}")
        
    except Exception as e:
        print(f"Error reading the Excel file: {e}")

if __name__ == "__main__":
    verify_collection_days_column()
