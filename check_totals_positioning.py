#!/usr/bin/env python3
"""
Quick script to check totals positioning in MP3 Duration Analysis sheet
"""

import pandas as pd

def check_totals_positioning():
    try:
        # Read the Excel file
        df = pd.read_excel('test_mp3_fixes_simple.xlsx', sheet_name='MP3 Duration Analysis', header=None)
        
        print("=== MP3 Duration Analysis Totals Positioning Check ===")
        print(f"Total sheet dimensions: {df.shape[0]} rows x {df.shape[1]} columns")
        print()
        
        # Find totals rows
        totals_rows = []
        for i in range(len(df)):
            cell_value = str(df.iloc[i, 0]) if pd.notna(df.iloc[i, 0]) else ""
            if "Total" in cell_value:
                totals_rows.append((i+1, cell_value))  # +1 for 1-based row numbering
        
        print("=== TOTALS ROWS FOUND ===")
        for row_num, content in totals_rows:
            print(f"Row {row_num}: {content}")
        print()
        
        # Check Period Breakdown table specifically
        print("=== PERIOD BREAKDOWN TABLE ===")
        period_start = None
        for i in range(len(df)):
            if str(df.iloc[i, 0]) == "Collection Period MP3 Duration Analysis":
                period_start = i
                break
        
        if period_start:
            print(f"Period table starts at row {period_start + 1}")
            # Show period table content
            for i in range(period_start + 2, min(period_start + 10, len(df))):
                row_content = [str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else "" for j in range(3)]
                if any(row_content):  # Only show non-empty rows
                    print(f"Row {i+1}: {row_content}")
        
        print()
        
        # Check Monthly Distribution table specifically  
        print("=== MONTHLY DISTRIBUTION TABLE ===")
        monthly_start = None
        for i in range(len(df)):
            if str(df.iloc[i, 0]) == "Monthly MP3 Duration Distribution":
                monthly_start = i
                break
        
        if monthly_start:
            print(f"Monthly table starts at row {monthly_start + 1}")
            # Show end of monthly table content
            for i in range(monthly_start + 10, min(monthly_start + 20, len(df))):
                row_content = [str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else "" for j in range(3)]
                if any(row_content):  # Only show non-empty rows
                    print(f"Row {i+1}: {row_content}")
        
    except Exception as e:
        print(f"Error checking totals positioning: {e}")

if __name__ == "__main__":
    check_totals_positioning()
