#!/usr/bin/env python3
"""
Data Cleaning Sheet Explorer
===========================

This script explores the structure of the Data Cleaning sheet to understand
how to extract the clean dataset totals correctly.
"""

import pandas as pd
import os

def explore_data_cleaning_sheet():
    """Explore the Data Cleaning sheet structure."""
    
    filename = "AR_Analysis_Report_20250802_130914.xlsx"
    if not os.path.exists(filename):
        print(f"❌ Report file not found: {filename}")
        return
    
    try:
        # Read the entire Data Cleaning sheet without headers
        df = pd.read_excel(filename, sheet_name='Data Cleaning', header=None)
        
        print("🔍 Data Cleaning Sheet Structure:")
        print("=" * 50)
        print(f"Sheet dimensions: {len(df)} rows × {len(df.columns)} columns")
        print()
        
        # Display the first 20 rows to understand structure
        print("📋 First 20 rows of Data Cleaning sheet:")
        for i in range(min(20, len(df))):
            row_data = []
            for j in range(min(8, len(df.columns))):  # Show first 8 columns
                cell_value = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                if len(cell_value) > 20:
                    cell_value = cell_value[:17] + "..."
                row_data.append(f"{cell_value:20}")
            print(f"Row {i:2d}: {' | '.join(row_data)}")
        
        print("\n🔍 Looking for key terms...")
        
        # Search for key terms
        key_terms = ['Clean Research Data', 'Final Dataset', 'JPG', 'MP3', 'Total']
        for term in key_terms:
            print(f"\n🔎 Searching for '{term}':")
            found_locations = []
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    cell_value = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                    if term.lower() in cell_value.lower():
                        found_locations.append((i, j, cell_value))
            
            if found_locations:
                for row, col, value in found_locations[:5]:  # Show first 5 matches
                    print(f"   Row {row}, Col {col}: {value}")
            else:
                print(f"   No matches found for '{term}'")
        
        # Look for numeric values that could be totals
        print(f"\n🔢 Looking for large numeric values (potential totals):")
        large_numbers = []
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell_value = df.iloc[i, j]
                if pd.notna(cell_value):
                    # Try to convert to number
                    try:
                        if isinstance(cell_value, str):
                            # Remove commas and try to convert
                            num_str = cell_value.replace(',', '')
                            if num_str.isdigit():
                                num_value = int(num_str)
                                if num_value > 1000:  # Potential total
                                    large_numbers.append((i, j, num_value, cell_value))
                        elif isinstance(cell_value, (int, float)) and cell_value > 1000:
                            large_numbers.append((i, j, int(cell_value), str(cell_value)))
                    except:
                        pass
        
        # Sort by value and show largest numbers
        large_numbers.sort(key=lambda x: x[2], reverse=True)
        for row, col, num_value, original_value in large_numbers[:10]:
            print(f"   Row {row}, Col {col}: {original_value} ({num_value:,})")
        
    except Exception as e:
        print(f"❌ Error exploring Data Cleaning sheet: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    explore_data_cleaning_sheet()
