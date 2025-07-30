#!/usr/bin/env python3
"""Simple column checker for Period Counts sheet"""

import pandas as pd

try:
    # Read the latest report
    df = pd.read_excel("AR_Analysis_Report_20250726_213437.xlsx", sheet_name="Period Counts (ACF_PACF)")
    
    print("COLUMN ANALYSIS:")
    print(f"Total columns: {len(df.columns)}")
    print(f"Unique columns: {len(set(df.columns))}")
    
    print("\nCOLUMN LIST:")
    for i, col in enumerate(df.columns, 1):
        print(f"{i:2d}. {col}")
    
    # Check for duplicates
    duplicates = []
    seen = set()
    for col in df.columns:
        if col in seen:
            duplicates.append(col)
        else:
            seen.add(col)
    
    if duplicates:
        print(f"\nDUPLICATE COLUMNS: {duplicates}")
        print("❌ ISSUE STILL EXISTS")
    else:
        print("\n✅ NO DUPLICATE COLUMNS FOUND - ISSUE RESOLVED!")
        
except Exception as e:
    print(f"Error: {e}")
