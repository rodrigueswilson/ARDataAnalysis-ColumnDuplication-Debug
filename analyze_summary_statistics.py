#!/usr/bin/env python3
"""
Analyze the Summary Statistics sheet to check for inconsistencies
in the "Total Collection Days" column.
"""

import pandas as pd
import glob
from ar_utils import calculate_collection_days_for_period

def find_latest_report():
    """Find the most recent AR Analysis Report Excel file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort()
    return files[-1]

def analyze_summary_statistics():
    """Analyze the Summary Statistics sheet for Total Collection Days inconsistencies."""
    latest_report = find_latest_report()
    if not latest_report:
        print("No report files found!")
        return
    
    print(f"Analyzing Summary Statistics in: {latest_report}")
    print("\n=== SUMMARY STATISTICS ANALYSIS ===")
    
    try:
        # Read the Summary Statistics sheet
        df = pd.read_excel(latest_report, sheet_name='Summary Statistics')
        
        print(f"Sheet dimensions: {len(df)} rows x {len(df.columns)} columns")
        print(f"Columns: {list(df.columns)}")
        print()
        
        # Look for Total Collection Days column
        collection_days_col = None
        for col in df.columns:
            if 'collection' in str(col).lower() and 'days' in str(col).lower():
                collection_days_col = col
                break
        
        if not collection_days_col:
            print("❌ No 'Total Collection Days' column found!")
            print("Available columns:")
            for i, col in enumerate(df.columns, 1):
                print(f"  {i:2d}. {col}")
            return
        
        print(f"Found collection days column: '{collection_days_col}'")
        print()
        
        # Display the data
        print("=== SUMMARY STATISTICS DATA ===")
        relevant_cols = []
        
        # Try to identify relevant columns
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['period', 'year', 'school', collection_days_col.lower()]):
                relevant_cols.append(col)
        
        if relevant_cols:
            print("Relevant columns found:")
            display_df = df[relevant_cols].copy()
            
            # Clean up display
            for col in display_df.columns:
                if display_df[col].dtype == 'object':
                    display_df[col] = display_df[col].astype(str)
            
            print(display_df.to_string(index=False))
        else:
            print("Showing all data:")
            print(df.to_string(index=False))
        
        print()
        
        # Analyze the Total Collection Days values
        print("=== COLLECTION DAYS ANALYSIS ===")
        collection_days_values = df[collection_days_col].dropna()
        
        print(f"Total Collection Days values:")
        print(f"  Count: {len(collection_days_values)}")
        print(f"  Min: {collection_days_values.min()}")
        print(f"  Max: {collection_days_values.max()}")
        print(f"  Mean: {collection_days_values.mean():.1f}")
        print(f"  Sum: {collection_days_values.sum()}")
        print()
        
        # Check for expected values if we can identify periods
        if 'Period' in df.columns or any('period' in str(col).lower() for col in df.columns):
            print("=== CONSISTENCY CHECK ===")
            
            # Try to find period column
            period_col = None
            for col in df.columns:
                if 'period' in str(col).lower():
                    period_col = col
                    break
            
            if period_col:
                print(f"Using period column: '{period_col}'")
                
                inconsistencies = []
                for idx, row in df.iterrows():
                    period = str(row[period_col])
                    reported_days = row[collection_days_col]
                    
                    if pd.isna(reported_days):
                        continue
                    
                    # Calculate expected collection days
                    expected_days = calculate_collection_days_for_period(period)
                    
                    if expected_days > 0:  # Valid period found
                        if abs(reported_days - expected_days) > 0.1:  # Allow for small floating point differences
                            inconsistencies.append({
                                'period': period,
                                'reported': reported_days,
                                'expected': expected_days,
                                'difference': reported_days - expected_days
                            })
                        
                        print(f"  {period}: Reported={reported_days}, Expected={expected_days}, Diff={reported_days - expected_days}")
                    else:
                        print(f"  {period}: Reported={reported_days}, Expected=Unknown (period not recognized)")
                
                if inconsistencies:
                    print(f"\n⚠️  Found {len(inconsistencies)} inconsistencies:")
                    for issue in inconsistencies:
                        print(f"    {issue['period']}: Off by {issue['difference']} days")
                else:
                    print(f"\n✅ All collection days values are consistent!")
            else:
                print("No period column found for consistency check")
        
        # Check for school year totals
        if any('year' in str(col).lower() for col in df.columns):
            print(f"\n=== SCHOOL YEAR ANALYSIS ===")
            year_col = None
            for col in df.columns:
                if 'year' in str(col).lower():
                    year_col = col
                    break
            
            if year_col:
                year_totals = df.groupby(year_col)[collection_days_col].sum()
                print("Collection days by school year:")
                for year, total in year_totals.items():
                    print(f"  {year}: {total} days")
                
                # Expected totals for known years
                expected_totals = {
                    '2021-2022': 34 + 59 + 63,  # P1 + P2 + P3
                    '2022-2023': 56 + 58 + 50   # P1 + P2 + P3
                }
                
                print("\nComparison with expected totals:")
                for year, reported_total in year_totals.items():
                    year_str = str(year)
                    if year_str in expected_totals:
                        expected = expected_totals[year_str]
                        diff = reported_total - expected
                        status = "✅" if abs(diff) < 0.1 else "⚠️"
                        print(f"  {year}: {status} Reported={reported_total}, Expected={expected}, Diff={diff}")
                    else:
                        print(f"  {year}: Reported={reported_total}, Expected=Unknown")
    
    except Exception as e:
        print(f"Error analyzing Summary Statistics: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_summary_statistics()
