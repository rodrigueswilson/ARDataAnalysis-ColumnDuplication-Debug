#!/usr/bin/env python3
"""
Focused analysis to check for collection day consistency issues across sheets,
excluding total/sum rows which are expected to have high values.
"""

import pandas as pd
import glob

def find_latest_report():
    """Find the most recent AR Analysis Report Excel file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort()
    return files[-1]

def check_collection_day_consistency():
    """Check for collection day consistency issues across all relevant sheets."""
    latest_report = find_latest_report()
    if not latest_report:
        print("No report files found!")
        return
    
    print(f"Analyzing: {latest_report}")
    print("\n=== COLLECTION DAY CONSISTENCY CHECK ===")
    
    try:
        excel_file = pd.ExcelFile(latest_report)
        
        # Focus on sheets that likely have collection day data
        target_sheets = [
            'Period Counts (ACF_PACF)',
            'Monthly Counts (ACF_PACF)', 
            'Weekly Counts (ACF_PACF)',
            'Daily Counts (ACF_PACF)',
            'Biweekly Counts (ACF_PACF)'
        ]
        
        issues_found = []
        
        for sheet_name in target_sheets:
            if sheet_name in excel_file.sheet_names:
                print(f"\n--- {sheet_name} ---")
                issues = analyze_sheet_collection_days(latest_report, sheet_name)
                if issues:
                    issues_found.extend(issues)
            else:
                print(f"\n--- {sheet_name} --- (Not found)")
        
        # Check other sheets that might have day-related data
        other_relevant_sheets = []
        for sheet_name in excel_file.sheet_names:
            if sheet_name not in target_sheets:
                df = pd.read_excel(latest_report, sheet_name=sheet_name, nrows=5)  # Just peek
                if any('Days' in str(col) or 'Day' in str(col) for col in df.columns):
                    other_relevant_sheets.append(sheet_name)
        
        if other_relevant_sheets:
            print(f"\n=== OTHER SHEETS WITH DAY-RELATED DATA ===")
            for sheet_name in other_relevant_sheets:
                print(f"\n--- {sheet_name} ---")
                issues = analyze_sheet_collection_days(latest_report, sheet_name)
                if issues:
                    issues_found.extend(issues)
        
        # Summary
        print(f"\n=== SUMMARY ===")
        if issues_found:
            print(f"Found {len(issues_found)} potential issues:")
            for issue in issues_found:
                print(f"  ⚠️  {issue}")
        else:
            print("✅ No collection day consistency issues found!")
            print("✅ All sheets appear to be using consistent logic")
    
    except Exception as e:
        print(f"Error: {e}")

def analyze_sheet_collection_days(file_path, sheet_name):
    """Analyze a specific sheet for collection day issues."""
    issues = []
    
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Look for Collection_Days and Days_With_Data columns
        has_collection_days = 'Collection_Days' in df.columns
        has_days_with_data = 'Days_With_Data' in df.columns
        
        print(f"  Rows: {len(df)}")
        print(f"  Has Collection_Days: {has_collection_days}")
        print(f"  Has Days_With_Data: {has_days_with_data}")
        
        if has_collection_days and has_days_with_data:
            # Check for discrepancies (excluding potential sum rows)
            df_filtered = df.copy()
            
            # Try to identify and exclude sum/total rows
            if 'Period' in df.columns:
                # Exclude rows where Period might be "Total" or similar
                df_filtered = df_filtered[~df_filtered['Period'].astype(str).str.contains('Total|Sum|All', case=False, na=False)]
            
            # Check for Days_With_Data > Collection_Days
            discrepancies = df_filtered[df_filtered['Days_With_Data'] > df_filtered['Collection_Days']]
            
            if not discrepancies.empty:
                issues.append(f"{sheet_name}: {len(discrepancies)} rows where Days_With_Data > Collection_Days")
                print(f"  ⚠️  {len(discrepancies)} discrepancy rows found")
                
                # Show details of discrepancies
                for idx, row in discrepancies.iterrows():
                    period = row.get('Period', f'Row {idx}')
                    collection_days = row['Collection_Days']
                    days_with_data = row['Days_With_Data']
                    print(f"    - {period}: Collection_Days={collection_days}, Days_With_Data={days_with_data}")
            else:
                print(f"  ✅ No discrepancies found")
        
        elif has_days_with_data:
            # Check for unusually high values (excluding potential sum rows)
            max_days = df['Days_With_Data'].max()
            
            # If there's a Period column, check individual periods
            if 'Period' in df.columns:
                individual_periods = df[~df['Period'].astype(str).str.contains('Total|Sum|All', case=False, na=False)]
                if not individual_periods.empty:
                    max_individual = individual_periods['Days_With_Data'].max()
                    if max_individual > 100:  # Still suspiciously high for individual periods
                        issues.append(f"{sheet_name}: Individual period has {max_individual} days with data")
                        print(f"  ⚠️  Individual period max: {max_individual} days")
                    else:
                        print(f"  ✅ Individual period max: {max_individual} days (reasonable)")
                        
                    # Check if there's a total row
                    total_rows = df[df['Period'].astype(str).str.contains('Total|Sum|All', case=False, na=False)]
                    if not total_rows.empty:
                        total_value = total_rows['Days_With_Data'].iloc[0]
                        print(f"  ℹ️  Total/Sum row value: {total_value} days")
            else:
                if max_days > 100:
                    issues.append(f"{sheet_name}: Days_With_Data max value is {max_days}")
                    print(f"  ⚠️  Max Days_With_Data: {max_days}")
                else:
                    print(f"  ✅ Max Days_With_Data: {max_days} (reasonable)")
        
        else:
            print(f"  ℹ️  No relevant day columns found")
    
    except Exception as e:
        print(f"  ❌ Error: {e}")
        issues.append(f"{sheet_name}: Analysis error - {e}")
    
    return issues

if __name__ == "__main__":
    check_collection_day_consistency()
