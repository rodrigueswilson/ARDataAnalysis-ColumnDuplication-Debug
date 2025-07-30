#!/usr/bin/env python3
"""
Comprehensive analysis of all sheets in the Excel report to identify
potential collection day counting issues.
"""

import pandas as pd
import glob
import os

def find_latest_report():
    """Find the most recent AR Analysis Report Excel file."""
    pattern = "AR_Analysis_Report_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        print("No report files found!")
        return None
    
    files.sort()
    latest_file = files[-1]
    print(f"Analyzing report: {latest_file}")
    return latest_file

def analyze_all_sheets():
    """Analyze all sheets for potential collection day counting issues."""
    latest_report = find_latest_report()
    if not latest_report:
        return
    
    try:
        # Get all sheet names
        excel_file = pd.ExcelFile(latest_report)
        sheet_names = excel_file.sheet_names
        
        print(f"\n=== COMPREHENSIVE SHEET ANALYSIS ===")
        print(f"Total sheets found: {len(sheet_names)}")
        print()
        
        # Categorize sheets by potential collection day relevance
        time_based_sheets = []
        acf_pacf_sheets = []
        summary_sheets = []
        other_sheets = []
        
        for sheet_name in sheet_names:
            if 'ACF_PACF' in sheet_name or 'ACF/PACF' in sheet_name:
                acf_pacf_sheets.append(sheet_name)
            elif any(keyword in sheet_name.lower() for keyword in ['daily', 'weekly', 'monthly', 'period', 'biweekly']):
                time_based_sheets.append(sheet_name)
            elif any(keyword in sheet_name.lower() for keyword in ['summary', 'dashboard', 'overview']):
                summary_sheets.append(sheet_name)
            else:
                other_sheets.append(sheet_name)
        
        print(f"=== SHEET CATEGORIZATION ===")
        print(f"ACF/PACF sheets: {len(acf_pacf_sheets)}")
        print(f"Time-based sheets: {len(time_based_sheets)}")
        print(f"Summary sheets: {len(summary_sheets)}")
        print(f"Other sheets: {len(other_sheets)}")
        print()
        
        # Analyze each category for collection day issues
        potential_issues = []
        
        print(f"=== DETAILED ANALYSIS ===")
        
        # 1. Analyze ACF/PACF sheets (highest priority)
        print(f"\n1. ACF/PACF SHEETS (High Priority)")
        print("-" * 50)
        for sheet_name in acf_pacf_sheets:
            issues = analyze_sheet_for_collection_days(latest_report, sheet_name)
            if issues:
                potential_issues.extend(issues)
        
        # 2. Analyze time-based sheets
        print(f"\n2. TIME-BASED SHEETS")
        print("-" * 50)
        for sheet_name in time_based_sheets:
            issues = analyze_sheet_for_collection_days(latest_report, sheet_name)
            if issues:
                potential_issues.extend(issues)
        
        # 3. Analyze summary sheets
        print(f"\n3. SUMMARY SHEETS")
        print("-" * 50)
        for sheet_name in summary_sheets:
            issues = analyze_sheet_for_collection_days(latest_report, sheet_name)
            if issues:
                potential_issues.extend(issues)
        
        # 4. Analyze other sheets
        print(f"\n4. OTHER SHEETS")
        print("-" * 50)
        for sheet_name in other_sheets:
            issues = analyze_sheet_for_collection_days(latest_report, sheet_name)
            if issues:
                potential_issues.extend(issues)
        
        # Summary of findings
        print(f"\n=== SUMMARY OF FINDINGS ===")
        if potential_issues:
            print(f"Found {len(potential_issues)} potential collection day issues:")
            for issue in potential_issues:
                print(f"  - {issue}")
        else:
            print("✓ No obvious collection day counting issues found in other sheets")
            print("✓ The Period Counts (ACF_PACF) fix appears to be the main issue")
        
    except Exception as e:
        print(f"Error analyzing sheets: {e}")

def analyze_sheet_for_collection_days(file_path, sheet_name):
    """Analyze a specific sheet for collection day counting issues."""
    issues = []
    
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        print(f"  Analyzing: {sheet_name}")
        print(f"    Rows: {len(df)}, Columns: {len(df.columns)}")
        
        # Look for relevant columns
        relevant_columns = []
        day_columns = []
        
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['collection', 'days', 'day']):
                relevant_columns.append(col)
            if any(keyword in col_lower for keyword in ['days_with_data', 'unique_days', 'total_days']):
                day_columns.append(col)
        
        if relevant_columns:
            print(f"    Relevant columns: {relevant_columns}")
        
        if day_columns:
            print(f"    Day-related columns: {day_columns}")
        
        # Check for specific patterns that might indicate issues
        
        # 1. Check for Collection_Days vs Days_With_Data discrepancies
        if 'Collection_Days' in df.columns and 'Days_With_Data' in df.columns:
            discrepancies = df[df['Days_With_Data'] > df['Collection_Days']]
            if not discrepancies.empty:
                issues.append(f"{sheet_name}: {len(discrepancies)} rows where Days_With_Data > Collection_Days")
                print(f"    ⚠️  Found {len(discrepancies)} discrepancy rows")
        
        # 2. Check for unusually high day counts
        for col in day_columns:
            if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                max_days = df[col].max()
                if max_days > 100:  # Suspiciously high for most time periods
                    issues.append(f"{sheet_name}: {col} has unusually high value ({max_days})")
                    print(f"    ⚠️  {col} max value: {max_days} (suspiciously high)")
        
        # 3. Check for negative day counts
        for col in day_columns:
            if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                negative_count = (df[col] < 0).sum()
                if negative_count > 0:
                    issues.append(f"{sheet_name}: {col} has {negative_count} negative values")
                    print(f"    ⚠️  {col} has {negative_count} negative values")
        
        # 4. Check for period-based sheets with inconsistent day counts
        if 'Period' in df.columns and any('Days' in str(col) for col in df.columns):
            period_analysis = analyze_period_consistency(df, sheet_name)
            if period_analysis:
                issues.extend(period_analysis)
        
        if not relevant_columns and not day_columns:
            print(f"    ✓ No day-related columns found")
        elif not issues:
            print(f"    ✓ No obvious issues detected")
        
    except Exception as e:
        print(f"    ❌ Error reading {sheet_name}: {e}")
        issues.append(f"{sheet_name}: Could not analyze due to error: {e}")
    
    return issues

def analyze_period_consistency(df, sheet_name):
    """Analyze period-based data for consistency issues."""
    issues = []
    
    try:
        if 'Period' in df.columns:
            # Group by period and check for consistency
            period_groups = df.groupby('Period')
            
            # Look for day-related columns
            day_cols = [col for col in df.columns if 'Days' in str(col) or 'Day' in str(col)]
            
            for col in day_cols:
                if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                    # Check if same periods have different day counts across rows
                    period_variance = period_groups[col].nunique()
                    inconsistent_periods = period_variance[period_variance > 1]
                    
                    if not inconsistent_periods.empty:
                        issues.append(f"{sheet_name}: {col} has inconsistent values within periods: {list(inconsistent_periods.index)}")
    
    except Exception as e:
        pass  # Skip analysis if there are issues
    
    return issues

if __name__ == "__main__":
    analyze_all_sheets()
