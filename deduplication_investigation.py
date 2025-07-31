#!/usr/bin/env python3
"""
Deduplication Logic Investigation Script
=======================================

This script specifically targets the contradiction where deduplication claims
to remove columns but the column count doesn't change. We'll examine the
pandas DataFrame internal state and column handling in detail.
"""

import pandas as pd
import numpy as np
from collections import Counter
import sys
from pathlib import Path

# Add the project root to the path
sys.path.insert(0, str(Path(__file__).parent))

def deep_column_analysis(df, label="DataFrame"):
    """Perform deep analysis of DataFrame column structure"""
    print(f"\n🔍 DEEP ANALYSIS: {label}")
    print(f"  DataFrame shape: {df.shape}")
    print(f"  Column count (len): {len(df.columns)}")
    print(f"  Column count (shape[1]): {df.shape[1]}")
    
    # Analyze column structure
    columns_list = list(df.columns)
    columns_set = set(df.columns)
    
    print(f"  Unique columns (set): {len(columns_set)}")
    print(f"  Total columns (list): {len(columns_list)}")
    
    # Check for duplicate column names
    column_counts = Counter(columns_list)
    duplicates = {col: count for col, count in column_counts.items() if count > 1}
    
    if duplicates:
        print(f"  ❌ DUPLICATES DETECTED: {len(duplicates)} unique duplicate names")
        for col, count in sorted(duplicates.items()):
            print(f"    '{col}': appears {count} times")
        
        # Show positions of duplicates
        print(f"  DUPLICATE POSITIONS:")
        for col, count in duplicates.items():
            positions = [i for i, c in enumerate(columns_list) if c == col]
            print(f"    '{col}': positions {positions}")
    else:
        print(f"  ✅ No duplicate column names found")
    
    # Check DataFrame index
    print(f"  DataFrame index: {type(df.index)}, length: {len(df.index)}")
    
    # Check for hidden characters in column names
    print(f"  COLUMN NAME ANALYSIS:")
    suspicious_columns = []
    for i, col in enumerate(columns_list):
        col_str = str(col)
        if len(col_str) != len(col_str.strip()):
            suspicious_columns.append((i, col, f"whitespace: '{col_str}'"))
        if '\x00' in col_str or any(ord(c) < 32 for c in col_str if c not in '\t\n\r'):
            suspicious_columns.append((i, col, f"control chars: {repr(col_str)}"))
    
    if suspicious_columns:
        print(f"    ❌ SUSPICIOUS COLUMN NAMES: {len(suspicious_columns)}")
        for pos, col, issue in suspicious_columns[:5]:  # Show first 5
            print(f"      Position {pos}: {issue}")
    else:
        print(f"    ✅ All column names appear clean")
    
    return {
        'total_columns': len(columns_list),
        'unique_columns': len(columns_set),
        'duplicates': duplicates,
        'suspicious_columns': suspicious_columns
    }

def test_pandas_deduplication_methods(df):
    """Test different pandas deduplication approaches"""
    print(f"\n🧪 TESTING PANDAS DEDUPLICATION METHODS")
    
    original_count = len(df.columns)
    print(f"Original column count: {original_count}")
    
    # Method 1: Standard duplicated() approach (what we're using)
    print(f"\n  Method 1: df.loc[:, ~df.columns.duplicated()]")
    try:
        method1_result = df.loc[:, ~df.columns.duplicated()]
        print(f"    Result columns: {len(method1_result.columns)}")
        print(f"    Duplicated mask: {df.columns.duplicated().sum()} True values")
        print(f"    Duplicated positions: {list(np.where(df.columns.duplicated())[0])}")
        
        # Check what duplicated() actually detects
        dup_mask = df.columns.duplicated()
        print(f"    Duplicated details:")
        for i, (col, is_dup) in enumerate(zip(df.columns, dup_mask)):
            if is_dup:
                print(f"      Position {i}: '{col}' marked as duplicate")
        
    except Exception as e:
        print(f"    ❌ Method 1 failed: {e}")
        method1_result = df
    
    # Method 2: Using groupby to keep last occurrence
    print(f"\n  Method 2: Using groupby approach")
    try:
        # Create a mapping of unique column names
        col_mapping = {}
        for i, col in enumerate(df.columns):
            if col not in col_mapping:
                col_mapping[col] = []
            col_mapping[col].append(i)
        
        # Keep only the last occurrence of each column
        cols_to_keep = []
        for col, positions in col_mapping.items():
            cols_to_keep.append(positions[-1])  # Keep last occurrence
        
        method2_result = df.iloc[:, sorted(cols_to_keep)]
        print(f"    Result columns: {len(method2_result.columns)}")
        print(f"    Removed positions: {set(range(len(df.columns))) - set(cols_to_keep)}")
        
    except Exception as e:
        print(f"    ❌ Method 2 failed: {e}")
        method2_result = df
    
    # Method 3: Manual column name deduplication
    print(f"\n  Method 3: Manual column renaming")
    try:
        new_columns = []
        seen_columns = {}
        
        for col in df.columns:
            if col not in seen_columns:
                new_columns.append(col)
                seen_columns[col] = 1
            else:
                seen_columns[col] += 1
                new_columns.append(f"{col}_dup_{seen_columns[col]}")
        
        method3_df = df.copy()
        method3_df.columns = new_columns
        
        # Now remove the renamed duplicates
        method3_result = method3_df.loc[:, ~method3_df.columns.str.contains('_dup_')]
        print(f"    Result columns: {len(method3_result.columns)}")
        print(f"    Renamed duplicates: {len([c for c in new_columns if '_dup_' in str(c)])}")
        
    except Exception as e:
        print(f"    ❌ Method 3 failed: {e}")
        method3_result = df
    
    # Method 4: Using set operations
    print(f"\n  Method 4: Set-based approach")
    try:
        unique_cols = list(dict.fromkeys(df.columns))  # Preserves order, removes duplicates
        method4_result = df[unique_cols]
        print(f"    Result columns: {len(method4_result.columns)}")
        print(f"    Unique column list length: {len(unique_cols)}")
        
    except Exception as e:
        print(f"    ❌ Method 4 failed: {e}")
        method4_result = df
    
    return {
        'method1': method1_result,
        'method2': method2_result,
        'method3': method3_result,
        'method4': method4_result
    }

def investigate_deduplication_contradiction():
    """Create a test case that reproduces the deduplication contradiction"""
    print("=" * 60)
    print("DEDUPLICATION CONTRADICTION INVESTIGATION")
    print("=" * 60)
    
    # Create a test DataFrame with known duplicate columns
    print("\n1. CREATING TEST DATAFRAME WITH KNOWN DUPLICATES")
    
    # Create base data
    base_data = {
        'Date': ['2023-01-01', '2023-01-02', '2023-01-03'],
        'Total_Files': [10, 15, 12],
        'MP3_Files': [5, 8, 6],
        'JPG_Files': [5, 7, 6]
    }
    
    # Create DataFrame and deliberately add duplicates
    df = pd.DataFrame(base_data)
    
    # Add duplicate columns using different methods
    df['Total_Files'] = df['Total_Files'] * 2  # Overwrite (should not create duplicate)
    df = pd.concat([df, df[['Total_Files', 'MP3_Files']]], axis=1)  # This should create duplicates
    df = pd.concat([df, df[['Total_Files']]], axis=1)  # Add another duplicate
    
    print(f"Test DataFrame created with shape: {df.shape}")
    
    # Analyze the test DataFrame
    analysis = deep_column_analysis(df, "Test DataFrame with Duplicates")
    
    # Test deduplication methods
    results = test_pandas_deduplication_methods(df)
    
    # Compare results
    print(f"\n📊 DEDUPLICATION RESULTS COMPARISON:")
    for method_name, result_df in results.items():
        print(f"  {method_name}: {len(df.columns)} -> {len(result_df.columns)} columns")
        if len(result_df.columns) < len(df.columns):
            print(f"    ✅ Successfully reduced columns")
        else:
            print(f"    ❌ No reduction in columns")
    
    return df, results

def reproduce_production_issue():
    """Try to reproduce the exact production issue"""
    print(f"\n" + "=" * 60)
    print("REPRODUCING PRODUCTION ISSUE")
    print("=" * 60)
    
    try:
        # Import the actual production modules
        from db_utils import get_db_connection
        from report_generator.sheet_creators import SheetCreator
        from report_generator.formatters import ExcelFormatter
        import json
        
        # Get real data from the system
        db = get_db_connection()
        sheet_creator = SheetCreator(db, formatter=ExcelFormatter())
        
        # Load the configuration
        config_path = Path(__file__).parent / "report_config.json"
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        # Find the Daily Counts ACF/PACF sheet
        daily_sheet = next((s for s in config.get('sheets', []) if 'Daily Counts (ACF_PACF)' in s.get('name', '')), None)
        
        if daily_sheet:
            print(f"Found test sheet: {daily_sheet['name']}")
            
            # Get the pipeline data
            from pipelines import PIPELINES
            pipeline = PIPELINES.get(daily_sheet['pipeline'])
            
            if pipeline:
                print(f"Running pipeline: {daily_sheet['pipeline']}")
                
                # Run the aggregation
                df = sheet_creator._run_aggregation(pipeline)
                print(f"Pipeline result shape: {df.shape}")
                
                # Analyze the pipeline result
                analysis1 = deep_column_analysis(df, "After Pipeline")
                
                # Apply ACF/PACF analysis
                from ar_utils import add_acf_pacf_analysis
                df_with_acf = add_acf_pacf_analysis(df)
                
                analysis2 = deep_column_analysis(df_with_acf, "After ACF/PACF")
                
                # Test our deduplication on the real data
                print(f"\n🔧 TESTING DEDUPLICATION ON REAL DATA")
                
                original_count = len(df_with_acf.columns)
                
                # Use the exact same logic as in production
                duplicated_mask = df_with_acf.columns.duplicated()
                num_duplicates = duplicated_mask.sum()
                
                print(f"Original columns: {original_count}")
                print(f"Duplicated mask sum: {num_duplicates}")
                print(f"Duplicated positions: {list(np.where(duplicated_mask)[0])}")
                
                if num_duplicates > 0:
                    print(f"Columns marked as duplicates:")
                    for i, (col, is_dup) in enumerate(zip(df_with_acf.columns, duplicated_mask)):
                        if is_dup:
                            print(f"  Position {i}: '{col}'")
                
                # Apply deduplication
                deduplicated_df = df_with_acf.loc[:, ~duplicated_mask]
                final_count = len(deduplicated_df.columns)
                
                print(f"After deduplication: {final_count}")
                print(f"Expected reduction: {num_duplicates}")
                print(f"Actual reduction: {original_count - final_count}")
                
                if num_duplicates > 0 and final_count == original_count:
                    print(f"❌ CONTRADICTION REPRODUCED!")
                    print(f"   Claims to remove {num_duplicates} duplicates but count unchanged")
                    
                    # Deep dive into the contradiction
                    print(f"\n🕵️ INVESTIGATING THE CONTRADICTION:")
                    
                    # Check if the mask is actually working
                    mask_true_count = (~duplicated_mask).sum()
                    print(f"   Inverse mask true count: {mask_true_count}")
                    print(f"   Should result in {mask_true_count} columns")
                    
                    # Check if there's an issue with the indexing
                    print(f"   Trying manual indexing...")
                    manual_df = df_with_acf.iloc[:, ~duplicated_mask]
                    print(f"   Manual indexing result: {len(manual_df.columns)} columns")
                    
                    # Check column data types
                    print(f"   Column index type: {type(df_with_acf.columns)}")
                    print(f"   Duplicated mask type: {type(duplicated_mask)}")
                    
                else:
                    print(f"✅ Deduplication worked correctly")
                
            else:
                print(f"❌ Pipeline not found: {daily_sheet['pipeline']}")
        else:
            print("❌ Daily Counts ACF/PACF sheet not found in config")
            
    except Exception as e:
        print(f"❌ Error reproducing production issue: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # Run both investigations
    print("Starting deduplication investigation...")
    
    # First, test with artificial data
    test_df, test_results = investigate_deduplication_contradiction()
    
    # Then, try to reproduce with real production data
    reproduce_production_issue()