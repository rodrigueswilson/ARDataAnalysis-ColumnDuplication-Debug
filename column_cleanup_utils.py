"""
Column Cleanup Utilities
========================

Utilities for cleaning up duplicate ACF/PACF columns caused by
inconsistent naming conventions.
"""

import pandas as pd
import re


def cleanup_duplicate_acf_pacf_columns(df, value_col="Total_Files"):
    """
    Clean up duplicate ACF/PACF columns in a DataFrame.
    
    This function removes duplicate ACF/PACF columns that may have been created
    by inconsistent naming conventions, keeping only the most recent/complete set.
    
    Args:
        df: DataFrame to clean up
        value_col: Base column name for ACF/PACF analysis
        
    Returns:
        DataFrame with duplicate columns removed
    """
    df_clean = df.copy()
    
    # Find all ACF/PACF columns with both naming patterns
    acf_pacf_pattern = rf'{value_col}_(ACF|PACF)_Lag_?(\d+)(_Significant)?'
    acf_pacf_cols = {}
    
    for col in df_clean.columns:
        match = re.match(acf_pacf_pattern, str(col))
        if match:
            analysis_type = match.group(1)  # ACF or PACF
            lag = int(match.group(2))
            is_significant = match.group(3) is not None
            
            # Create standardized column name (with underscore)
            if is_significant:
                std_name = f'{value_col}_{analysis_type}_Lag_{lag}_Significant'
            else:
                std_name = f'{value_col}_{analysis_type}_Lag_{lag}'
            
            # Keep track of columns for this standardized name
            if std_name not in acf_pacf_cols:
                acf_pacf_cols[std_name] = []
            acf_pacf_cols[std_name].append(col)
    
    # Remove duplicates, keeping the last column for each standardized name
    columns_to_drop = []
    columns_to_rename = {}
    
    for std_name, col_list in acf_pacf_cols.items():
        if len(col_list) > 1:
            # Multiple columns for the same logical column
            print(f"    [CLEANUP] Found {len(col_list)} duplicate columns for {std_name}: {col_list}")
            
            # Keep the last one (most recent) and drop others
            keep_col = col_list[-1]
            drop_cols = col_list[:-1]
            
            columns_to_drop.extend(drop_cols)
            if keep_col != std_name:
                columns_to_rename[keep_col] = std_name
            
            print(f"    [CLEANUP] Keeping: {keep_col}, Dropping: {drop_cols}")
    
    # Apply cleanup
    if columns_to_drop:
        df_clean = df_clean.drop(columns=columns_to_drop)
        print(f"    [CLEANUP] Dropped {len(columns_to_drop)} duplicate columns")
    
    if columns_to_rename:
        df_clean = df_clean.rename(columns=columns_to_rename)
        print(f"    [CLEANUP] Renamed {len(columns_to_rename)} columns to standard format")
    
    return df_clean


def make_robust_existence_check(df, value_col):
    """
    Robust existence check that detects both naming conventions.
    
    Args:
        df: DataFrame to check
        value_col: Base column name for ACF/PACF analysis
        
    Returns:
        List of existing ACF/PACF columns
    """
    existing_acf_cols = []
    for col in df.columns:
        col_str = str(col)
        # Check for both underscore patterns: ACF_Lag_ and ACF_Lag
        if (f'{value_col}_ACF_Lag_' in col_str or 
            f'{value_col}_ACF_Lag' in col_str and f'{value_col}_ACF_Lag_' not in col_str):
            existing_acf_cols.append(col)
    
    return existing_acf_cols
