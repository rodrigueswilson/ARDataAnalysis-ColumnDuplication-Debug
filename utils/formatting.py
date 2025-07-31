"""
Data Formatting Utilities for AR Data Analysis

This module provides utilities for formatting data, converting between different
formats, and preparing data for display in reports and analysis.

Key Functions:
- Duration formatting (seconds to ISO, HH:MM:SS)
- Data type conversions
- Display formatting utilities
- Column reordering for reports
"""

import pandas as pd
import numpy as np
from typing import List, Union


def seconds_to_iso_duration(seconds: Union[int, float]) -> str:
    """
    Converts a duration in seconds to an ISO 8601 duration string (e.g., PT1M35S).
    
    Args:
        seconds: Duration in seconds
        
    Returns:
        ISO 8601 duration string
    """
    if pd.isna(seconds) or seconds < 0:
        return "PT0S"
    
    seconds = int(seconds)
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    
    duration_parts = []
    if hours > 0:
        duration_parts.append(f"{hours}H")
    if minutes > 0:
        duration_parts.append(f"{minutes}M")
    if secs > 0 or not duration_parts:  # Always include seconds if no other parts
        duration_parts.append(f"{secs}S")
    
    return "PT" + "".join(duration_parts)


def seconds_to_hms(seconds: Union[int, float]) -> str:
    """
    Converts a duration in seconds to an HH:MM:SS format string.
    
    Args:
        seconds: Duration in seconds
        
    Returns:
        Duration string in HH:MM:SS format
    """
    if pd.isna(seconds) or seconds < 0:
        return "00:00:00"
    
    seconds = int(seconds)
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def format_file_size(size_mb: Union[int, float], precision: int = 2) -> str:
    """
    Formats file size in MB with appropriate units.
    
    Args:
        size_mb: Size in megabytes
        precision: Number of decimal places
        
    Returns:
        Formatted size string with units
    """
    if pd.isna(size_mb) or size_mb < 0:
        return "0.00 MB"
    
    if size_mb < 1:
        return f"{size_mb * 1024:.{precision}f} KB"
    elif size_mb >= 1024:
        return f"{size_mb / 1024:.{precision}f} GB"
    else:
        return f"{size_mb:.{precision}f} MB"


def format_percentage(value: Union[int, float], precision: int = 1) -> str:
    """
    Formats a value as a percentage.
    
    Args:
        value: Value to format (0-1 range expected)
        precision: Number of decimal places
        
    Returns:
        Formatted percentage string
    """
    if pd.isna(value):
        return "N/A"
    
    return f"{value * 100:.{precision}f}%"


def format_count(count: Union[int, float]) -> str:
    """
    Formats a count value with appropriate thousands separators.
    
    Args:
        count: Count value to format
        
    Returns:
        Formatted count string
    """
    if pd.isna(count):
        return "0"
    
    return f"{int(count):,}"


def reorder_with_acf_pacf(df: pd.DataFrame, base_order: List[str]) -> pd.DataFrame:
    """
    Reorders DataFrame columns to place ACF/PACF columns after base metrics.

    This function intelligently finds all ACF/PACF-related columns (including
    significance indicators) and reorders them to appear immediately after the
    base metrics columns, maintaining a logical flow for analysis.

    Args:
        df: DataFrame to reorder
        base_order: List of base columns to appear first (e.g. Date, Total_Files)
        
    Returns:
        DataFrame with columns reordered according to the pattern
    """
    # Find ACF/PACF columns
    acf_pacf_cols = []
    for col in df.columns:
        if any(pattern in col for pattern in ['_ACF_', '_PACF_', '_Significant']):
            acf_pacf_cols.append(col)
    
    # Sort ACF/PACF columns for consistency
    acf_pacf_cols.sort()
    
    # Get remaining columns (not in base_order or ACF/PACF)
    remaining_cols = []
    for col in df.columns:
        if col not in base_order and col not in acf_pacf_cols:
            remaining_cols.append(col)
    
    # Build final column order: base -> ACF/PACF -> remaining
    final_order = []
    
    # Add base columns that exist in the DataFrame
    for col in base_order:
        if col in df.columns:
            final_order.append(col)
    
    # Add ACF/PACF columns
    final_order.extend(acf_pacf_cols)
    
    # Add remaining columns
    final_order.extend(remaining_cols)
    
    return df[final_order]


def reorder_with_forecast_columns(df: pd.DataFrame, base_order: List[str], value_col: str = "Total_Files") -> pd.DataFrame:
    """
    Reorders DataFrame columns with forecast columns after ACF/PACF columns.
    Extends the existing reorder_with_acf_pacf pattern.
    
    Args:
        df: DataFrame to reorder
        base_order: List of base column names in desired order
        value_col: Base column name for forecast columns
        
    Returns:
        DataFrame with reordered columns including forecasts
    """
    # First apply existing ACF/PACF reordering
    df_reordered = reorder_with_acf_pacf(df, base_order)

    # Find forecast columns
    forecast_cols = []
    for col in df.columns:
        if col.startswith(f'{value_col}_Forecast'):
            forecast_cols.append(col)

    # Sort forecast columns for consistency
    forecast_cols.sort()

    # Get current column order and add forecast columns at the end
    current_cols = list(df_reordered.columns)

    # Remove forecast columns from current position
    for col in forecast_cols:
        if col in current_cols:
            current_cols.remove(col)

    # Add forecast columns at the end
    final_order = current_cols + forecast_cols

    return df_reordered[final_order]


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans and standardizes column names in a DataFrame.
    
    Args:
        df: DataFrame with columns to clean
        
    Returns:
        DataFrame with cleaned column names
    """
    df_copy = df.copy()
    
    # Replace common problematic characters
    new_columns = []
    for col in df_copy.columns:
        # Replace spaces with underscores
        clean_col = str(col).replace(' ', '_')
        # Remove special characters except underscores
        clean_col = ''.join(c for c in clean_col if c.isalnum() or c == '_')
        # Ensure it doesn't start with a number
        if clean_col and clean_col[0].isdigit():
            clean_col = f"col_{clean_col}"
        new_columns.append(clean_col)
    
    df_copy.columns = new_columns
    return df_copy


def standardize_date_column(df: pd.DataFrame, date_col: str = 'Date') -> pd.DataFrame:
    """
    Standardizes a date column to ensure consistent formatting.
    
    Args:
        df: DataFrame containing the date column
        date_col: Name of the date column to standardize
        
    Returns:
        DataFrame with standardized date column
    """
    if date_col not in df.columns:
        return df
    
    df_copy = df.copy()
    
    # Convert to datetime if not already
    if not pd.api.types.is_datetime64_any_dtype(df_copy[date_col]):
        df_copy[date_col] = pd.to_datetime(df_copy[date_col], errors='coerce')
    
    # Format as YYYY-MM-DD string
    df_copy[date_col] = df_copy[date_col].dt.strftime('%Y-%m-%d')
    
    return df_copy
