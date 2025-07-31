"""
AR Data Analysis Utilities Package

This package provides modular utilities for the AR Data Analysis system,
decomposed from the original ar_utils.py into focused modules for better
maintainability and organization.

Modules:
- config: Configuration management and loading
- calendar: Calendar and date utilities
- formatting: Data formatting and display utilities
- time_series: Time series analysis (ACF/PACF, ARIMA)
- enrichment: Data enrichment and contextual information

This module provides backward compatibility by re-exporting all functions
from the original ar_utils.py interface.
"""

# Import all functions from submodules for backward compatibility
from .config import (
    get_school_calendar,
    get_non_collection_days,
    get_activity_schedule,
    get_config,
    reload_config,
    is_config_loaded
)

from .calendar import (
    calculate_collection_days_for_period,
    precompute_collection_days,
    is_collection_day,
    get_period_for_date,
    get_school_year_for_date,
    get_date_range_for_period,
    get_all_periods,
    get_periods_for_school_year
)

from .formatting import (
    seconds_to_iso_duration,
    seconds_to_hms,
    format_file_size,
    format_percentage,
    format_count,
    reorder_with_acf_pacf,
    reorder_with_forecast_columns,
    clean_column_names,
    standardize_date_column
)

from .time_series import (
    add_acf_pacf_analysis,
    infer_sheet_type,
    generate_arima_forecast,
    add_arima_forecast_to_dataframe,
    STATSMODELS_AVAILABLE
)

from .enrichment import (
    get_contextual_info,
    enrich_file_metadata,
    validate_enriched_data,
    get_file_category,
    summarize_enrichment_stats
)

# Re-export the CONFIG for backward compatibility
from .config import CONFIG

# Package metadata
__version__ = "2.0.0"
__author__ = "AR Data Analysis Team"
__description__ = "Modular utilities for AR Data Analysis system"

# List of all exported functions for backward compatibility
__all__ = [
    # Configuration functions
    'get_school_calendar',
    'get_non_collection_days', 
    'get_activity_schedule',
    'get_config',
    'reload_config',
    'is_config_loaded',
    'CONFIG',
    
    # Calendar functions
    'calculate_collection_days_for_period',
    'precompute_collection_days',
    'is_collection_day',
    'get_period_for_date',
    'get_school_year_for_date',
    'get_date_range_for_period',
    'get_all_periods',
    'get_periods_for_school_year',
    
    # Formatting functions
    'seconds_to_iso_duration',
    'seconds_to_hms',
    'format_file_size',
    'format_percentage',
    'format_count',
    'reorder_with_acf_pacf',
    'reorder_with_forecast_columns',
    'clean_column_names',
    'standardize_date_column',
    
    # Time series functions
    'add_acf_pacf_analysis',
    'infer_sheet_type',
    'generate_arima_forecast',
    'add_arima_forecast_to_dataframe',
    'STATSMODELS_AVAILABLE',
    
    # Enrichment functions
    'get_contextual_info',
    'enrich_file_metadata',
    'validate_enriched_data',
    'get_file_category',
    'summarize_enrichment_stats'
]


def get_utils_summary():
    """
    Returns a summary of available utility functions organized by module.
    
    Returns:
        Dictionary containing function lists by module
    """
    return {
        'config': [
            'get_school_calendar', 'get_non_collection_days', 'get_activity_schedule',
            'get_config', 'reload_config', 'is_config_loaded'
        ],
        'calendar': [
            'calculate_collection_days_for_period', 'precompute_collection_days',
            'is_collection_day', 'get_period_for_date', 'get_school_year_for_date',
            'get_date_range_for_period', 'get_all_periods', 'get_periods_for_school_year'
        ],
        'formatting': [
            'seconds_to_iso_duration', 'seconds_to_hms', 'format_file_size',
            'format_percentage', 'format_count', 'reorder_with_acf_pacf',
            'reorder_with_forecast_columns', 'clean_column_names', 'standardize_date_column'
        ],
        'time_series': [
            'add_acf_pacf_analysis', 'infer_sheet_type', 'generate_arima_forecast',
            'add_arima_forecast_to_dataframe'
        ],
        'enrichment': [
            'get_contextual_info', 'enrich_file_metadata', 'validate_enriched_data',
            'get_file_category', 'summarize_enrichment_stats'
        ]
    }
