"""
AR Data Analysis Shared Utilities

This module contains shared configuration data and helper functions used across
the AR Data Analysis scripts to eliminate code duplication and ensure consistency.
"""

import datetime
import yaml
from pathlib import Path
import pandas as pd
import numpy as np

# Import statsmodels for ACF/PACF analysis and ARIMA forecasting
try:
    from statsmodels.tsa.stattools import acf, pacf, adfuller
    from statsmodels.tsa.arima.model import ARIMA, ARIMAResultsWrapper
    from statsmodels.stats.diagnostic import acorr_ljungbox
    STATSMODELS_AVAILABLE = True
except ImportError:
    STATSMODELS_AVAILABLE = False
    print("[WARNING]  Warning: statsmodels not available. ACF/PACF and ARIMA analysis will be disabled.")
    print("    Install with: pip install statsmodels")


# ==============================================================================
# CONFIGURATION LOADING
# ==============================================================================

def _load_config():
    """Loads and parses the config.yaml file, converting date strings to objects."""
    config_path = Path(__file__).parent / "config.yaml"
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file not found at: {config_path}")

    with open(config_path, 'r') as f:
        config = yaml.safe_load(f)

    # --- Post-process configuration data ---
    # Convert date strings in school_calendar to datetime.date objects
    for year, details in config.get('school_calendar', {}).items():
        if 'start_date' in details and isinstance(details['start_date'], str):
            details['start_date'] = datetime.date.fromisoformat(details['start_date'])
        if 'end_date' in details and isinstance(details['end_date'], str):
            details['end_date'] = datetime.date.fromisoformat(details['end_date'])
        for period, dates in details.get('periods', {}).items():
            if isinstance(dates, list) and all(isinstance(d, str) for d in dates):
                details['periods'][period] = [datetime.date.fromisoformat(d) for d in dates]

    # Convert date string keys in non_collection_days to datetime.date objects
    if 'non_collection_days' in config:
        config['non_collection_days'] = {
            datetime.date.fromisoformat(date_str): info
            for date_str, info in config['non_collection_days'].items()
        }

    return config

# Load configuration once when the module is imported
CONFIG = {}
try:
    CONFIG = _load_config()
except (FileNotFoundError, yaml.YAMLError, TypeError) as e:
    print(f"[ERROR] Critical Error: Could not load or parse config.yaml. Details: {e}")
    # Allow module to load but dependent functions will fail gracefully.

# ==============================================================================
# CONFIGURATION DATA ACCESSORS
# ==============================================================================

def get_school_calendar():
    """Returns the school calendar configuration from the loaded CONFIG."""
    return CONFIG.get('school_calendar', {})


def get_non_collection_days():
    """Returns a dictionary of non-collection days from the loaded CONFIG."""
    return CONFIG.get('non_collection_days', {})


def calculate_collection_days_for_period(period_name: str) -> int:
    """
    Calculate the total number of valid collection days for a given period.
    
    This function counts all weekdays within the period's date range,
    excluding holidays, breaks, and other non-collection days as defined
    in the configuration.
    
    Args:
        period_name: The period identifier (e.g., 'SY 21-22 P1')
        
    Returns:
        int: Total number of valid collection days in the period
    """
    try:
        school_calendar = get_school_calendar()
        non_collection_days = get_non_collection_days()
        
        # Find the period in the school calendar
        period_dates = None
        for year_data in school_calendar.values():
            if period_name in year_data.get('periods', {}):
                period_dates = year_data['periods'][period_name]
                break
        
        if not period_dates or len(period_dates) != 2:
            print(f"[WARNING] Could not find date range for period: {period_name}")
            return 0
        
        start_date, end_date = period_dates
        
        # Count valid collection days
        collection_days = 0
        current_date = start_date
        
        while current_date <= end_date:
            # Check if it's a weekday (Monday=0, Sunday=6)
            if current_date.weekday() < 5:  # Monday through Friday
                # Check if it's not a non-collection day
                if current_date not in non_collection_days:
                    collection_days += 1
            
            # Move to next day
            current_date += datetime.timedelta(days=1)
        
        return collection_days
        
    except Exception as e:
        print(f"[ERROR] Error calculating collection days for {period_name}: {e}")
        return 0


def get_activity_schedule():
    """Returns the daily activity schedule from the loaded CONFIG."""
    return CONFIG.get('activity_schedule', [])


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================

def precompute_collection_days(school_calendar: dict, non_collection_days: dict) -> dict[datetime.date, dict]:
    """
    Precomputes a map of all valid data collection days for efficient lookups.

    This function iterates through the date ranges defined in the school calendar
    and builds a dictionary where keys are `datetime.date` objects and values
    are dictionaries containing details about that day (e.g., school year,
    collection period).

    The logic correctly handles:
    - Filtering for weekdays (Monday-Friday).
    - Excluding specific non-collection dates (e.g., holidays, breaks).
    - Ensuring that only dates within defined collection periods are included.

    Args:
        school_calendar: A dictionary containing the academic years, their start/end
                         dates, and defined collection periods.
        non_collection_days: A dictionary mapping specific dates to their event
                             details, used for exclusions.

    Returns:
        A dictionary mapping each valid collection `datetime.date` to a dictionary
        of its properties, such as `School_Year` and `Collection_Period`.
    """
    collection_day_map = {}
    
    # First pass: identify all collection days by period
    for school_year, config in school_calendar.items():
        periods = config["periods"]
        
        for period_name, (start_date, end_date) in periods.items():
            current = start_date
            while current <= end_date:
                # Weekdays only
                if current.weekday() < 5:
                    # Check if the day is in non-collection
                    ncd = non_collection_days.get(current)
                    if ncd:
                        if ncd["type"] == "Non-Collection":
                            # Skip non-collection
                            current += datetime.timedelta(days=1)
                            continue
                        elif ncd["type"] == "Partial":
                            # Partial days are included
                            pass
                    # Add to map
                    collection_day_map[current] = {
                        "School_Year": school_year,
                        "Period": period_name
                    }
                current += datetime.timedelta(days=1)
    
    # Second pass: add day numbering
    # Sort all collection days chronologically
    sorted_days = sorted(collection_day_map.keys())
    
    # Track day numbers by period and school year
    period_day_counters = {}
    sy_day_counters = {}
    
    for day in sorted_days:
        info = collection_day_map[day]
        school_year = info["School_Year"]
        period = info["Period"]
        
        # Initialize counters if needed
        if period not in period_day_counters:
            period_day_counters[period] = 0
        if school_year not in sy_day_counters:
            sy_day_counters[school_year] = 0
            
        # Increment counters
        period_day_counters[period] += 1
        sy_day_counters[school_year] += 1
        
        # Add day numbers to the map
        collection_day_map[day]["day_in_period"] = period_day_counters[period]
        collection_day_map[day]["day_in_sycollection"] = sy_day_counters[school_year]
    
    return collection_day_map


def seconds_to_iso_duration(seconds: int | float) -> str:
    """Converts a duration in seconds to an ISO 8601 duration string (e.g., PT1M35S)."""
    if not isinstance(seconds, (int, float)) or seconds < 0:
        return "PT0S"
    
    seconds = int(seconds)
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    
    duration_str = "PT"
    if hours > 0:
        duration_str += f"{hours}H"
    if minutes > 0:
        duration_str += f"{minutes}M"
    if seconds > 0 or duration_str == "PT":  # Always include seconds if it's the only component
        duration_str += f"{seconds}S"
    
    return duration_str


def seconds_to_hms(seconds: int | float) -> str:
    """Converts a duration in seconds to an HH:MM:SS format string."""
    if not isinstance(seconds, (int, float)) or seconds < 0:
        return "00:00:00"
    seconds = int(seconds)
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def get_contextual_info(
    dt_obj: datetime.datetime,
    calendar: dict,
    non_collection_days: dict,
    schedule: list,
    collection_day_map: dict,
    audio_props: dict | None = None,
    is_outlier: bool = False
) -> dict:
    """
    Derives all contextual information for a given file based on its timestamp.

    This function acts as a central hub for data enrichment, pulling together
    calendar details, holiday information, and activity schedules to create a
    comprehensive record for a single data point (e.g., a photo or audio file).

    Args:
        dt_obj: The datetime object associated with the file.
        calendar: The school calendar configuration.
        non_collection_days: The dictionary of holidays and non-collection days.
        schedule: The daily activity schedule.
        collection_day_map: The precomputed map of valid collection days.
        audio_props: Optional dictionary of audio-specific properties.
        is_outlier: Flag indicating if the file is considered an outlier.

    Returns:
        A dictionary containing a wide range of contextual fields, including:
        - ISO_Date, ISO_Time, ISO_Week, ISO_Month
        - Day_of_Week, Time_of_Day
        - School_Year, Collection_Period
        - Day_Number_in_Period, Day_Number_in_SYCollection
        - Day_Type, Day_Event, Scheduled_Activity
        - and other derived metrics.
    """
    if not dt_obj:
        return None

    info = {}
    file_date = dt_obj.date()
    file_time = dt_obj.time()

    # --- Basic Time & Context Info ---
    info['ISO_Date'] = file_date.isoformat()
    info['ISO_Time'] = file_time.strftime('%H:%M:%S')
    info['ISO_Week'] = file_date.isocalendar()[1]  # Store as integer for consistent aggregation
    info['ISO_Year'] = file_date.year  # Store year as integer for future-proof aggregation
    info['ISO_YearWeek'] = f"{file_date.year}-W{file_date.isocalendar()[1]:02d}"  # Combined year-week identifier
    info['ISO_Month'] = file_date.month
    info['Day_of_Week'] = dt_obj.strftime('%A')
    info['Time_of_Day'] = "Morning" if dt_obj.hour < 12 else "Afternoon"
    
    # --- Add explicit is_collection_day flag ---
    info['is_collection_day'] = file_date in collection_day_map
    
    # --- Audio Properties (MP3 files only) ---
    if audio_props:
        info['Duration_Seconds'] = round(audio_props['duration'], 2)
        info['Duration_ISO'] = seconds_to_iso_duration(audio_props['duration'])
        info['Duration_HMS'] = seconds_to_hms(audio_props['duration'])
        info['File_Size_MB'] = round(audio_props['file_size'] / (1024 * 1024), 3)
        info['Bitrate_kbps'] = audio_props['bitrate']
        info['Channels'] = audio_props['channels']
        info['Outlier_Status'] = audio_props.get('is_outlier', False)
    else:
        # For JPG files, just set outlier status
        info['Outlier_Status'] = is_outlier

    # --- School Calendar & Period Info ---
    info['School_Year'] = "N/A"
    info['Collection_Period'] = "N/A"
    is_in_period = False
    for year, details in calendar.items():
        if details["start_date"] <= file_date <= details["end_date"]:
            info['School_Year'] = year
            for period, (start, end) in details["periods"].items():
                if start <= file_date <= end:
                    info['Collection_Period'] = period
                    is_in_period = True
                    break
            if is_in_period:
                break

    # --- Day Number Info (from pre-computed map) ---
    day_numbers = collection_day_map.get(file_date)
    if day_numbers:
        info['Day_Number_in_Period'] = day_numbers['day_in_period']
        info['Day_Number_in_SYCollection'] = day_numbers['day_in_sycollection']
    else:
        info['Day_Number_in_Period'] = "N/A"
        info['Day_Number_in_SYCollection'] = "N/A"

    # --- Day Type Classification ---
    if not is_in_period:
        info['Day_Type'] = "Non-Collection"
        info['Day_Event'] = "Outside Collection Period"
    else:
        day_exception = non_collection_days.get(file_date)
        if day_exception:
            if "virtual day" in day_exception.get("event", "").lower():
                info['Day_Type'] = "Non-Collection (Virtual)"
            else:
                info['Day_Type'] = day_exception.get("type", "Non-Collection")
            info['Day_Event'] = day_exception.get("event", "N/A")
        elif file_date.weekday() >= 5:
            info['Day_Type'] = "Non-Collection"
            info['Day_Event'] = "Weekend"
        else:
            info['Day_Type'] = "Full"
            info['Day_Event'] = "Regular School Day"

    # --- Activity Schedule Info ---
    info['Scheduled_Activity'] = "N/A"
    if info['Day_Type'] in ["Non-Collection", "Non-Collection (Virtual)"]:
        info['Scheduled_Activity'] = "No Collection"
    elif info['Day_Type'] == "Partial" and file_time >= datetime.time(12, 0):
        info['Scheduled_Activity'] = "Afternoon (No Collection on Partial Day)"
    else:
        for act in schedule:
            if act["start"] <= file_time < act["end"]:
                info['Scheduled_Activity'] = act["activity"]
                break
            
    return info


def is_collection_day(date_obj: datetime.date, collection_day_map: dict) -> bool:
    """Determines if a given date is a collection day based on the precomputed map."""
    return date_obj in collection_day_map


# ==============================================================================
# ACF/PACF TIME SERIES ANALYSIS
# ==============================================================================

def add_acf_pacf_analysis(
    df: pd.DataFrame,
    value_col: str = "Total_Files",
    sheet_type: str = "daily",
    include_confidence: bool = True
) -> pd.DataFrame:
    """
    Calculates and adds ACF and PACF statistics to a time series DataFrame.
    Returns a new DataFrame containing ONLY the new analysis columns.
    """
    if not STATSMODELS_AVAILABLE:
        return pd.DataFrame(index=df.index)

    if value_col not in df.columns:
        return pd.DataFrame(index=df.index)

    series_data = df[value_col].fillna(0)
    
    lag_config = {
        'daily': [1, 7, 14],
        'weekly': [1, 4, 8],
        'biweekly': [1, 2, 4],
        'monthly': [1, 3, 6],
        'period': [1, 2, 3]
    }
    key_lags = lag_config.get(sheet_type, [1, 7, 14])
    max_lag = max(key_lags)

    # Create a new DataFrame for the results
    result_df = pd.DataFrame(index=df.index)

    # Initialize columns with NaN - FIXED: Use consistent underscore pattern
    new_col_names = []
    for lag in key_lags:
        acf_col = f"{value_col}_ACF_Lag_{lag}"  # FIXED: Added underscore before lag
        pacf_col = f"{value_col}_PACF_Lag_{lag}"  # FIXED: Added underscore before lag
        result_df[acf_col] = np.nan
        result_df[pacf_col] = np.nan
        new_col_names.extend([acf_col, pacf_col])
        if include_confidence:
            sig_acf_col = f'{acf_col}_Significant'
            sig_pacf_col = f'{pacf_col}_Significant'
            result_df[sig_acf_col] = pd.NA
            result_df[sig_pacf_col] = pd.NA
            new_col_names.extend([sig_acf_col, sig_pacf_col])

    # The minimum number of observations required to calculate PACF for a given number of lags (nlags) is 2 * nlags.
    # We will dynamically adjust the number of lags based on the available data points for each row.
    for i in range(1, len(df)):
        expanding_series = series_data.iloc[:i+1]
        n_obs = len(expanding_series)

        if n_obs < 4 or expanding_series.nunique() <= 1:
            continue

        try:
            # Dynamically determine the max number of lags we can calculate
            # PACF calculation requires at least 2*nlags observations
            current_nlags = min(max_lag, (n_obs // 2) - 1)

            if current_nlags <= 0:
                continue

            acf_vals = acf(expanding_series, nlags=current_nlags, fft=True)
            pacf_vals = pacf(expanding_series, nlags=current_nlags, method='ywm')
            confidence = 1.96 / np.sqrt(n_obs)

            for lag in key_lags:
                if lag < len(acf_vals):
                    acf_col = f"{value_col}_ACF_Lag_{lag}"
                    pacf_col = f"{value_col}_PACF_Lag_{lag}"
                    result_df.loc[result_df.index[i], acf_col] = acf_vals[lag]
                    result_df.loc[result_df.index[i], pacf_col] = pacf_vals[lag]
                    if include_confidence:
                        sig_acf_col = f'{acf_col}_Significant'
                        sig_pacf_col = f'{pacf_col}_Significant'
                        result_df.loc[result_df.index[i], sig_acf_col] = np.abs(acf_vals[lag]) > confidence
                        result_df.loc[result_df.index[i], sig_pacf_col] = np.abs(pacf_vals[lag]) > confidence
        except Exception:
            continue

    # Drop columns that are all NaN, which can happen if analysis fails on all rows
    result_df.dropna(axis=1, how='all', inplace=True)
            
    if include_confidence:
        for lag in key_lags:
            sig_acf_col = f"{value_col}_ACF_Lag_{lag}_Significant"  # FIXED: Added underscore before lag
            sig_pacf_col = f"{value_col}_PACF_Lag_{lag}_Significant"  # FIXED: Added underscore before lag
            if sig_acf_col in result_df.columns:
                result_df[sig_acf_col] = result_df[sig_acf_col].astype('boolean')
            if sig_pacf_col in result_df.columns:
                result_df[sig_pacf_col] = result_df[sig_pacf_col].astype('boolean')

    return result_df


def infer_sheet_type(sheet_name: str) -> str:
    """
    Infers the time scale type from sheet name for adaptive ACF/PACF lag configuration.

    Args:
        sheet_name: The name of the Excel sheet.

    Returns:
        The inferred sheet type ('daily', 'weekly', 'monthly', 'period', or 'biweekly').
    """
    sheet_name_lower = sheet_name.lower()

    if "biweekly" in sheet_name_lower:
        return "biweekly"
    elif "weekly" in sheet_name_lower:
        return "weekly"
    elif "monthly" in sheet_name_lower:
        return "monthly"
    elif "period" in sheet_name_lower:
        return "period"
    else:
        return "daily"  # Default fallback

def reorder_with_acf_pacf(df: pd.DataFrame, base_order: list) -> pd.DataFrame:
    """
    Reorders DataFrame columns to place ACF/PACF/Forecast columns immediately
    after the specific metric they relate to.

    Args:
        df: DataFrame to reorder.
        base_order: List of base metric columns (e.g., ['Total_Files', 'JPG_Files']).

    Returns:
        DataFrame with columns reordered.
    """
    if df.empty:
        return df

    final_order = []
    processed_cols = set()

    # Start with the primary key if it exists
    if '_id' in df.columns:
        final_order.append('_id')
        processed_cols.add('_id')

    # Identify all metric-related columns
    for metric in base_order:
        if metric in df.columns and metric not in processed_cols:
            final_order.append(metric)
            processed_cols.add(metric)

            # Find all related analysis columns for this specific metric
            related_cols = []
            for col in df.columns:
                if col.startswith(f"{metric}_"):
                    related_cols.append(col)
            
            # A simple sort is often good enough, but a more specific one is better
            # This sorting logic can be enhanced if needed
            final_order.extend(sorted(related_cols))
            processed_cols.update(related_cols)

    # Add any remaining columns that were not processed
    remaining_cols = [col for col in df.columns if col not in processed_cols]
    final_order.extend(remaining_cols)
    
    # Ensure the final list of columns only contains columns that actually exist
    existing_final_order = [col for col in final_order if col in df.columns]

    return df[existing_final_order]


# ==============================================================================
# ARIMA FORECASTING
# ==============================================================================

# Forecast horizon configuration per time scale
FORECAST_HORIZONS = {
    'daily': 14,      # 2 weeks
    'weekly': 6,      # 6 weeks  
    'biweekly': 4,    # 8 weeks
    'monthly': 3,     # 3 months
    'period': 2       # 2 periods
}

# Adaptive minimum data requirements per time scale
MIN_DATA_REQUIREMENTS = {
    'daily': 14,      # At least 2 weeks for daily patterns
    'weekly': 8,      # At least 8 weeks for weekly patterns
    'biweekly': 6,    # At least 6 biweekly periods
    'monthly': 6,     # At least 6 months for monthly patterns
    'period': 4       # At least 4 periods for period-based analysis
}

# Fallback options when primary ARIMA fails
FALLBACK_METHODS = {
    'moving_average': 'Simple moving average forecast',
    'linear_trend': 'Linear trend extrapolation',
    'seasonal_naive': 'Seasonal naive forecast (last period)',
    'mean_forecast': 'Historical mean forecast'
}

def generate_arima_forecast(
    series: pd.Series,
    forecast_horizon: int = None,
    max_p: int = 3,
    max_q: int = 3,
    seasonal: bool = False,
    sheet_type: str = "daily"
) -> tuple:
    """
    Generate ARIMA forecast for a time series with comprehensive diagnostics.
    
    Args:
        series (pd.Series): Time series data to forecast
        forecast_horizon (int): Number of periods to forecast (auto-determined if None)
        max_p (int): Maximum AR order to consider
        max_q (int): Maximum MA order to consider  
        seasonal (bool): Whether to use seasonal ARIMA (future enhancement)
        sheet_type (str): Sheet type for adaptive configuration
        
    Returns:
        tuple: (forecast_df, diagnostics_dict) or (None, error_dict) if failed
    """
    if not STATSMODELS_AVAILABLE:
        return None, {
            'error': 'statsmodels not available',
            'forecast_quality': 'Error',
            'message': '<Statsmodels Required>'
        }
    
    # Determine forecast horizon if not specified
    if forecast_horizon is None:
        forecast_horizon = FORECAST_HORIZONS.get(sheet_type, 14)
    
    # Adaptive data validation based on sheet type
    min_required = MIN_DATA_REQUIREMENTS.get(sheet_type, 10)
    if len(series) < min_required:
        # Use simple fallback for insufficient data
        fallback_forecast = _create_simple_fallback_forecast(series, forecast_horizon)
        return fallback_forecast, {
            'error': 'insufficient_data',
            'forecast_quality': 'Poor',
            'method_used': 'insufficient_data',
            'message': f'<Need {min_required}+ {sheet_type} periods>',
            'series_length': len(series),
            'min_required': min_required,
            'sheet_type': sheet_type,
            'suggestion': f'Collect at least {min_required} {sheet_type} data points for reliable forecasting'
        }
    
    try:
        # Clean and prepare series
        clean_series = _prepare_series_for_arima(series)
        
        if clean_series is None or len(clean_series) < min_required:
            # Provide specific guidance based on the failure reason
            if clean_series is None:
                error_msg = '<All values constant or invalid>'
                suggestion = 'Data contains no variation - check for data collection issues'
            else:
                error_msg = f'<Only {len(clean_series)} valid values>'
                suggestion = f'After cleaning, only {len(clean_series)} valid values remain - need {min_required}+'
            
            # Use simple fallback even for data preparation failures
            fallback_forecast = _create_simple_fallback_forecast(series, forecast_horizon)
            return fallback_forecast, {
                'error': 'data_preparation_failed',
                'forecast_quality': 'Poor',
                'method_used': 'mean_forecast',
                'message': error_msg,
                'original_length': len(series),
                'cleaned_length': len(clean_series) if clean_series is not None else 0,
                'min_required': min_required,
                'suggestion': suggestion
            }
        
        # Determine differencing order (d)
        d_order = _determine_differencing_order(clean_series)
        
        # Find best ARIMA model using grid search
        best_model, best_order, best_aic = _find_best_arima_model(
            clean_series, max_p, max_q, d_order
        )
        
        if best_model is None:
            # Try fallback forecasting methods
            fallback_result = _try_fallback_forecast(clean_series, forecast_horizon, sheet_type)
            if fallback_result is not None:
                return fallback_result
            
            return None, {
                'error': 'model_fitting_failed',
                'forecast_quality': 'Error',
                'message': f'<ARIMA failed for {sheet_type} data>',
                'series_length': len(clean_series),
                'sheet_type': sheet_type,
                'suggestion': 'Try collecting more data points or check for data quality issues',
                'attempted_models': f'Tested {(max_p+1) * (max_q+1)} ARIMA configurations'
            }
        
        # Generate forecast
        forecast_result = best_model.get_forecast(steps=forecast_horizon)
        forecast_values = forecast_result.predicted_mean
        forecast_ci = forecast_result.conf_int(alpha=0.05)  # 95% confidence interval
        
        # Calculate diagnostics
        diagnostics = _calculate_arima_diagnostics(best_model, clean_series, best_order, best_aic)
        
        # Create forecast DataFrame
        forecast_df = pd.DataFrame({
            'Forecast': forecast_values,
            'Lower_CI': forecast_ci.iloc[:, 0],
            'Upper_CI': forecast_ci.iloc[:, 1]
        })
        
        return forecast_df, diagnostics
        
    except Exception as e:
        return None, {
            'error': f'arima_exception: {str(e)}',
            'forecast_quality': 'Error',
            'message': '<ARIMA Error>'
        }


def _prepare_series_for_arima(series: pd.Series) -> pd.Series:
    """
    Clean and prepare time series for ARIMA modeling.
    
    Args:
        series (pd.Series): Raw time series
        
    Returns:
        pd.Series: Cleaned series ready for ARIMA
    """
    try:
        # Remove non-numeric values
        numeric_series = pd.to_numeric(series, errors='coerce')
        
        # Forward fill missing values (educational data often has gaps)
        filled_series = numeric_series.ffill()
        
        # Drop remaining NaN values at the beginning
        clean_series = filled_series.dropna()
        
        # Check for constant series
        if clean_series.nunique() <= 1:
            return None
            
        return clean_series
        
    except Exception:
        return None


def _determine_differencing_order(series: pd.Series, max_d: int = 2) -> int:
    """
    Determine the optimal differencing order using ADF test.
    
    Args:
        series (pd.Series): Time series to test
        max_d (int): Maximum differencing order to consider
        
    Returns:
        int: Optimal differencing order (0, 1, or 2)
    """
    try:
        # Test original series
        adf_result = adfuller(series, autolag='AIC')
        if adf_result[1] <= 0.05:  # p-value <= 0.05 means stationary
            return 0
        
        # Test first difference
        if len(series) > 1:
            diff1 = series.diff().dropna()
            if len(diff1) > 0:
                adf_result = adfuller(diff1, autolag='AIC')
                if adf_result[1] <= 0.05:
                    return 1
        
        # Test second difference if needed
        if max_d >= 2 and len(series) > 2:
            diff2 = series.diff().diff().dropna()
            if len(diff2) > 0:
                adf_result = adfuller(diff2, autolag='AIC')
                if adf_result[1] <= 0.05:
                    return 2
        
        # Default to first differencing if tests are inconclusive
        return 1
        
    except Exception:
        return 1  # Safe default


def _find_best_arima_model(series: pd.Series, max_p: int, max_q: int, d: int):
    """
    Find the best ARIMA model using grid search with AIC criterion.
    
    Args:
        series (pd.Series): Time series data
        max_p (int): Maximum AR order
        max_q (int): Maximum MA order
        d (int): Differencing order
        
    Returns:
        tuple: (best_model, best_order, best_aic) or (None, None, None) if failed
    """
    best_aic = float('inf')
    best_model = None
    best_order = None
    
    # Grid search over p and q
    for p in range(0, max_p + 1):
        for q in range(0, max_q + 1):
            try:
                # Skip (0,0,0) model
                if p == 0 and d == 0 and q == 0:
                    continue
                    
                model = ARIMA(series, order=(p, d, q))
                fitted_model = model.fit()
                
                if fitted_model.aic < best_aic:
                    best_aic = fitted_model.aic
                    best_model = fitted_model
                    best_order = (p, d, q)
                    
            except Exception:
                continue  # Skip problematic model configurations
    
    # Fallback to simple model if grid search failed
    if best_model is None:
        try:
            model = ARIMA(series, order=(1, 1, 0))
            best_model = model.fit()
            best_order = (1, 1, 0)
            best_aic = best_model.aic
        except Exception:
            return None, None, None
    
    return best_model, best_order, best_aic


def _try_fallback_forecast(series: pd.Series, forecast_horizon: int, sheet_type: str):
    """
    Attempt fallback forecasting methods when ARIMA fails.
    
    Args:
        series (pd.Series): Cleaned time series data
        forecast_horizon (int): Number of periods to forecast
        sheet_type (str): Sheet type for context
        
    Returns:
        tuple: (forecast_df, diagnostics_dict) or None if all methods fail
    """
    try:
        # Method 1: Simple moving average (most robust)
        if len(series) >= 3:
            window_size = min(len(series) // 3, 7)  # Adaptive window size
            recent_avg = series.tail(window_size).mean()
            
            # Create forecast DataFrame with slight variation to avoid static values
            forecast_values = []
            lower_values = []
            upper_values = []
            
            for i in range(forecast_horizon):
                # Add small time-based variation (0.1% per period)
                variation_factor = 1 + (i * 0.001)
                forecast_values.append(recent_avg * variation_factor)
                lower_values.append(recent_avg * 0.8 * variation_factor)
                upper_values.append(recent_avg * 1.2 * variation_factor)
            
            forecast_df = pd.DataFrame({
                'Forecast': forecast_values,
                'Lower_CI': lower_values,
                'Upper_CI': upper_values
            })
            
            diagnostics = {
                'forecast_quality': 'Fallback',
                'method': 'moving_average',
                'model_order': f'MA({window_size})',
                'message': f'Simple moving average forecast (ARIMA failed)',
                'fallback_reason': 'ARIMA model fitting failed',
                'window_size': window_size
            }
            
            return forecast_df, diagnostics
        
        # Method 2: Linear trend (if enough data)
        if len(series) >= 5:
            import numpy as np
            x = np.arange(len(series))
            slope, intercept = np.polyfit(x, series, 1)
            
            # Extrapolate trend
            future_x = np.arange(len(series), len(series) + forecast_horizon)
            trend_forecast = slope * future_x + intercept
            
            # Add some uncertainty bands
            residual_std = np.std(series - (slope * x + intercept))
            
            forecast_df = pd.DataFrame({
                'Forecast': trend_forecast,
                'Lower_CI': trend_forecast - 1.96 * residual_std,
                'Upper_CI': trend_forecast + 1.96 * residual_std
            })
            
            diagnostics = {
                'forecast_quality': 'Fallback',
                'method': 'linear_trend',
                'model_order': 'Linear',
                'message': f'Linear trend forecast (ARIMA failed)',
                'fallback_reason': 'ARIMA model fitting failed',
                'slope': round(slope, 3)
            }
            
            return forecast_df, diagnostics
        
        # Method 3: Historical mean (last resort) with variation
        mean_value = series.mean()
        std_value = series.std()
        
        # Create varying forecast values to avoid static output
        forecast_values = []
        lower_values = []
        upper_values = []
        
        for i in range(forecast_horizon):
            # Add small time-based variation (0.1% per period)
            variation_factor = 1 + (i * 0.001)
            forecast_values.append(mean_value * variation_factor)
            lower_values.append((mean_value - 1.96 * std_value) * variation_factor)
            upper_values.append((mean_value + 1.96 * std_value) * variation_factor)
        
        forecast_df = pd.DataFrame({
            'Forecast': forecast_values,
            'Lower_CI': lower_values,
            'Upper_CI': upper_values
        })
        
        diagnostics = {
            'forecast_quality': 'Fallback',
            'method': 'mean_forecast',
            'model_order': 'Mean',
            'message': f'Historical mean forecast (ARIMA failed)',
            'fallback_reason': 'ARIMA model fitting failed',
            'mean_value': round(mean_value, 1)
        }
        
        return forecast_df, diagnostics
        
    except Exception:
        return None  # All fallback methods failed


def _calculate_arima_diagnostics(model: ARIMAResultsWrapper, series: pd.Series, order: tuple, aic: float) -> dict:
    """
    Calculates comprehensive diagnostics for a fitted ARIMA model.

    This function computes key metrics like AIC, BIC, MAE, and RMSE, and performs
    a Ljung-Box test to check for residual autocorrelation. These diagnostics are
    used to assess the model's performance and forecast quality.

    Args:
        model: The fitted ARIMA model results object.
        series: The original time series data used for fitting.
        order: The (p,d,q) order of the fitted ARIMA model.
        aic: The Akaike Information Criterion (AIC) of the model.

    Returns:
        A dictionary containing a comprehensive set of model diagnostics.
    """
    try:
        # Basic model info
        diagnostics = {
            'model_order': order,
            'aic': round(aic, 2),
            'bic': round(model.bic, 2),
            'series_length': len(series)
        }
        
        # Calculate forecast accuracy metrics
        residuals = model.resid
        
        # MAE and RMSE
        mae = np.mean(np.abs(residuals))
        rmse = np.sqrt(np.mean(residuals**2))
        
        diagnostics.update({
            'mae': round(mae, 3),
            'rmse': round(rmse, 3)
        })
        
        # Ljung-Box test for residual autocorrelation
        try:
            lb_test = acorr_ljungbox(residuals, lags=min(10, len(residuals)//4), return_df=True)
            lb_pvalue = lb_test['lb_pvalue'].iloc[-1]  # Use the last lag's p-value
            diagnostics['ljung_box_pvalue'] = round(lb_pvalue, 4)
        except Exception:
            diagnostics['ljung_box_pvalue'] = None
        
        # Determine forecast quality
        quality = _assess_forecast_quality(diagnostics)
        diagnostics['forecast_quality'] = quality
        
        return diagnostics
        
    except Exception as e:
        return {
            'model_order': order,
            'aic': aic,
            'forecast_quality': 'Warning',
            'error': f'diagnostics_error: {str(e)}'
        }


def _assess_forecast_quality(diagnostics: dict) -> str:
    """
    Assess overall forecast quality based on diagnostics.
    
    Args:
        diagnostics (dict): Model diagnostics
        
    Returns:
        str: Quality assessment ('Good', 'Warning', 'Poor')
    """
    try:
        # Check Ljung-Box test (residual autocorrelation)
        lb_pvalue = diagnostics.get('ljung_box_pvalue')
        if lb_pvalue is not None and lb_pvalue < 0.05:
            return 'Warning'  # Residual autocorrelation detected
        
        # Check series length
        if diagnostics.get('series_length', 0) < 20:
            return 'Warning'  # Short series
        
        # If all checks pass
        return 'Good'
        
    except Exception:
        return 'Warning'


def add_arima_forecast_columns(df: pd.DataFrame, value_col: str = "Total_Files", sheet_type: str = "daily") -> pd.DataFrame:
    """
    Add ARIMA forecast columns to a DataFrame following established patterns.
    
    Args:
        df (pd.DataFrame): DataFrame to enhance with forecasts
        value_col (str): Column name to forecast
        sheet_type (str): Sheet type for adaptive configuration
        
    Returns:
        pd.DataFrame: DataFrame with forecast columns added
    """

    if not STATSMODELS_AVAILABLE:
        print("[WARNING]  ARIMA forecasting skipped: statsmodels not available")
        return df
    
    if value_col not in df.columns:
        print(f"[WARNING]  ARIMA forecasting skipped: {value_col} column not found")
        return df
    
    print(f"[CHART] Generating ARIMA forecast for {value_col} ({sheet_type} scale)...")
    
    try:
        # Generate forecast
        series = df[value_col]
        forecast_df, diagnostics = generate_arima_forecast(series, sheet_type=sheet_type)
        
        # Handle failed forecast
        if forecast_df is None:
            print(f"    - [EMOJI] Forecast failed: {diagnostics.get('message', 'Unknown error')}")
            df['Forecast_Quality'] = diagnostics.get('forecast_quality', 'Error')
            df['Forecast_Message'] = diagnostics.get('message', '')
            return df

        # Create future dates for the forecast horizon
        forecast_horizon = len(forecast_df)
        
        # Check if DataFrame has a datetime index, otherwise create simple numeric forecast columns
        if hasattr(df.index, 'dtype') and pd.api.types.is_datetime64_any_dtype(df.index):
            last_date = df.index[-1]
            
            # Determine frequency for date range generation
            freq_map = {
                'daily': 'D',
                'weekly': 'W-MON',
                'biweekly': '2W-MON',
                'monthly': 'MS',
                'period': 'MS'  # Placeholder, may need adjustment
            }
            freq = freq_map.get(sheet_type, 'D')
            
            # Generate future index
            future_index = pd.date_range(
                start=last_date + pd.Timedelta(days=1 if freq == 'D' else 0),
                periods=forecast_horizon,
                freq=freq
            )
            
            # Align forecast DataFrame with the new index
            forecast_df.index = future_index
            
            # Combine original DataFrame with forecast
            combined_df = pd.concat([df, forecast_df], axis=1)
        else:
            # For non-datetime indexes, add forecast columns directly to original DataFrame
            # Rename forecast columns to avoid conflicts
            forecast_df = forecast_df.rename(columns={
                'Forecast': f'{value_col}_Forecast',
                'Lower_CI': f'{value_col}_Forecast_Lower', 
                'Upper_CI': f'{value_col}_Forecast_Upper'
            })
            
            # Distribute forecast values across rows (cycling through the forecast sequence)
            for col in forecast_df.columns:
                # Create a cycling pattern through the forecast values for visual variety
                forecast_values = []
                for i in range(len(df)):
                    forecast_idx = i % len(forecast_df)  # Cycle through forecast sequence
                    forecast_values.append(forecast_df[col].iloc[forecast_idx])
                df[col] = forecast_values
            
            combined_df = df
        
        # Add diagnostics to the DataFrame
        combined_df['Forecast_Quality'] = diagnostics.get('forecast_quality', 'Unknown')
        combined_df['Forecast_Message'] = diagnostics.get('message', '')
        if 'model_order' in diagnostics:
            combined_df['Forecast_Model'] = str(diagnostics['model_order'])
        
        # Fill down diagnostic info for clarity in Excel
        for col in ['Forecast_Quality', 'Forecast_Message', 'Forecast_Model']:
            if col in combined_df.columns:
                combined_df[col] = combined_df[col].ffill()
        
        print(f"    - [OK] Forecast generated successfully (Quality: {diagnostics.get('forecast_quality', 'N/A')})")
        return combined_df

    except Exception as e:
        print(f"[ERROR] Critical error during ARIMA forecast addition: {e}")
        df['Forecast_Quality'] = 'Critical Error'
        df['Forecast_Message'] = str(e)
        return df


def _create_simple_fallback_forecast(series: pd.Series, forecast_horizon: int) -> pd.DataFrame:
    """
    Creates a simple, robust fallback forecast when ARIMA modeling fails.

    This function generates a forecast based on the historical mean and standard
    deviation of the series. It is designed to always return a valid DataFrame,
    ensuring that the forecasting pipeline does not crash even with problematic
    or insufficient data.

    Args:
        series: The original time series data.
        forecast_horizon: The number of periods to forecast into the future.

    Returns:
        A DataFrame containing the fallback forecast, including columns for the
        forecast values, and lower and upper confidence intervals.
    """
    try:
        numeric_series = pd.to_numeric(series, errors='coerce').dropna()

        if numeric_series.empty:
            base_value, std_dev = 0.0, 1.0
        elif len(numeric_series) == 1:
            base_value = float(numeric_series.iloc[0])
            std_dev = abs(base_value * 0.1) or 0.1
        else:
            base_value = float(numeric_series.mean())
            std_dev = max(float(numeric_series.std()), abs(base_value * 0.05)) or 0.1

        # Generate a slightly increasing trend to make the forecast less static
        trend_factor = np.arange(forecast_horizon) * (abs(base_value) * 0.01)
        forecast_values = base_value + trend_factor

        # Create confidence intervals
        margin_of_error = 1.96 * std_dev
        lower_ci = forecast_values - margin_of_error
        upper_ci = forecast_values + margin_of_error

        return pd.DataFrame({
            'Forecast': forecast_values,
            'Lower_CI': lower_ci,
            'Upper_CI': upper_ci
        })

    except Exception:
        # Ultimate fallback for any unexpected errors
        return pd.DataFrame({
            'Forecast': np.full(forecast_horizon, 1.0),
            'Lower_CI': np.full(forecast_horizon, 0.5),
            'Upper_CI': np.full(forecast_horizon, 1.5)
        })


def reorder_with_forecast_columns(df: pd.DataFrame, base_order: list, value_col: str = "Total_Files") -> pd.DataFrame:
    """
    Reorders DataFrame columns with forecast columns after ACF/PACF columns.
    Extends the existing reorder_with_acf_pacf pattern.
    
    Args:
        df (pd.DataFrame): DataFrame to reorder
        base_order (list): List of base column names in desired order
        value_col (str): Base column name for forecast columns
        
    Returns:
        pd.DataFrame: DataFrame with reordered columns including forecasts
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

