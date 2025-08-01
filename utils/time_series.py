"""
Time Series Analysis Utilities for AR Data Analysis

This module provides utilities for time series analysis including ACF/PACF calculations,
ARIMA forecasting, and related statistical analysis functions.

Key Functions:
- ACF/PACF analysis with adaptive lag configuration
- ARIMA forecasting with model selection
- Time series chart generation
- Statistical validation and diagnostics
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Optional, Tuple, List

# Import statsmodels for ACF/PACF analysis and ARIMA forecasting
try:
    from statsmodels.tsa.stattools import acf, pacf, adfuller
    from statsmodels.tsa.arima.model import ARIMA, ARIMAResultsWrapper
    from statsmodels.stats.diagnostic import acorr_ljungbox
    STATSMODELS_AVAILABLE = True
except ImportError:
    STATSMODELS_AVAILABLE = False
    print("[WARNING] Warning: statsmodels not available. ACF/PACF and ARIMA analysis will be disabled.")
    print("    Install with: pip install statsmodels")


def add_acf_pacf_analysis(
    df: pd.DataFrame,
    value_col: str = "Total_Files",
    sheet_type: str = "daily",
    include_confidence: bool = True
) -> pd.DataFrame:
    """
    Calculates and adds ACF and PACF statistics to a time series DataFrame.

    This function computes the Autocorrelation Function (ACF) and Partial
    Autocorrelation Function (PACF) for a specified column in the DataFrame.
    It adapts the number of lags based on the time scale and data availability.

    Args:
        df: DataFrame containing the time series data.
        value_col: Name of the column to analyze (default: "Total_Files").
        sheet_type: Time scale type for adaptive lag configuration.
        include_confidence: Whether to include confidence interval significance.

    Returns:
        DataFrame with additional ACF and PACF columns added, along with
        confidence interval significance.
    """
    if not STATSMODELS_AVAILABLE:
        print("[WARNING] Statsmodels not available. Skipping ACF/PACF analysis.")
        return df

    if value_col not in df.columns:
        print(f"[WARNING] Column '{value_col}' not found. Skipping ACF/PACF analysis.")
        return df

    # Check if ACF/PACF columns already exist to prevent duplication
    # Robust check for both naming conventions: ACF_Lag_ and ACF_Lag
    existing_acf_cols = []
    for col in df.columns:
        col_str = str(col)
        # Check for both underscore patterns: ACF_Lag_ and ACF_Lag
        if (f'{value_col}_ACF_Lag_' in col_str or 
            (f'{value_col}_ACF_Lag' in col_str and f'{value_col}_ACF_Lag_' not in col_str)):
            existing_acf_cols.append(col)
    if existing_acf_cols:
        print(f"[INFO] ACF/PACF columns already exist ({len(existing_acf_cols)} found). Skipping duplicate analysis.")
        print(f"[DEBUG] Existing ACF columns: {existing_acf_cols}")
        print(f"[DEBUG] All columns: {list(df.columns)}")
        return df

    # Create a copy to avoid modifying the original DataFrame
    df_with_acf_pacf = df.copy()

    try:
        # Convert the target column to numeric, handling any non-numeric values
        series = pd.to_numeric(df_with_acf_pacf[value_col], errors='coerce').dropna()

        if len(series) < 10:
            print(f"[WARNING] Insufficient data for ACF/PACF analysis (n={len(series)}). Minimum 10 observations required.")
            return df

        # Adaptive lag configuration based on sheet type
        lag_configs = {
            'daily': [1, 7, 14],      # 1 day, 1 week, 2 weeks
            'weekly': [1, 4, 8],      # 1 week, 1 month, 2 months  
            'biweekly': [1, 2, 4],    # 1 period, 2 periods, 4 periods
            'monthly': [1, 3, 6],     # 1 month, 1 quarter, 6 months
            'period': [1, 2, 4]       # 1 period, 2 periods, 4 periods
        }

        lags_to_compute = lag_configs.get(sheet_type, [1, 2, 3])

        # Compute ACF and PACF
        max_lag = min(max(lags_to_compute), len(series) // 2 - 1)
        if max_lag < 1:
            print(f"[WARNING] Series too short for meaningful ACF/PACF analysis (n={len(series)}).")
            return df

        # Calculate ACF and PACF
        acf_values, acf_confint = acf(series, nlags=max_lag, alpha=0.05, fft=False)
        
        # Dynamic PACF lag capping to prevent mathematical errors
        max_valid_pacf_lag = len(series) // 2 - 1
        pacf_max_lag = min(max_lag, max_valid_pacf_lag)
        
        if pacf_max_lag >= 1:
            pacf_values, pacf_confint = pacf(series, nlags=pacf_max_lag, alpha=0.05)
        else:
            pacf_values = np.array([1.0])  # PACF at lag 0 is always 1
            pacf_confint = np.array([[0.95, 1.05]])

        # Add ACF and PACF columns for the specified lags
        for lag in lags_to_compute:
            if lag <= max_lag:
                # ACF values
                df_with_acf_pacf[f'{value_col}_ACF_Lag_{lag}'] = round(acf_values[lag], 3)
                
                # PACF values (with dynamic capping)
                if lag <= pacf_max_lag:
                    df_with_acf_pacf[f'{value_col}_PACF_Lag_{lag}'] = round(pacf_values[lag], 3)
                else:
                    df_with_acf_pacf[f'{value_col}_PACF_Lag_{lag}'] = "<Exceeded Max Lag>"

                # Confidence interval significance (if requested)
                if include_confidence and lag <= max_lag:
                    acf_lower, acf_upper = acf_confint[lag]
                    is_significant = abs(acf_values[lag]) > max(abs(acf_lower), abs(acf_upper))
                    df_with_acf_pacf[f'{value_col}_ACF_Lag_{lag}_Significant'] = is_significant
            else:
                # Handle lags beyond available data
                df_with_acf_pacf[f'{value_col}_ACF_Lag_{lag}'] = "<Insufficient Data>"
                df_with_acf_pacf[f'{value_col}_PACF_Lag_{lag}'] = "<Insufficient Data>"
                if include_confidence:
                    df_with_acf_pacf[f'{value_col}_ACF_Lag_{lag}_Significant'] = "<N/A>"

        print(f"    - [OK] ACF/PACF analysis completed for {sheet_type} data (n={len(series)}, max_lag={max_lag})")
        print(f"    - [DEBUG] Added columns: {[col for col in df_with_acf_pacf.columns if col not in df.columns]}")
        print(f"    - [DEBUG] Final column count: {len(df_with_acf_pacf.columns)}")

    except Exception as e:
        print(f"[ERROR] Error in ACF/PACF analysis: {e}")
        # Add error indicators
        for lag in [1, 7, 14]:  # Default lags
            df_with_acf_pacf[f'{value_col}_ACF_Lag_{lag}'] = "<Error>"
            df_with_acf_pacf[f'{value_col}_PACF_Lag_{lag}'] = "<Error>"

    return df_with_acf_pacf


def infer_sheet_type(sheet_name: str) -> str:
    """
    Infers the time scale type from sheet name for adaptive ACF/PACF lag configuration.

    Args:
        sheet_name: The name of the Excel sheet.

    Returns:
        The inferred sheet type ('daily', 'weekly', 'monthly', 'period', or 'biweekly').
    """
    sheet_lower = sheet_name.lower()
    
    if 'daily' in sheet_lower:
        return 'daily'
    elif 'weekly' in sheet_lower:
        return 'weekly'
    elif 'biweekly' in sheet_lower:
        return 'biweekly'
    elif 'monthly' in sheet_lower:
        return 'monthly'
    elif 'period' in sheet_lower:
        return 'period'
    else:
        return 'daily'  # Default fallback


def generate_arima_forecast(
    series: pd.Series,
    forecast_horizon: int = 6,
    value_col: str = "Total_Files"
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Generates ARIMA forecast for a time series with comprehensive error handling.

    This function performs automatic ARIMA model selection using grid search,
    generates forecasts with confidence intervals, and provides detailed
    diagnostics about the forecasting process.

    Args:
        series: The time series data to forecast.
        forecast_horizon: Number of periods to forecast ahead.
        value_col: Name of the value column for labeling forecast columns.

    Returns:
        Tuple containing:
        - DataFrame with forecast values and confidence intervals
        - Dictionary with diagnostic information about the forecasting process
    """
    if not STATSMODELS_AVAILABLE:
        return _create_simple_fallback_forecast(series, forecast_horizon), {
            'forecast_quality': 'Unavailable',
            'forecast_message': 'Statsmodels not available',
            'model_order': 'N/A'
        }

    try:
        # Convert to numeric and clean the series
        numeric_series = pd.to_numeric(series, errors='coerce').dropna()

        if len(numeric_series) < 10:
            return _create_simple_fallback_forecast(series, forecast_horizon), {
                'forecast_quality': 'Insufficient Data',
                'forecast_message': f'Only {len(numeric_series)} valid observations',
                'model_order': 'N/A'
            }

        # Test for stationarity
        try:
            adf_result = adfuller(numeric_series)
            is_stationary = adf_result[1] <= 0.05
        except:
            is_stationary = False

        # Grid search for optimal ARIMA parameters
        best_aic = float('inf')
        best_order = None
        best_model = None

        # Limited grid search for efficiency
        p_values = range(0, 3)
        d_values = [0, 1] if is_stationary else [1, 2]
        q_values = range(0, 3)

        for p in p_values:
            for d in d_values:
                for q in q_values:
                    try:
                        model = ARIMA(numeric_series, order=(p, d, q))
                        fitted_model = model.fit()
                        
                        if fitted_model.aic < best_aic:
                            best_aic = fitted_model.aic
                            best_order = (p, d, q)
                            best_model = fitted_model
                    except:
                        continue

        if best_model is None:
            return _create_simple_fallback_forecast(series, forecast_horizon), {
                'forecast_quality': 'Model Fitting Failed',
                'forecast_message': 'Unable to fit any ARIMA model',
                'model_order': 'N/A'
            }

        # Generate forecast
        forecast_result = best_model.forecast(steps=forecast_horizon, alpha=0.05)
        forecast_values = forecast_result
        
        # Get confidence intervals
        forecast_ci = best_model.get_forecast(steps=forecast_horizon, alpha=0.05).conf_int()
        lower_ci = forecast_ci.iloc[:, 0].values
        upper_ci = forecast_ci.iloc[:, 1].values

        # Create forecast DataFrame
        forecast_df = pd.DataFrame({
            f'{value_col}_Forecast': forecast_values,
            f'{value_col}_Forecast_Lower': lower_ci,
            f'{value_col}_Forecast_Upper': upper_ci
        })

        # Diagnostic information
        diagnostics = {
            'forecast_quality': 'Good',
            'forecast_message': f'ARIMA{best_order} model fitted successfully',
            'model_order': best_order,
            'aic': best_aic,
            'is_stationary': is_stationary
        }

        return forecast_df, diagnostics

    except Exception as e:
        return _create_simple_fallback_forecast(series, forecast_horizon), {
            'forecast_quality': 'Error',
            'forecast_message': str(e),
            'model_order': 'N/A'
        }


def add_arima_forecast_to_dataframe(
    df: pd.DataFrame,
    value_col: str = "Total_Files",
    sheet_type: str = "daily"
) -> pd.DataFrame:
    """
    Adds ARIMA forecast columns to an existing DataFrame.

    Args:
        df: DataFrame containing the time series data
        value_col: Name of the column to forecast
        sheet_type: Time scale type for adaptive forecast horizon

    Returns:
        DataFrame with forecast columns added
    """
    if value_col not in df.columns:
        print(f"[WARNING] Column '{value_col}' not found. Skipping ARIMA forecast.")
        return df

    try:
        # Adaptive forecast horizon based on sheet type
        horizon_configs = {
            'daily': 14,      # 2 weeks
            'weekly': 6,      # 6 weeks
            'biweekly': 4,    # 4 periods
            'monthly': 3,     # 3 months
            'period': 2       # 2 periods
        }

        forecast_horizon = horizon_configs.get(sheet_type, 6)
        
        # Extract the series for forecasting
        series = df[value_col]
        
        # Generate forecast
        forecast_df, diagnostics = generate_arima_forecast(
            series, forecast_horizon, value_col
        )

        # Combine original data with forecast
        combined_df = df.copy()

        # Add forecast columns to all rows (will be filled with forecast values at the end)
        for col in forecast_df.columns:
            combined_df[col] = '<N/A>'

        # Add the actual forecast values to the last few rows
        if len(forecast_df) > 0:
            start_idx = max(0, len(combined_df) - len(forecast_df))
            for i, (_, forecast_row) in enumerate(forecast_df.iterrows()):
                if start_idx + i < len(combined_df):
                    for col in forecast_df.columns:
                        combined_df.iloc[start_idx + i, combined_df.columns.get_loc(col)] = forecast_row[col]

        # Add diagnostic information
        if 'forecast_quality' in diagnostics:
            combined_df['Forecast_Quality'] = diagnostics['forecast_quality']
        if 'forecast_message' in diagnostics:
            combined_df['Forecast_Message'] = diagnostics['forecast_message']
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
