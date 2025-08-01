# ARIMA/SARIMAX Non-invertible/Non-stationary Warnings

## Problem Description

During the AR Data Analysis report generation process, warnings are encountered related to ARIMA/SARIMAX model fitting:

> ARIMA/SARIMAX warnings: Non-invertible/non-stationary starting parameters; optimizer used zeros.

These warnings indicate that the time series modeling process is encountering mathematical challenges that could affect forecast quality and reliability.

### Technical Background

1. **Non-invertible MA Parameters**:
   - In ARIMA models, the Moving Average (MA) component must be invertible for the model to be valid
   - Non-invertible MA processes cannot be represented as convergent infinite-order autoregressive processes
   - When MA parameters lie outside the invertibility region, forecasts may be unstable or unreliable

2. **Non-stationary Parameters**:
   - Stationarity is a key assumption in time series analysis
   - A stationary time series has constant statistical properties over time (mean, variance, autocorrelation)
   - Non-stationary processes can lead to spurious relationships and unreliable forecasts

3. **Optimizer Using Zeros**:
   - When the optimization algorithm encounters difficulties, it may resort to using zeros as parameter values
   - This is a fallback strategy by statsmodels when proper convergence cannot be achieved
   - Results in suboptimal model fits and potentially misleading forecasts

### Current Implementation

The current implementation in `ar_utils.py` uses a grid search approach to find the best ARIMA model:

```python
def _find_best_arima_model(series: pd.Series, max_p: int, max_q: int, d: int):
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
```

Key issues with this approach:
- No explicit stationarity testing before model fitting
- No constraints on parameter values during fitting
- Broad exception handling that silently skips problematic configurations
- Limited fallback strategy if all configurations fail

## Potential Solutions

### 1. Parameter Constraint Enforcement

Explicitly enforce stationarity and invertibility constraints during model fitting:

```python
model = ARIMA(
    series, 
    order=(p, d, q),
    enforce_stationarity=True,
    enforce_invertibility=True
)
fitted_model = model.fit()
```

**Pros**:
- Directly addresses the warning messages
- Ensures parameters remain within valid mathematical bounds
- Simple implementation with minimal code changes

**Cons**:
- May constrain the model too much for some time series
- Could potentially reduce fit quality for certain datasets
- Doesn't address potential underlying data issues

### 2. Pre-testing for Stationarity

Implement explicit stationarity testing before model fitting:

```python
from statsmodels.tsa.stattools import adfuller

def determine_differencing_order(series, max_d=2):
    """Determine appropriate differencing order using ADF test"""
    for d in range(max_d + 1):
        if d > 0:
            diff_series = series.diff(d).dropna()
        else:
            diff_series = series
            
        adf_result = adfuller(diff_series)
        p_value = adf_result[1]
        
        if p_value < 0.05:  # Stationary at 5% significance
            return d
            
    return max_d  # Default to max_d if still not stationary
```

**Pros**:
- Addresses the root cause of non-stationarity
- More scientific approach based on statistical testing
- Potentially improves forecast quality

**Cons**:
- Adds complexity to the modeling process
- May not address invertibility issues in MA components
- Requires careful handling of edge cases

### 3. Better Initial Parameter Selection

Use ACF/PACF analysis to guide parameter selection:

```python
from statsmodels.tsa.stattools import acf, pacf

def suggest_arima_orders(series, max_lag=10):
    """Suggest ARIMA orders based on ACF/PACF analysis"""
    diff_series = series.diff(1).dropna()  # First difference
    
    acf_values = acf(diff_series, nlags=max_lag)
    pacf_values = pacf(diff_series, nlags=max_lag)
    
    # Suggested p: Find significant PACF lags
    p_candidates = [i for i in range(1, len(pacf_values)) 
                   if abs(pacf_values[i]) > 1.96/np.sqrt(len(diff_series))]
    
    # Suggested q: Find significant ACF lags
    q_candidates = [i for i in range(1, len(acf_values)) 
                   if abs(acf_values[i]) > 1.96/np.sqrt(len(diff_series))]
    
    # Limit to top 3 candidates
    p_candidates = sorted(p_candidates)[:3] if p_candidates else [0, 1]
    q_candidates = sorted(q_candidates)[:3] if q_candidates else [0, 1]
    
    return p_candidates, q_candidates
```

**Pros**:
- More intelligent parameter selection based on data characteristics
- Potentially reduces the search space
- Aligns with statistical best practices

**Cons**:
- More complex implementation
- May still encounter invertibility issues
- Requires careful interpretation of ACF/PACF results

### 4. Warning Suppression

Use Python's warning handling to suppress specific warnings:

```python
import warnings

def fit_arima_with_suppressed_warnings(series, order):
    """Fit ARIMA model with suppressed warnings"""
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", 
                              message="Non-invertible starting MA parameters found.")
        warnings.filterwarnings("ignore", 
                              message="Non-stationary starting autoregressive parameters found.")
        
        model = ARIMA(series, order=order)
        return model.fit()
```

**Pros**:
- Simple implementation
- Cleans up output logs
- No changes to core modeling logic

**Cons**:
- Doesn't fix the underlying issues
- Masks potentially important warnings
- May hide problems that should be addressed

### 5. Enhanced Fallback Strategy

Implement a more sophisticated fallback strategy:

```python
def robust_arima_forecast(series, forecast_horizon):
    """Robust ARIMA forecasting with multiple fallback options"""
    # Try auto_arima from pmdarima if available
    try:
        import pmdarima as pm
        model = pm.auto_arima(
            series,
            start_p=0, max_p=5,
            start_q=0, max_q=5,
            d=None, max_d=2,
            seasonal=False,
            stepwise=True,
            suppress_warnings=True,
            error_action="ignore"
        )
        return model.predict(n_periods=forecast_horizon)
    except (ImportError, Exception):
        pass
    
    # Try standard ARIMA with common configurations
    for order in [(1,1,1), (0,1,1), (1,1,0), (2,1,2)]:
        try:
            model = ARIMA(series, order=order)
            fitted = model.fit()
            return fitted.forecast(forecast_horizon)
        except Exception:
            continue
    
    # Simple exponential smoothing fallback
    try:
        from statsmodels.tsa.holtwinters import SimpleExpSmoothing
        model = SimpleExpSmoothing(series).fit()
        return model.forecast(forecast_horizon)
    except Exception:
        pass
    
    # Last resort: naive forecast (repeat last value)
    return np.array([series.iloc[-1]] * forecast_horizon)
```

**Pros**:
- More robust forecasting approach
- Multiple layers of fallback options
- Better handling of difficult time series

**Cons**:
- Significantly more complex
- May introduce additional dependencies
- Could be slower due to multiple model attempts

## Experimental Scripting Guidelines

Before implementing changes to the main codebase, it's recommended to create experimental scripts to test solutions. Here's a suggested approach:

### 1. Create an Experimental Directory Structure

```
ARDataAnalysis_Refactored/
└── experiments/
    ├── arima_improvements/
    │   ├── data/
    │   ├── notebooks/
    │   ├── scripts/
    │   └── results/
    └── README.md
```

### 2. Data Collection Script

Create a script to extract relevant time series data from the MongoDB database:

```python
# experiments/arima_improvements/scripts/collect_data.py

import pymongo
import pandas as pd
import os
import sys
import json

# Add parent directory to path to import project modules
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..')))

from config import MONGODB_CONNECTION_STRING
from ar_utils import get_collection_periods

def extract_time_series_samples():
    """Extract sample time series from MongoDB for experimentation"""
    client = pymongo.MongoClient(MONGODB_CONNECTION_STRING)
    db = client.ar_data
    
    # Extract daily counts
    pipeline = [
        {"$group": {"_id": {"date": "$date"}, "count": {"$sum": 1}}},
        {"$sort": {"_id.date": 1}}
    ]
    
    daily_counts = list(db.files.aggregate(pipeline))
    daily_df = pd.DataFrame([
        {"date": doc["_id"]["date"], "count": doc["count"]} 
        for doc in daily_counts
    ])
    
    # Extract weekly counts
    # Similar aggregation for weekly data...
    
    # Save to CSV files
    os.makedirs("../data", exist_ok=True)
    daily_df.to_csv("../data/daily_counts_sample.csv", index=False)
    
    print(f"Extracted {len(daily_df)} daily records")
    
if __name__ == "__main__":
    extract_time_series_samples()
```

### 3. Experimentation Notebook

Create a Jupyter notebook to experiment with different solutions:

```python
# experiments/arima_improvements/notebooks/arima_solutions.ipynb

# Import necessary libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.stattools import adfuller, acf, pacf
import warnings
import sys
import os

# Load sample data
daily_df = pd.read_csv("../data/daily_counts_sample.csv")
daily_df["date"] = pd.to_datetime(daily_df["date"])
daily_df.set_index("date", inplace=True)

# Visualize the time series
plt.figure(figsize=(12, 6))
daily_df["count"].plot()
plt.title("Daily File Counts")
plt.grid(True)
plt.show()

# Test Solution 1: Parameter Constraint Enforcement
def test_parameter_constraints():
    # Original approach
    model_original = ARIMA(daily_df["count"], order=(2, 1, 2))
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        result_original = model_original.fit()
        warnings_original = len([warning for warning in w if "invertible" in str(warning)])
    
    # Constrained approach
    model_constrained = ARIMA(
        daily_df["count"], 
        order=(2, 1, 2),
        enforce_stationarity=True,
        enforce_invertibility=True
    )
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        result_constrained = model_constrained.fit()
        warnings_constrained = len([warning for warning in w if "invertible" in str(warning)])
    
    print(f"Original approach warnings: {warnings_original}")
    print(f"Constrained approach warnings: {warnings_constrained}")
    
    # Compare forecasts
    forecast_horizon = 14
    forecast_original = result_original.forecast(forecast_horizon)
    forecast_constrained = result_constrained.forecast(forecast_horizon)
    
    plt.figure(figsize=(12, 6))
    daily_df["count"][-30:].plot(label="Historical Data")
    forecast_original.plot(label="Original Forecast")
    forecast_constrained.plot(label="Constrained Forecast")
    plt.title("Forecast Comparison")
    plt.legend()
    plt.grid(True)
    plt.show()
    
    return result_original, result_constrained

# Test other solutions similarly...

# Comprehensive evaluation function
def evaluate_all_solutions():
    # Implementation of all solutions and comparison
    pass

# Run experiments
original_model, constrained_model = test_parameter_constraints()
# Run other solution tests...
```

### 4. Benchmark Script

Create a script to benchmark different solutions:

```python
# experiments/arima_improvements/scripts/benchmark.py

import pandas as pd
import numpy as np
import time
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
import warnings
import sys
import os
import json

# Add parent directory to path to import project modules
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..')))

from ar_utils import _find_best_arima_model

# Load sample data
daily_df = pd.read_csv("../data/daily_counts_sample.csv")
daily_df["date"] = pd.to_datetime(daily_df["date"])
daily_df.set_index("date", inplace=True)

# Original implementation
def original_implementation(series):
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        start_time = time.time()
        model, order, aic = _find_best_arima_model(series, max_p=2, max_q=2, d=1)
        duration = time.time() - start_time
        warning_count = len([warning for warning in w if "invertible" in str(warning)])
    
    return {
        "method": "original",
        "order": order,
        "aic": aic,
        "duration": duration,
        "warnings": warning_count
    }

# Solution 1: Parameter constraints
def solution_1(series):
    # Implementation with parameter constraints
    pass

# Solution 2: Stationarity testing
def solution_2(series):
    # Implementation with stationarity testing
    pass

# Run benchmarks
results = []
results.append(original_implementation(daily_df["count"]))
# Add other solution benchmarks

# Save results
with open("../results/benchmark_results.json", "w") as f:
    json.dump(results, f, indent=2)

# Generate comparison chart
# ...
```

### 5. Integration Test Script

Create a script to test integration with the main codebase:

```python
# experiments/arima_improvements/scripts/integration_test.py

import sys
import os
import pandas as pd
import warnings

# Add parent directory to path to import project modules
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..')))

# Import the original module
import ar_utils

# Create a modified version of the module for testing
def create_modified_module():
    # Read the original module
    with open("../../../ar_utils.py", "r") as f:
        code = f.read()
    
    # Apply modifications based on the selected solution
    # For example, adding parameter constraints to _find_best_arima_model
    modified_code = code.replace(
        "model = ARIMA(series, order=(p, d, q))",
        "model = ARIMA(series, order=(p, d, q), enforce_stationarity=True, enforce_invertibility=True)"
    )
    
    # Write to a temporary file
    with open("../temp/modified_ar_utils.py", "w") as f:
        f.write(modified_code)

# Test with sample data
def test_integration():
    # Create sample DataFrame
    df = pd.read_csv("../data/daily_counts_sample.csv")
    df["date"] = pd.to_datetime(df["date"])
    
    # Test original implementation
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        result_original = ar_utils.add_arima_forecast_columns(
            df.copy(), value_col="count", sheet_type="daily"
        )
        warnings_original = len([warning for warning in w if "invertible" in str(warning)])
    
    # Import modified module
    sys.path.insert(0, os.path.abspath("../temp"))
    import modified_ar_utils
    
    # Test modified implementation
    with warnings.catch_warnings(record=True) as w:
        warnings.simplefilter("always")
        result_modified = modified_ar_utils.add_arima_forecast_columns(
            df.copy(), value_col="count", sheet_type="daily"
        )
        warnings_modified = len([warning for warning in w if "invertible" in str(warning)])
    
    print(f"Original implementation warnings: {warnings_original}")
    print(f"Modified implementation warnings: {warnings_modified}")
    
    # Compare results
    # ...

if __name__ == "__main__":
    os.makedirs("../temp", exist_ok=True)
    create_modified_module()
    test_integration()
```

## Recommended Approach

Based on the analysis of the problem and potential solutions, the recommended approach is a combination of:

1. **Parameter Constraint Enforcement** (primary solution):
   - Add `enforce_stationarity=True` and `enforce_invertibility=True` to ARIMA model creation
   - This directly addresses the warning messages with minimal code changes

2. **Pre-testing for Stationarity** (secondary enhancement):
   - Add stationarity testing to determine appropriate differencing order
   - This improves model quality and reduces the likelihood of non-stationary warnings

3. **Enhanced Error Handling** (robustness improvement):
   - Improve exception handling to provide more specific error messages
   - Add logging of model fitting issues for better diagnostics

This combined approach addresses the immediate warning issues while also improving the overall robustness and quality of the ARIMA forecasting system.

## Implementation Plan

1. Create experimental scripts to validate solutions
2. Benchmark different approaches using real data
3. Select the best performing solution
4. Implement changes in a feature branch
5. Add comprehensive tests
6. Update documentation
7. Submit for code review

By following this process, we can ensure that the solution is thoroughly tested before being integrated into the main codebase.
