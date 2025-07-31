# Quick Start Guide for ChatGPT Code Review

## 🚨 THE PROBLEM
**4x Column Duplication in ACF/PACF Excel Sheets**
- Expected: ~15 columns per sheet
- Actual: ~50 columns per sheet
- Pipeline caching fixes implemented but duplication persists

## 🎯 WHAT TO ANALYZE

### 1. Primary Files to Review
```
ar_utils.py                           # ACF/PACF analysis functions
sheet_creators/pipeline.py            # Sheet creation logic  
report_generator/core.py              # Main orchestration
sheet_creators/base.py                # Base sheet creator (recently fixed)
report_generator/dashboard.py         # Dashboard creator (recently fixed)
```

### 2. Key Functions to Examine
```python
# In ar_utils.py
add_acf_pacf_analysis()              # Adds ACF/PACF columns
reorder_with_acf_pacf()              # Reorders columns

# In sheet_creators/pipeline.py  
_create_pipeline_sheet()             # Creates individual sheets

# In report_generator/core.py
generate_report()                    # Main orchestration method
```

## 🔍 DEBUGGING CLUES

### What's Working ✅
- Pipeline caching prevents duplicate executions
- Individual ACF/PACF functions work correctly in isolation
- Cache hit/miss messages confirm caching is functional

### What's Broken ❌
- Full workflow produces 4x column duplication
- Issue occurs only in orchestrated report generation
- Duplication pattern suggests ACF/PACF columns added 4 times

## 🤔 INVESTIGATION QUESTIONS

1. **DataFrame Mutations**: Are DataFrames being shared/mutated across sheets?
2. **Hidden Loops**: Is there orchestration logic calling ACF/PACF analysis multiple times?
3. **Reference Issues**: Are DataFrame references being passed around and modified?
4. **Column Accumulation**: Where exactly are columns being accumulated?

## 📊 EXPECTED BEHAVIOR
```python
# Normal ACF/PACF sheet should have:
base_columns = ['Date', 'Total_Files', 'MP3_Files', 'JPG_Files', ...]  # ~7-8 cols
acf_columns = ['Total_Files_ACF_Lag_1', 'Total_Files_ACF_Lag_2', ...]   # ~5 cols  
pacf_columns = ['Total_Files_PACF_Lag_1', 'Total_Files_PACF_Lag_2', ...] # ~5 cols
# Total: ~15-18 columns
```

## 🚀 SUGGESTED APPROACH

1. **Trace ACF/PACF column addition** in `add_acf_pacf_analysis()`
2. **Check DataFrame handling** in `_create_pipeline_sheet()`
3. **Look for hidden orchestration loops** in `generate_report()`
4. **Identify shared DataFrame references** across the codebase

## 💡 RECENT FIXES (For Context)
- Added `_run_aggregation_cached()` to prevent duplicate pipeline executions
- Implemented pipeline caching in `base.py` and `dashboard.py`
- Each sheet now fetches fresh pipeline data

**The mystery**: Caching works, no duplicate executions, but 4x duplication persists!
