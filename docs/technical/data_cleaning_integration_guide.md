# Data Cleaning Utils Integration Guide

## Overview

This guide documents how to integrate the newly created `DataCleaningUtils` class with the existing AR Data Analysis codebase. The integration follows a safe, incremental approach to minimize risk and ensure continuous functionality.

## Background

The AR Data Analysis system uses a 2×2 matrix filtering approach based on:
- `is_collection_day` (TRUE/FALSE): Whether a file was recorded on a designated collection day
- `Outlier_Status` (TRUE/FALSE): Whether a file is flagged as an outlier

This creates four mutually exclusive categories:
1. School Normal (final research dataset): `is_collection_day=TRUE & Outlier_Status=FALSE`
2. School Outliers: `is_collection_day=TRUE & Outlier_Status=TRUE`
3. Non-School Normal: `is_collection_day=FALSE & Outlier_Status=FALSE`
4. Non-School Outliers: `is_collection_day=FALSE & Outlier_Status=TRUE`

The data cleaning logic was previously embedded in `SheetCreator.create_data_cleaning_sheet()`. We've now extracted this logic into a reusable `DataCleaningUtils` class.

## Implementation Status

- ✅ Created `DataCleaningUtils` class in `utils/data_cleaning.py`
- ✅ Implemented test script `test_data_cleaning_utils.py` for validation
- ✅ Created proof of concept integration in `data_cleaning_integration.py`
- ⬜ Integrate with `SheetCreator.create_data_cleaning_sheet()`
- ⬜ Extend integration to other sheets

## Integration Steps

### Step 1: Import and Initialize DataCleaningUtils

In `report_generator/sheet_creators/__init__.py`, add the import at the top of the file:

```python
from utils.data_cleaning import DataCleaningUtils
```

Then in the `__init__` method of the `SheetCreator` class, initialize the utils:

```python
def __init__(self, db, formatter):
    self.db = db
    self.formatter = formatter
    # Initialize data cleaning utils
    self.data_cleaning_utils = DataCleaningUtils(db)
```

### Step 2: Refactor create_data_cleaning_sheet Method (Safe Approach)

The safest approach is to add a wrapper method first that uses the new utility class:

```python
def create_data_cleaning_sheet_with_utils(self, workbook):
    """
    Creates the Data Cleaning sheet using the DataCleaningUtils class.
    This is a wrapper for the original create_data_cleaning_sheet method.
    """
    # Use the utils class to get data
    result = self.data_cleaning_utils.get_complete_cleaning_data()
    intersection_data = result['intersection_data']
    totals = result['totals']
    
    # Use the existing method for formatting and report creation
    # but pass the pre-calculated data
    return self._create_data_cleaning_sheet_impl(workbook, intersection_data, totals)
```

Then, extract the implementation from the original method:

```python
def _create_data_cleaning_sheet_impl(self, workbook, intersection_data=None, totals=None):
    """
    Implementation of the Data Cleaning sheet creation.
    Can either calculate data (original behavior) or use pre-calculated data.
    """
    # Create worksheet
    ws = workbook.create_sheet("Data Cleaning")
    
    # Rest of the original method, but check if data is provided
    if intersection_data is None or totals is None:
        # Original data calculation logic
        # [...]
    
    # Rest of the original method using the data
    # [...]
```

### Step 3: Add Year Breakdown Integration

For the year-by-year breakdown, add a similar approach:

```python
def _create_year_breakdown_table(self, ws, start_row, use_utils=False):
    """
    Creates the year breakdown table.
    Can either use the original logic or the utility class.
    """
    if use_utils:
        year_data = self.data_cleaning_utils.get_year_breakdown_data()
        # Format and add to worksheet
        # [...]
    else:
        # Original implementation
        # [...]
```

### Step 4: Testing Strategy

To ensure the integration works correctly:

1. Create a test script that generates reports using both the original and new methods
2. Compare the outputs cell by cell to verify they're identical
3. Start by comparing just key metrics (totals, percentages)
4. Only proceed to full integration after confirming results match

### Step 5: Full Integration (After Validation)

Once you've verified that both approaches produce identical results:

1. Update the original `create_data_cleaning_sheet` to use the utility class
2. Keep the implementation method as a fallback option
3. Add tests to verify the results match expectations

```python
def create_data_cleaning_sheet(self, workbook):
    """
    Creates the Data Cleaning sheet.
    Uses the DataCleaningUtils class for data calculation.
    """
    try:
        # Use the utils class for data calculation
        result = self.data_cleaning_utils.get_complete_cleaning_data()
        intersection_data = result['intersection_data']
        totals = result['totals']
        
        # Use the implementation method for formatting
        return self._create_data_cleaning_sheet_impl(workbook, intersection_data, totals)
    except Exception as e:
        # Log the error
        print(f"[ERROR] Failed to use DataCleaningUtils: {e}")
        print("[INFO] Falling back to original implementation")
        # Fall back to original implementation
        return self._create_data_cleaning_sheet_impl(workbook)
```

### Step 6: Extension to Other Sheets

After successful integration with the data cleaning sheet:

1. Identify other sheets or reports that use similar data cleaning logic
2. Create utility methods in DataCleaningUtils for their specific needs
3. Integrate with those sheets using the same incremental approach
4. Verify results at each step

## Integration Verification Checklist

For each integration point:

- [ ] Numeric values match the original implementation
- [ ] Percentages are calculated correctly and sum to 100%
- [ ] Year breakdown data matches the complete analysis
- [ ] Table formatting and layout are preserved
- [ ] All academic terminology is consistent
- [ ] Explanatory notes and context are preserved
- [ ] No regression in other functionality

## Conclusion

This integration approach minimizes risk by:
1. Keeping the original code intact during testing
2. Adding new functionality in parallel
3. Providing fallback mechanisms
4. Verifying results at each step
5. Making small, incremental changes

By following this guide, the data cleaning logic can be safely modularized while ensuring consistent results and preserving the academic integrity of the AR Data Analysis system.
