#!/usr/bin/env python3
"""
Advanced DataFrame Tracking Debug Script
========================================

This script implements comprehensive DataFrame tracking to identify the exact
source of 4x column duplication by monitoring DataFrame IDs, memory addresses,
and operation sequences.
"""

import pandas as pd
import sys
import json
import traceback
from pathlib import Path
from collections import defaultdict

# Add the project root to the path
sys.path.insert(0, str(Path(__file__).parent))

# Global tracking
dataframe_operations = defaultdict(list)
dataframe_genealogy = {}

def track_dataframe_operation(df, operation_name, source_location=None):
    """Track every operation performed on DataFrames"""
    if not hasattr(df, 'columns'):
        return df
    
    df_id = id(df)
    df_shape = (len(df), len(df.columns))
    
    # Get call stack for source tracking
    if source_location is None:
        stack = traceback.extract_stack()
        source_location = f"{Path(stack[-3].filename).name}:{stack[-3].lineno}"
    
    # Check for column duplicates
    duplicate_cols = [col for col in df.columns if list(df.columns).count(col) > 1]
    
    operation_info = {
        'operation': operation_name,
        'df_id': df_id,
        'shape': df_shape,
        'source': source_location,
        'has_duplicates': bool(duplicate_cols),
        'duplicate_columns': list(set(duplicate_cols)) if duplicate_cols else [],
        'total_columns': len(df.columns),
        'unique_columns': len(set(df.columns))
    }
    
    dataframe_operations[df_id].append(operation_info)
    
    print(f"🔍 {operation_name} | ID:{df_id} | Shape:{df_shape} | Dupes:{len(duplicate_cols)} | {source_location}")
    
    if duplicate_cols:
        from collections import Counter
        col_counts = Counter(df.columns)
        print(f"   ❌ DUPLICATES: {[(col, count) for col, count in col_counts.items() if count > 1]}")
    
    return df

def monkey_patch_dataframe_operations():
    """Patch all major DataFrame operations for tracking"""
    
    # Patch pandas concat
    original_concat = pd.concat
    def tracked_concat(*args, **kwargs):
        print(f"\n🔗 pd.concat called from {traceback.extract_stack()[-2].filename}:{traceback.extract_stack()[-2].lineno}")
        
        if args and hasattr(args[0], '__iter__'):
            dfs = list(args[0]) if not isinstance(args[0], pd.DataFrame) else [args[0]]
            print(f"   Input DataFrames: {len(dfs)}")
            for i, df in enumerate(dfs):
                if hasattr(df, 'columns'):
                    track_dataframe_operation(df, f"concat_input_{i}")
        
        result = original_concat(*args, **kwargs)
        track_dataframe_operation(result, "concat_result")
        return result
    
    pd.concat = tracked_concat
    
    # Patch DataFrame copy
    original_copy = pd.DataFrame.copy
    def tracked_copy(self, *args, **kwargs):
        result = original_copy(self, *args, **kwargs)
        track_dataframe_operation(self, "copy_source")
        track_dataframe_operation(result, "copy_result")
        dataframe_genealogy[id(result)] = id(self)
        return result
    
    pd.DataFrame.copy = tracked_copy
    
    # Patch DataFrame assignment
    original_setitem = pd.DataFrame.__setitem__
    def tracked_setitem(self, key, value):
        track_dataframe_operation(self, f"setitem_{key}_before")
        result = original_setitem(self, key, value)
        track_dataframe_operation(self, f"setitem_{key}_after")
        return result
    
    pd.DataFrame.__setitem__ = tracked_setitem

def patch_ar_utils_functions():
    """Patch ar_utils functions for detailed tracking"""
    
    import ar_utils
    
    # Patch add_acf_pacf_analysis
    original_add_acf_pacf = ar_utils.add_acf_pacf_analysis
    def tracked_add_acf_pacf(df, value_col='Total_Files', sheet_type='daily', include_confidence=True):
        print(f"\n📊 ACF/PACF Analysis Starting")
        track_dataframe_operation(df, "acf_pacf_input")
        
        # Check if ACF columns already exist (replicate original logic)
        existing_acf_cols = [col for col in df.columns if f'{value_col}_ACF_Lag_' in str(col)]
        print(f"   Existing ACF columns found: {len(existing_acf_cols)}")
        if existing_acf_cols:
            print(f"   Existing columns: {existing_acf_cols}")
        
        # Call original function
        result = original_add_acf_pacf(df, value_col, sheet_type, include_confidence)
        track_dataframe_operation(result, "acf_pacf_output")
        
        # Check if same DataFrame object
        if id(df) == id(result):
            print(f"   ⚠️  SAME DATAFRAME OBJECT RETURNED")
        else:
            print(f"   ✅ NEW DATAFRAME OBJECT CREATED")
        
        return result
    
    ar_utils.add_acf_pacf_analysis = tracked_add_acf_pacf
    
    # Patch reorder_with_acf_pacf
    original_reorder = ar_utils.reorder_with_acf_pacf
    def tracked_reorder(df, base_order):
        print(f"\n🔄 DataFrame Reordering Starting")
        track_dataframe_operation(df, "reorder_input")
        
        result = original_reorder(df, base_order)
        track_dataframe_operation(result, "reorder_output")
        
        # Check column order changes
        if list(df.columns) != list(result.columns):
            print(f"   Column order changed: {len(df.columns)} -> {len(result.columns)}")
        
        return result
    
    ar_utils.reorder_with_acf_pacf = tracked_reorder

def patch_sheet_creator_methods():
    """Patch sheet creator methods for tracking"""
    
    from report_generator.sheet_creators import SheetCreator
    
    # Patch _run_aggregation
    original_run_agg = SheetCreator._run_aggregation
    def tracked_run_agg(self, pipeline, use_base_filter=True, collection_name='media_records'):
        print(f"\n💾 MongoDB Aggregation Starting")
        result = original_run_agg(self, pipeline, use_base_filter, collection_name)
        track_dataframe_operation(result, "mongodb_aggregation")
        return result
    
    SheetCreator._run_aggregation = tracked_run_agg
    
    # Patch _fill_missing_collection_days
    original_fill_missing = SheetCreator._fill_missing_collection_days
    def tracked_fill_missing(self, df, pipeline_name):
        print(f"\n🕳️  Fill Missing Days Starting")
        track_dataframe_operation(df, "fill_missing_input")
        
        result = original_fill_missing(self, df, pipeline_name)
        track_dataframe_operation(result, "fill_missing_output")
        
        return result
    
    SheetCreator._fill_missing_collection_days = tracked_fill_missing

def analyze_duplication_patterns():
    """Analyze the collected tracking data for duplication patterns"""
    print(f"\n" + "="*60)
    print(f"DUPLICATION PATTERN ANALYSIS")
    print(f"="*60)
    
    # Find DataFrames with duplicates
    problematic_dfs = []
    for df_id, operations in dataframe_operations.items():
        for op in operations:
            if op['has_duplicates']:
                problematic_dfs.append((df_id, op))
    
    if problematic_dfs:
        print(f"\n❌ DATAFRAMES WITH DUPLICATES: {len(problematic_dfs)}")
        for df_id, op in problematic_dfs:
            print(f"  DF {df_id}: {op['operation']} - {op['duplicate_columns']}")
            
            # Show the operation sequence for this DataFrame
            print(f"    Operation history for DF {df_id}:")
            for hist_op in dataframe_operations[df_id]:
                status = "❌" if hist_op['has_duplicates'] else "✅"
                print(f"      {status} {hist_op['operation']}: {hist_op['total_columns']} cols ({hist_op['unique_columns']} unique)")
    
    # Look for 4x multiplication pattern
    multiplication_patterns = []
    for df_id, operations in dataframe_operations.items():
        if len(operations) >= 2:
            start_cols = operations[0]['unique_columns']
            for op in operations[1:]:
                if op['unique_columns'] > 0 and start_cols > 0:
                    ratio = op['total_columns'] / start_cols
                    if abs(ratio - 4.0) < 0.1:  # Close to 4x
                        multiplication_patterns.append((df_id, start_cols, op['total_columns'], ratio, op['operation']))
    
    if multiplication_patterns:
        print(f"\n🎯 4X MULTIPLICATION PATTERNS DETECTED:")
        for df_id, start, end, ratio, operation in multiplication_patterns:
            print(f"  DF {df_id}: {start} -> {end} columns ({ratio:.1f}x) at {operation}")

def debug_single_pipeline_with_tracking():
    """Debug a single ACF/PACF pipeline with comprehensive tracking"""
    
    # Apply all patches
    monkey_patch_dataframe_operations()
    patch_ar_utils_functions()
    patch_sheet_creator_methods()
    
    print("🚀 Starting tracked single pipeline debug...")
    
    try:
        from db_utils import get_db_connection
        from report_generator.sheet_creators import SheetCreator
        from report_generator.formatters import ExcelFormatter
        
        # Setup
        db = get_db_connection()
        formatter = ExcelFormatter()
        sheet_creator = SheetCreator(db, formatter)
        
        # Load config and find a test sheet
        config_path = Path(__file__).parent / "report_config.json"
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        test_sheet = next((s for s in config.get('sheets', []) if 'Daily Counts (ACF_PACF)' in s.get('name', '')), None)
        
        if test_sheet:
            print(f"Testing sheet: {test_sheet['name']}")
            
            # Create workbook and test single sheet
            import openpyxl
            workbook = openpyxl.Workbook()
            workbook.remove(workbook.active)
            
            # Track the sheet creation process
            sheet_creator._create_pipeline_sheet(workbook, test_sheet)
            
            # Analyze results
            analyze_duplication_patterns()
            
            workbook.close()
        
    except Exception as e:
        print(f"❌ Error during tracked debug: {e}")
        traceback.print_exc()
        
        # Still analyze what we captured
        analyze_duplication_patterns()

if __name__ == "__main__":
    debug_single_pipeline_with_tracking()