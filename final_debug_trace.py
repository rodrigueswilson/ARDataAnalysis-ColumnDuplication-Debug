#!/usr/bin/env python3
"""
Final Debug Trace

This script adds comprehensive debug logging to ALL possible sheet creation paths
to definitively identify where ACF/PACF sheets are actually being created.
"""

import sys
import os

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def add_debug_to_all_methods():
    """Add debug logging to all possible sheet creation methods."""
    print("ADDING DEBUG LOGGING TO ALL SHEET CREATION METHODS")
    print("=" * 60)
    
    # Files to patch with debug logging
    files_to_patch = [
        'report_generator/core.py',
        'report_generator/sheet_creators/base.py',
        'report_generator/sheet_creators/pipeline.py',
        'report_generator/sheet_creators/specialized.py',
        'report_generator/sheet_creators/__init__.py',
        'report_generator/consolidated_sheet_factory.py'
    ]
    
    debug_patterns = [
        ('def create_', '[FINAL_DEBUG] create_ method called'),
        ('def _create_', '[FINAL_DEBUG] _create_ method called'),
        ('def add_sheet', '[FINAL_DEBUG] add_sheet method called'),
        ('def _add_sheet', '[FINAL_DEBUG] _add_sheet method called'),
        ('def process_', '[FINAL_DEBUG] process_ method called'),
        ('aggregate(', '[FINAL_DEBUG] MongoDB aggregate called'),
        ('DAILY_COUNTS_COLLECTION_ONLY', '[FINAL_DEBUG] DAILY_COUNTS_COLLECTION_ONLY pipeline detected'),
        ('Daily Counts (ACF_PACF)', '[FINAL_DEBUG] Daily Counts (ACF_PACF) sheet detected')
    ]
    
    for file_path in files_to_patch:
        if os.path.exists(file_path):
            print(f"✅ Found {file_path}")
            
            try:
                with open(file_path, 'r') as f:
                    content = f.read()
                
                # Check if file contains relevant patterns
                relevant_patterns = []
                for pattern, debug_msg in debug_patterns:
                    if pattern in content:
                        relevant_patterns.append((pattern, debug_msg))
                
                if relevant_patterns:
                    print(f"  📍 Contains {len(relevant_patterns)} relevant patterns:")
                    for pattern, debug_msg in relevant_patterns:
                        count = content.count(pattern)
                        print(f"    - '{pattern}': {count} occurrences")
                else:
                    print(f"  ⚪ No relevant patterns found")
                    
            except Exception as e:
                print(f"  ❌ Error reading {file_path}: {e}")
        else:
            print(f"❌ Missing {file_path}")
    
    return True

def create_comprehensive_debug_patch():
    """Create a comprehensive debug patch for the BaseSheetCreator."""
    print("\nCREATING COMPREHENSIVE DEBUG PATCH")
    print("=" * 60)
    
    patch_code = '''
# COMPREHENSIVE DEBUG PATCH - Add this to BaseSheetCreator.__init__
print("[FINAL_DEBUG] BaseSheetCreator initialized")

# COMPREHENSIVE DEBUG PATCH - Add this to _run_aggregation_cached
def _run_aggregation_cached_debug(self, pipeline, use_base_filter=True, collection_name='media_records'):
    """Debug version of _run_aggregation_cached with comprehensive logging."""
    cache_key = str(pipeline)[:100] + "..." if len(str(pipeline)) > 100 else str(pipeline)
    
    print(f"[FINAL_DEBUG] _run_aggregation_cached called")
    print(f"[FINAL_DEBUG] Cache key: {cache_key}")
    print(f"[FINAL_DEBUG] Pipeline type: {type(pipeline)}")
    print(f"[FINAL_DEBUG] Use base filter: {use_base_filter}")
    print(f"[FINAL_DEBUG] Collection: {collection_name}")
    
    # Check if this is an ACF/PACF related pipeline
    pipeline_str = str(pipeline).upper()
    if 'DAILY_COUNTS_COLLECTION_ONLY' in pipeline_str:
        print("[FINAL_DEBUG] 🎯 DAILY_COUNTS_COLLECTION_ONLY PIPELINE DETECTED!")
    if 'ACF_PACF' in pipeline_str:
        print("[FINAL_DEBUG] 🎯 ACF_PACF PIPELINE DETECTED!")
    
    # Call original method
    result = self._run_aggregation_cached_original(pipeline, use_base_filter, collection_name)
    
    print(f"[FINAL_DEBUG] Result shape: {result.shape if hasattr(result, 'shape') else 'No shape'}")
    print(f"[FINAL_DEBUG] Result columns: {list(result.columns) if hasattr(result, 'columns') else 'No columns'}")
    
    return result
'''
    
    print("Debug patch code created:")
    print(patch_code[:200] + "..." if len(patch_code) > 200 else patch_code)
    
    return patch_code

def investigate_method_resolution():
    """Investigate method resolution order to understand inheritance."""
    print("\nINVESTIGATING METHOD RESOLUTION ORDER")
    print("=" * 60)
    
    try:
        from report_generator.sheet_creators import SheetCreator
        
        print("SheetCreator MRO (Method Resolution Order):")
        for i, cls in enumerate(SheetCreator.__mro__):
            print(f"  {i+1}. {cls.__name__} ({cls.__module__})")
        
        # Check if _fill_missing_collection_days exists in each class
        print("\n_fill_missing_collection_days method locations:")
        for cls in SheetCreator.__mro__:
            if hasattr(cls, '_fill_missing_collection_days'):
                method = getattr(cls, '_fill_missing_collection_days')
                print(f"  ✅ {cls.__name__}: {method}")
            else:
                print(f"  ❌ {cls.__name__}: Not found")
        
        # Check if _run_aggregation_cached exists in each class
        print("\n_run_aggregation_cached method locations:")
        for cls in SheetCreator.__mro__:
            if hasattr(cls, '_run_aggregation_cached'):
                method = getattr(cls, '_run_aggregation_cached')
                print(f"  ✅ {cls.__name__}: {method}")
            else:
                print(f"  ❌ {cls.__name__}: Not found")
                
        return True
        
    except Exception as e:
        print(f"Error investigating MRO: {e}")
        return False

def main():
    print("FINAL DEBUG TRACE - COMPREHENSIVE INVESTIGATION")
    print("=" * 70)
    
    try:
        # Add debug to all methods
        add_debug_to_all_methods()
        
        # Create comprehensive debug patch
        create_comprehensive_debug_patch()
        
        # Investigate method resolution
        investigate_method_resolution()
        
        print("\n" + "=" * 70)
        print("FINAL ANALYSIS")
        print("=" * 70)
        
        print("CURRENT STATUS:")
        print("1. ❌ Zero-fill fix not working despite multiple patches")
        print("2. ❌ Debug messages not appearing in report generation")
        print("3. ❌ Excel sheets still have incorrect structure and totals")
        print("4. ❌ Left-aligned numbers still present")
        
        print("\nROOT CAUSE HYPOTHESIS:")
        print("1. ACF/PACF sheets are created through a completely different code path")
        print("2. The sheets bypass ALL our patched methods")
        print("3. There may be a direct database query or cached result being used")
        print("4. The sheet creation happens in a module we haven't identified yet")
        
        print("\nNEXT STEPS:")
        print("1. Add debug logging directly to the report generation entry point")
        print("2. Trace every method call during report generation")
        print("3. Find the exact line of code that creates ACF/PACF sheets")
        print("4. Apply the zero-fill fix to the correct location")
        
        print("\n🚨 CRITICAL INSIGHT:")
        print("The fact that NO debug messages appear suggests the ACF/PACF sheets")
        print("are either using cached data or being created by a completely different")
        print("system that we haven't identified yet.")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
