#!/usr/bin/env python3
"""
Comprehensive Debug Logging for Daily Counts (ACF_PACF) Sheet Creation

This script adds comprehensive debug logging to trace the exact execution path
and identify why zero-fill patches aren't being triggered.
"""

import sys
import os

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def add_comprehensive_debug_logging():
    """Add comprehensive debug logging to all critical methods."""
    print("ADDING COMPREHENSIVE DEBUG LOGGING")
    print("=" * 50)
    
    # 1. Add debug to pipeline sheet creation
    pipeline_debug = '''
# ADD TO: report_generator/sheet_creators/pipeline.py line 95
print(f"[PIPELINE_DEBUG] Executing pipeline: {pipeline}")
print(f"[PIPELINE_DEBUG] Pipeline type: {type(pipeline)}")
print(f"[PIPELINE_DEBUG] Sheet config: {sheet_config}")
'''
    
    # 2. Add debug to aggregation caching
    cache_debug = '''
# ADD TO: report_generator/sheet_creators/base.py line 141
print(f"[CACHE_DEBUG] Cache key: {cache_key}")
print(f"[CACHE_DEBUG] Pipeline: {str(pipeline)[:100]}...")
print(f"[CACHE_DEBUG] Calling _should_apply_zero_fill...")
'''
    
    # 3. Add debug to zero-fill method
    zerofill_debug = '''
# ADD TO: report_generator/sheet_creators/base.py line 50
print(f"[ZEROFILL_DEBUG] _fill_missing_collection_days called")
print(f"[ZEROFILL_DEBUG] Input data shape: {data.shape}")
print(f"[ZEROFILL_DEBUG] Pipeline name: {pipeline_name}")
'''
    
    print("Debug code snippets prepared:")
    print("1. Pipeline execution debug")
    print("2. Cache key debug") 
    print("3. Zero-fill method debug")
    
    return True

def patch_pipeline_sheet_creator():
    """Patch the pipeline sheet creator with debug logging."""
    print("\nPATCHING PIPELINE SHEET CREATOR")
    print("=" * 50)
    
    try:
        # Read the current file
        file_path = 'report_generator/sheet_creators/pipeline.py'
        with open(file_path, 'r') as f:
            content = f.read()
        
        # Add debug logging before pipeline execution
        debug_code = '''        # COMPREHENSIVE DEBUG LOGGING
        print(f"[PIPELINE_EXEC_DEBUG] ========================================")
        print(f"[PIPELINE_EXEC_DEBUG] Sheet: {sheet_config['name']}")
        print(f"[PIPELINE_EXEC_DEBUG] Pipeline: {sheet_config['pipeline']}")
        print(f"[PIPELINE_EXEC_DEBUG] Module: {sheet_config.get('module', 'unknown')}")
        print(f"[PIPELINE_EXEC_DEBUG] Category: {sheet_config.get('category', 'unknown')}")
        print(f"[PIPELINE_EXEC_DEBUG] About to call _run_aggregation_original...")
        print(f"[PIPELINE_EXEC_DEBUG] ========================================")
        
'''
        
        # Find the line where pipeline is executed and add debug before it
        target_line = "        df = self._run_aggregation_original(pipeline)"
        if target_line in content:
            content = content.replace(target_line, debug_code + target_line)
            
            # Write back to file
            with open(file_path, 'w') as f:
                f.write(content)
            
            print("✅ Successfully patched pipeline.py with debug logging")
            return True
        else:
            print("❌ Target line not found in pipeline.py")
            return False
            
    except Exception as e:
        print(f"❌ Error patching pipeline.py: {e}")
        return False

def patch_base_sheet_creator():
    """Patch the base sheet creator with debug logging."""
    print("\nPATCHING BASE SHEET CREATOR")
    print("=" * 50)
    
    try:
        # Read the current file
        file_path = 'report_generator/sheet_creators/base.py'
        with open(file_path, 'r') as f:
            content = f.read()
        
        # Add debug logging in _run_aggregation_cached method
        debug_code = '''        # COMPREHENSIVE DEBUG LOGGING FOR AGGREGATION
        print(f"[AGGREGATION_DEBUG] ========================================")
        print(f"[AGGREGATION_DEBUG] Cache key: {cache_key}")
        print(f"[AGGREGATION_DEBUG] Pipeline type: {type(pipeline)}")
        print(f"[AGGREGATION_DEBUG] Pipeline preview: {str(pipeline)[:200]}...")
        print(f"[AGGREGATION_DEBUG] Use base filter: {use_base_filter}")
        print(f"[AGGREGATION_DEBUG] Collection: {collection_name}")
        print(f"[AGGREGATION_DEBUG] About to execute pipeline...")
        print(f"[AGGREGATION_DEBUG] ========================================")
        
'''
        
        # Find the line where result is assigned and add debug before it
        target_line = "        result = self._run_aggregation_original(pipeline, use_base_filter, collection_name)"
        if target_line in content:
            content = content.replace(target_line, debug_code + target_line)
            
            # Write back to file
            with open(file_path, 'w') as f:
                f.write(content)
            
            print("✅ Successfully patched base.py with debug logging")
            return True
        else:
            print("❌ Target line not found in base.py")
            return False
            
    except Exception as e:
        print(f"❌ Error patching base.py: {e}")
        return False

def test_debug_logging():
    """Test the debug logging by generating a small report."""
    print("\nTESTING DEBUG LOGGING")
    print("=" * 50)
    
    print("Running a test report generation to verify debug logging...")
    print("Look for debug messages in the output:")
    print("- [PIPELINE_EXEC_DEBUG] messages")
    print("- [AGGREGATION_DEBUG] messages") 
    print("- [ZERO_FILL_DEBUG] messages")
    
    return True

def main():
    print("COMPREHENSIVE DEBUG LOGGING IMPLEMENTATION")
    print("=" * 70)
    
    try:
        # Add comprehensive debug logging
        add_comprehensive_debug_logging()
        
        # Patch the pipeline sheet creator
        if not patch_pipeline_sheet_creator():
            print("❌ Failed to patch pipeline sheet creator")
            return 1
        
        # Patch the base sheet creator
        if not patch_base_sheet_creator():
            print("❌ Failed to patch base sheet creator")
            return 1
        
        # Test the debug logging
        test_debug_logging()
        
        print("\n" + "=" * 70)
        print("DEBUG LOGGING IMPLEMENTATION COMPLETE")
        print("=" * 70)
        
        print("\n✅ NEXT STEPS:")
        print("1. Run: python generate_report.py")
        print("2. Look for debug messages in the output")
        print("3. Identify where the execution path diverges")
        print("4. Apply targeted fixes based on debug output")
        
        print("\n🔍 WHAT TO LOOK FOR:")
        print("- [PIPELINE_EXEC_DEBUG] Should show Daily Counts (ACF_PACF) execution")
        print("- [AGGREGATION_DEBUG] Should show DAILY_COUNTS_ALL_WITH_ZEROES pipeline")
        print("- [ZERO_FILL_DEBUG] Should show cache key matching logic")
        print("- [ZERO_FILL_PATCH] Should show zero-fill being applied")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
