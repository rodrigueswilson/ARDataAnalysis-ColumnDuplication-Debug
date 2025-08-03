#!/usr/bin/env python3
"""
Simple Data Cleaning Test

This script tests if the Data Cleaning sheet creation is working by examining
the actual report generation process and checking for exceptions.
"""

import sys
import os
import traceback

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_data_cleaning_in_report_generation():
    """Test Data Cleaning sheet creation within the full report generation."""
    print("SIMPLE DATA CLEANING TEST")
    print("=" * 60)
    
    try:
        # Import the main report generator
        from generate_report import main as generate_main
        
        print("✅ Successfully imported generate_report")
        
        # Monkey patch the create_data_cleaning_sheet method to add debug info
        print("\n🔍 PATCHING create_data_cleaning_sheet FOR DEBUG")
        print("=" * 50)
        
        # Find and patch the method
        try:
            from report_generator.sheet_creators.base import BaseSheetCreator
            
            # Store original method
            original_method = BaseSheetCreator.create_data_cleaning_sheet
            
            def debug_create_data_cleaning_sheet(self, workbook):
                """Debug wrapper for create_data_cleaning_sheet."""
                print("[DEBUG_PATCH] create_data_cleaning_sheet called!")
                print(f"[DEBUG_PATCH] Workbook sheets before: {workbook.sheetnames}")
                
                try:
                    # Call original method
                    result = original_method(self, workbook)
                    print(f"[DEBUG_PATCH] create_data_cleaning_sheet completed successfully")
                    print(f"[DEBUG_PATCH] Workbook sheets after: {workbook.sheetnames}")
                    return result
                    
                except Exception as e:
                    print(f"[DEBUG_PATCH] ❌ Exception in create_data_cleaning_sheet: {e}")
                    print(f"[DEBUG_PATCH] Full traceback:")
                    traceback.print_exc()
                    raise
            
            # Apply patch
            BaseSheetCreator.create_data_cleaning_sheet = debug_create_data_cleaning_sheet
            print("✅ Successfully patched create_data_cleaning_sheet")
            
        except Exception as e:
            print(f"❌ Failed to patch method: {e}")
            return False
        
        # Also patch the core.py call to see if it's being reached
        try:
            import report_generator.core as core_module
            
            # Find the ReportGenerator class
            original_generate = core_module.ReportGenerator.generate_report
            
            def debug_generate_report(self):
                """Debug wrapper for generate_report."""
                print("[DEBUG_CORE] generate_report called!")
                
                try:
                    result = original_generate(self)
                    print("[DEBUG_CORE] generate_report completed")
                    return result
                except Exception as e:
                    print(f"[DEBUG_CORE] ❌ Exception in generate_report: {e}")
                    traceback.print_exc()
                    raise
            
            core_module.ReportGenerator.generate_report = debug_generate_report
            print("✅ Successfully patched generate_report")
            
        except Exception as e:
            print(f"❌ Failed to patch generate_report: {e}")
        
        # Run report generation
        print(f"\n🔄 RUNNING REPORT GENERATION WITH DEBUG PATCHES")
        print("=" * 50)
        
        try:
            # Run the report generation
            result = generate_main()
            print(f"✅ Report generation completed with result: {result}")
            return True
            
        except Exception as e:
            print(f"❌ Report generation failed: {e}")
            traceback.print_exc()
            return False
            
    except Exception as e:
        print(f"❌ Error in test setup: {e}")
        traceback.print_exc()
        return False

def check_recent_changes_impact():
    """Check what our recent changes might have broken."""
    print(f"\n🔍 CHECKING RECENT CHANGES IMPACT")
    print("=" * 50)
    
    print("📋 OUR RECENT CHANGES TO base.py:")
    print("   1. Added debug logging to _run_aggregation_cached")
    print("   2. Modified _should_apply_zero_fill method")
    print("   3. Added _fill_missing_collection_days patches")
    
    print(f"\n🚨 POTENTIAL IMPACT ON Data Cleaning:")
    print("   - create_data_cleaning_sheet uses direct MongoDB aggregation")
    print("   - Does NOT use _run_aggregation_cached method")
    print("   - Should NOT be affected by our pipeline changes")
    print("   - But might be affected by import/initialization issues")
    
    print(f"\n🎯 HYPOTHESIS:")
    print("   - Data Cleaning sheet creation is being called")
    print("   - But failing silently due to an exception")
    print("   - Our debug patches should reveal the exact issue")

def main():
    print("SIMPLE DATA CLEANING SHEET TEST")
    print("=" * 70)
    
    try:
        # Check recent changes impact
        check_recent_changes_impact()
        
        # Test Data Cleaning in report generation
        success = test_data_cleaning_in_report_generation()
        
        print("\n" + "=" * 70)
        print("TEST RESULTS")
        print("=" * 70)
        
        if success:
            print("✅ Data Cleaning sheet creation appears to be working")
            print("   Check the debug output above for details")
        else:
            print("❌ Data Cleaning sheet creation is broken")
            print("   Review the exception details above")
        
        return 0 if success else 1
        
    except Exception as e:
        print(f"❌ Error in main: {e}")
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
