#!/usr/bin/env python3
"""
Test script to debug pipeline import issues.
"""

def test_pipeline_import():
    print("=== PIPELINE IMPORT DEBUG TEST ===")
    
    try:
        print("[1] Importing pipelines module...")
        import pipelines
        print("    ✅ pipelines module imported successfully")
        
        print("[2] Checking PIPELINES dictionary...")
        if hasattr(pipelines, 'PIPELINES'):
            print(f"    ✅ PIPELINES dictionary found with {len(pipelines.PIPELINES)} entries")
        else:
            print("    ❌ PIPELINES dictionary not found")
            return False
        
        print("[3] Checking for MP3 duration pipelines...")
        required_pipelines = ['MP3_DURATION_BY_SCHOOL_YEAR', 'MP3_DURATION_BY_PERIOD', 'MP3_DURATION_BY_MONTH']
        
        for pipeline_name in required_pipelines:
            if pipeline_name in pipelines.PIPELINES:
                pipeline = pipelines.PIPELINES[pipeline_name]
                print(f"    ✅ {pipeline_name}: {len(pipeline)} stages")
            else:
                print(f"    ❌ {pipeline_name} NOT found in PIPELINES")
        
        print("[4] Checking pipeline variables directly...")
        for pipeline_name in required_pipelines:
            if hasattr(pipelines, pipeline_name):
                pipeline = getattr(pipelines, pipeline_name)
                print(f"    ✅ {pipeline_name} variable: {len(pipeline)} stages")
            else:
                print(f"    ❌ {pipeline_name} variable not found")
        
        print("[5] Available pipeline keys:")
        available_keys = [k for k in pipelines.PIPELINES.keys() if 'MP3' in k or 'DURATION' in k]
        if available_keys:
            print(f"    MP3/Duration related keys: {available_keys}")
        else:
            print("    No MP3/Duration related keys found")
            
        print(f"[6] All pipeline keys containing 'MP3': {[k for k in pipelines.PIPELINES.keys() if 'MP3' in k]}")
        
        return True
        
    except Exception as e:
        print(f"❌ Import test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_pipeline_import()
    print(f"\n=== TEST {'PASSED' if success else 'FAILED'} ===")
