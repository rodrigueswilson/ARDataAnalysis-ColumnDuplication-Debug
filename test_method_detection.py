#!/usr/bin/env python3
"""
Test Method Detection

This script tests why the create_data_cleaning_sheet method isn't being detected.
"""

import sys
import os
import inspect

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_method_detection():
    """Test if we can detect the create_data_cleaning_sheet method."""
    print("METHOD DETECTION TEST")
    print("=" * 50)
    
    try:
        # Import the BaseSheetCreator class
        from report_generator.sheet_creators.base import BaseSheetCreator
        
        print("✅ Successfully imported BaseSheetCreator")
        
        # Check if the method exists
        if hasattr(BaseSheetCreator, 'create_data_cleaning_sheet'):
            print("✅ create_data_cleaning_sheet method EXISTS")
            
            # Get method info
            method = getattr(BaseSheetCreator, 'create_data_cleaning_sheet')
            print(f"📋 Method type: {type(method)}")
            print(f"📋 Method callable: {callable(method)}")
            
            # Get method signature
            try:
                sig = inspect.signature(method)
                print(f"📋 Method signature: {sig}")
            except Exception as e:
                print(f"⚠️  Could not get signature: {e}")
            
            # Get method source location
            try:
                source_file = inspect.getfile(method)
                source_lines = inspect.getsourcelines(method)
                print(f"📋 Source file: {source_file}")
                print(f"📋 Source lines: {source_lines[1]}-{source_lines[1] + len(source_lines[0])}")
            except Exception as e:
                print(f"⚠️  Could not get source info: {e}")
            
        else:
            print("❌ create_data_cleaning_sheet method NOT FOUND")
            
            # List all methods to see what's available
            print("\n📋 Available methods in BaseSheetCreator:")
            methods = [name for name in dir(BaseSheetCreator) if not name.startswith('_')]
            for method_name in sorted(methods):
                if 'data' in method_name.lower() or 'clean' in method_name.lower():
                    print(f"   🔍 {method_name}")
                elif method_name.startswith('create'):
                    print(f"   📄 {method_name}")
        
        # Test instantiation
        print(f"\n🔍 TESTING INSTANTIATION:")
        try:
            # Try to create an instance (this might fail due to missing dependencies)
            instance = BaseSheetCreator(None, None)
            print("✅ BaseSheetCreator instantiated successfully")
            
            if hasattr(instance, 'create_data_cleaning_sheet'):
                print("✅ Instance has create_data_cleaning_sheet method")
            else:
                print("❌ Instance does NOT have create_data_cleaning_sheet method")
                
        except Exception as e:
            print(f"⚠️  Could not instantiate BaseSheetCreator: {e}")
            print("   (This is expected due to missing database/config)")
        
        return True
        
    except Exception as e:
        print(f"❌ Import error: {e}")
        import traceback
        traceback.print_exc()
        return False

def check_file_syntax():
    """Check if there are syntax errors in base.py that might prevent method loading."""
    print(f"\n🔍 CHECKING FILE SYNTAX")
    print("=" * 50)
    
    try:
        # Try to compile the base.py file
        with open('report_generator/sheet_creators/base.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Try to compile
        compile(content, 'report_generator/sheet_creators/base.py', 'exec')
        print("✅ base.py syntax is valid")
        
        # Check for the method definition in the raw content
        if 'def create_data_cleaning_sheet' in content:
            print("✅ create_data_cleaning_sheet definition found in file")
            
            # Count the occurrences
            count = content.count('def create_data_cleaning_sheet')
            print(f"📊 Found {count} definition(s)")
            
            # Check indentation
            lines = content.split('\n')
            for i, line in enumerate(lines, 1):
                if 'def create_data_cleaning_sheet' in line:
                    print(f"📋 Line {i}: {repr(line)}")
                    # Check if it's properly indented (should be class method)
                    if not line.startswith('    def '):
                        print("⚠️  WARNING: Method may not be properly indented as class method")
                    else:
                        print("✅ Method appears to be properly indented")
        else:
            print("❌ create_data_cleaning_sheet definition NOT found in file")
        
        return True
        
    except SyntaxError as e:
        print(f"❌ Syntax error in base.py: {e}")
        print(f"   Line {e.lineno}: {e.text}")
        return False
    except Exception as e:
        print(f"❌ Error checking syntax: {e}")
        return False

def main():
    print("FOCUSED METHOD DETECTION TEST")
    print("=" * 60)
    
    try:
        # Check file syntax first
        syntax_ok = check_file_syntax()
        
        # Test method detection
        detection_ok = test_method_detection()
        
        print("\n" + "=" * 60)
        print("DETECTION TEST RESULTS")
        print("=" * 60)
        
        if syntax_ok and detection_ok:
            print("✅ Method should be accessible - issue may be elsewhere")
        else:
            print("❌ Found issues with method detection")
            print(f"   Syntax OK: {'✅' if syntax_ok else '❌'}")
            print(f"   Detection OK: {'✅' if detection_ok else '❌'}")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error in main: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
