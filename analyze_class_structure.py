#!/usr/bin/env python3
"""
Comprehensive Class Structure Analysis

This script analyzes the exact structure of base.py to find why the 
create_data_cleaning_sheet method isn't being loaded into the class.
"""

import sys
import os
import ast
import traceback

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def analyze_file_structure():
    """Analyze the AST structure of base.py to find the issue."""
    print("COMPREHENSIVE CLASS STRUCTURE ANALYSIS")
    print("=" * 60)
    
    try:
        # Read the file
        with open('report_generator/sheet_creators/base.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"✅ File read successfully ({len(content)} characters)")
        
        # Parse the AST
        try:
            tree = ast.parse(content)
            print("✅ AST parsing successful")
        except SyntaxError as e:
            print(f"❌ Syntax error in file: {e}")
            print(f"   Line {e.lineno}: {e.text}")
            return False
        
        # Analyze the AST structure
        print(f"\n🔍 AST STRUCTURE ANALYSIS:")
        print("=" * 40)
        
        classes_found = []
        functions_found = []
        
        for node in ast.walk(tree):
            if isinstance(node, ast.ClassDef):
                classes_found.append({
                    'name': node.name,
                    'line': node.lineno,
                    'methods': [n.name for n in node.body if isinstance(n, ast.FunctionDef)]
                })
            elif isinstance(node, ast.FunctionDef) and node.col_offset == 0:
                # Top-level function (not in a class)
                functions_found.append({
                    'name': node.name,
                    'line': node.lineno
                })
        
        print(f"📊 Classes found: {len(classes_found)}")
        for cls in classes_found:
            print(f"   📄 {cls['name']} (line {cls['line']})")
            print(f"      Methods: {len(cls['methods'])}")
            if 'create_data_cleaning_sheet' in cls['methods']:
                print(f"      ✅ create_data_cleaning_sheet FOUND in {cls['name']}")
            else:
                print(f"      ❌ create_data_cleaning_sheet NOT in {cls['name']}")
            
            # Show first few methods
            for method in cls['methods'][:5]:
                print(f"         - {method}")
            if len(cls['methods']) > 5:
                print(f"         ... and {len(cls['methods']) - 5} more")
        
        print(f"\n📊 Top-level functions found: {len(functions_found)}")
        for func in functions_found:
            print(f"   🔧 {func['name']} (line {func['line']})")
            if func['name'] == 'create_data_cleaning_sheet':
                print(f"      🚨 FOUND create_data_cleaning_sheet as TOP-LEVEL function!")
                print(f"         This explains why it's not accessible as a class method!")
        
        # Check if create_data_cleaning_sheet is orphaned
        orphaned_method = any(f['name'] == 'create_data_cleaning_sheet' for f in functions_found)
        in_class = any('create_data_cleaning_sheet' in cls['methods'] for cls in classes_found)
        
        print(f"\n🎯 DIAGNOSIS:")
        print(f"   create_data_cleaning_sheet in class: {'✅' if in_class else '❌'}")
        print(f"   create_data_cleaning_sheet orphaned: {'🚨' if orphaned_method else '✅'}")
        
        if orphaned_method:
            print(f"\n🚨 ROOT CAUSE IDENTIFIED:")
            print(f"   The create_data_cleaning_sheet method is defined as a TOP-LEVEL function")
            print(f"   instead of being inside the BaseSheetCreator class!")
            print(f"   This is why it's not accessible as a class method.")
        
        return not orphaned_method
        
    except Exception as e:
        print(f"❌ Error analyzing file structure: {e}")
        traceback.print_exc()
        return False

def find_exact_issue_location():
    """Find the exact location where the class structure breaks."""
    print(f"\n🔍 FINDING EXACT ISSUE LOCATION")
    print("=" * 50)
    
    try:
        with open('report_generator/sheet_creators/base.py', 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        print(f"📊 Total lines: {len(lines)}")
        
        # Find class definition
        class_start = None
        for i, line in enumerate(lines):
            if line.strip().startswith('class BaseSheetCreator'):
                class_start = i + 1
                print(f"✅ BaseSheetCreator class starts at line {class_start}")
                break
        
        if not class_start:
            print("❌ BaseSheetCreator class definition not found!")
            return False
        
        # Find create_data_cleaning_sheet definition
        method_line = None
        for i, line in enumerate(lines):
            if 'def create_data_cleaning_sheet' in line:
                method_line = i + 1
                print(f"✅ create_data_cleaning_sheet found at line {method_line}")
                print(f"   Line content: {repr(line.rstrip())}")
                
                # Check indentation
                indentation = len(line) - len(line.lstrip())
                print(f"   Indentation: {indentation} spaces")
                
                if indentation == 4:
                    print("   ✅ Correct indentation for class method")
                elif indentation == 0:
                    print("   🚨 PROBLEM: No indentation - this is a top-level function!")
                else:
                    print(f"   ⚠️  Unusual indentation: {indentation} spaces")
                
                break
        
        if not method_line:
            print("❌ create_data_cleaning_sheet definition not found!")
            return False
        
        # Check if the method is after the class definition
        if method_line > class_start:
            print(f"✅ Method is after class definition (good)")
            
            # Check for any class-ending issues between class start and method
            for i in range(class_start, method_line):
                line = lines[i].rstrip()
                if line and not line.startswith(' ') and not line.startswith('\t'):
                    if not line.startswith('#') and line != '':
                        print(f"🚨 POTENTIAL ISSUE at line {i+1}: {repr(line)}")
                        print(f"   This line might be breaking the class definition!")
        else:
            print(f"❌ Method is BEFORE class definition - this shouldn't happen!")
        
        return True
        
    except Exception as e:
        print(f"❌ Error finding issue location: {e}")
        return False

def main():
    print("COMPREHENSIVE CLASS STRUCTURE ANALYSIS")
    print("=" * 70)
    
    try:
        # Analyze file structure
        structure_ok = analyze_file_structure()
        
        # Find exact issue location
        location_found = find_exact_issue_location()
        
        print("\n" + "=" * 70)
        print("ANALYSIS RESULTS")
        print("=" * 70)
        
        if structure_ok and location_found:
            print("✅ Class structure appears correct")
            print("   Issue may be elsewhere")
        else:
            print("🚨 STRUCTURAL ISSUES FOUND")
            print("   The create_data_cleaning_sheet method is orphaned")
            print("   It needs to be moved inside the BaseSheetCreator class")
        
        print(f"\n🎯 NEXT STEPS:")
        if not structure_ok:
            print("   1. Fix the orphaned create_data_cleaning_sheet method")
            print("   2. Ensure it's properly indented inside BaseSheetCreator class")
            print("   3. Test method accessibility")
        else:
            print("   1. Investigate import or module loading issues")
            print("   2. Check for circular imports or other problems")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error in main: {e}")
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
