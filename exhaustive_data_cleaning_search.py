#!/usr/bin/env python3
"""
Exhaustive Data Cleaning Sheet Search

This script performs an exhaustive search for all references to the Data Cleaning sheet
to understand what happened and identify what our recent changes may have broken.
"""

import sys
import os
import subprocess
import re

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def run_grep_search(pattern, description, case_sensitive=False):
    """Run a grep search and return results."""
    print(f"\n🔍 {description}")
    print("=" * 60)
    
    try:
        # Build grep command
        cmd = ["grep", "-r", "-n"]
        if not case_sensitive:
            cmd.append("-i")
        cmd.extend([pattern, "."])
        
        # Exclude certain directories and file types
        excludes = [
            "--exclude-dir=.git",
            "--exclude-dir=__pycache__",
            "--exclude-dir=.pytest_cache",
            "--exclude=*.pyc",
            "--exclude=*.xlsx",
            "--exclude=*.log"
        ]
        cmd.extend(excludes)
        
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=".")
        
        if result.returncode == 0 and result.stdout.strip():
            lines = result.stdout.strip().split('\n')
            print(f"📊 Found {len(lines)} matches:")
            
            # Group by file
            file_matches = {}
            for line in lines:
                if ':' in line:
                    file_path = line.split(':', 1)[0]
                    if file_path not in file_matches:
                        file_matches[file_path] = []
                    file_matches[file_path].append(line)
            
            # Show results grouped by file
            for file_path, matches in file_matches.items():
                print(f"\n📄 {file_path} ({len(matches)} matches):")
                for match in matches[:5]:  # Show first 5 matches per file
                    parts = match.split(':', 2)
                    if len(parts) >= 3:
                        line_num = parts[1]
                        content = parts[2].strip()
                        print(f"   Line {line_num}: {content}")
                if len(matches) > 5:
                    print(f"   ... and {len(matches) - 5} more matches")
            
            return file_matches
        else:
            print("❌ No matches found")
            return {}
            
    except Exception as e:
        print(f"❌ Error running grep: {e}")
        return {}

def search_data_cleaning_references():
    """Search for all Data Cleaning related references."""
    print("EXHAUSTIVE DATA CLEANING SHEET SEARCH")
    print("=" * 80)
    
    # Search patterns
    searches = [
        ("Data Cleaning", "Direct 'Data Cleaning' references"),
        ("data.cleaning", "data.cleaning references (case insensitive)", False),
        ("DataCleaning", "DataCleaning class/variable references"),
        ("data_cleaning", "data_cleaning underscore references"),
        ("DASHBOARD_DATA_CLEANING", "DASHBOARD_DATA_CLEANING pipeline references"),
        ("DATA_CLEANING_ANALYSIS", "DATA_CLEANING_ANALYSIS pipeline references"),
        ("sheet_name.*[Dd]ata.*[Cc]leaning", "Sheet name configurations"),
        ("module.*data_cleaning", "Module references to data_cleaning"),
        ("pipeline.*[Dd]ata.*[Cc]leaning", "Pipeline configurations"),
    ]
    
    all_results = {}
    
    for pattern, description, *args in searches:
        case_sensitive = args[0] if args else True
        results = run_grep_search(pattern, description, case_sensitive)
        all_results[pattern] = results
    
    return all_results

def analyze_recent_changes():
    """Analyze what files we recently modified that might affect Data Cleaning."""
    print(f"\n🔍 ANALYZING RECENT CHANGES THAT MIGHT AFFECT DATA CLEANING")
    print("=" * 80)
    
    # Files we know we modified
    modified_files = [
        "report_generator/sheet_creators/base.py",
        "report_generator/sheet_creators/pipeline.py", 
        "report_generator/core.py",
        "report_config.json",
        "generate_report.py"
    ]
    
    print("📋 FILES WE RECENTLY MODIFIED:")
    for file_path in modified_files:
        if os.path.exists(file_path):
            print(f"   ✅ {file_path}")
            
            # Check if this file has Data Cleaning references
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                    
                # Look for Data Cleaning references
                data_cleaning_refs = []
                lines = content.split('\n')
                for i, line in enumerate(lines, 1):
                    if 'data' in line.lower() and 'clean' in line.lower():
                        data_cleaning_refs.append((i, line.strip()))
                
                if data_cleaning_refs:
                    print(f"      🚨 Contains Data Cleaning references:")
                    for line_num, line_content in data_cleaning_refs[:3]:
                        print(f"         Line {line_num}: {line_content}")
                    if len(data_cleaning_refs) > 3:
                        print(f"         ... and {len(data_cleaning_refs) - 3} more")
                else:
                    print(f"      ✅ No obvious Data Cleaning references")
                    
            except Exception as e:
                print(f"      ❌ Error reading file: {e}")
        else:
            print(f"   ❌ {file_path} (not found)")

def check_data_cleaning_module_structure():
    """Check the structure of data cleaning related modules."""
    print(f"\n🔍 CHECKING DATA CLEANING MODULE STRUCTURE")
    print("=" * 60)
    
    # Look for data cleaning modules
    potential_modules = [
        "utils/data_cleaning.py",
        "report_generator/pipelines/data_cleaning.py",
        "report_generator/sheet_creators/data_cleaning.py",
        "data_cleaning.py"
    ]
    
    print("📋 SEARCHING FOR DATA CLEANING MODULES:")
    for module_path in potential_modules:
        if os.path.exists(module_path):
            print(f"   ✅ Found: {module_path}")
            
            # Check file size and basic structure
            try:
                stat = os.stat(module_path)
                print(f"      Size: {stat.st_size} bytes")
                
                with open(module_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                    lines = len(content.split('\n'))
                    print(f"      Lines: {lines}")
                    
                    # Look for class/function definitions
                    classes = re.findall(r'^class\s+(\w+)', content, re.MULTILINE)
                    functions = re.findall(r'^def\s+(\w+)', content, re.MULTILINE)
                    
                    if classes:
                        print(f"      Classes: {', '.join(classes[:5])}")
                    if functions:
                        print(f"      Functions: {', '.join(functions[:5])}")
                        
            except Exception as e:
                print(f"      ❌ Error analyzing: {e}")
        else:
            print(f"   ❌ Not found: {module_path}")

def identify_potential_breakage():
    """Identify what might have broken the Data Cleaning sheet."""
    print(f"\n🚨 IDENTIFYING POTENTIAL BREAKAGE POINTS")
    print("=" * 60)
    
    print("🎯 HYPOTHESIS: Our recent changes broke Data Cleaning sheet")
    print("\n📋 POTENTIAL CAUSES:")
    
    causes = [
        "1. Changes to base.py _run_aggregation_cached method affected all sheets",
        "2. Debug logging additions interfered with sheet creation",
        "3. Zero-fill logic changes affected pipeline execution",
        "4. Pipeline configuration changes in report_config.json",
        "5. Import or module loading issues from our modifications",
        "6. Cache key or pipeline matching logic changes",
        "7. Exception handling changes that silently fail Data Cleaning creation"
    ]
    
    for cause in causes:
        print(f"   {cause}")
    
    print(f"\n🔧 INVESTIGATION PRIORITIES:")
    priorities = [
        "1. Check if Data Cleaning sheet was in original report_config.json",
        "2. Review all changes to base.py for unintended side effects", 
        "3. Test Data Cleaning sheet generation in isolation",
        "4. Compare current vs previous report_config.json",
        "5. Check if our debug logging broke sheet creation flow"
    ]
    
    for priority in priorities:
        print(f"   {priority}")

def main():
    print("EXHAUSTIVE DATA CLEANING SHEET INVESTIGATION")
    print("=" * 90)
    
    try:
        # Search for all Data Cleaning references
        search_results = search_data_cleaning_references()
        
        # Analyze recent changes
        analyze_recent_changes()
        
        # Check module structure
        check_data_cleaning_module_structure()
        
        # Identify potential breakage
        identify_potential_breakage()
        
        print("\n" + "=" * 90)
        print("EXHAUSTIVE SEARCH COMPLETE")
        print("=" * 90)
        
        # Summary
        total_files_with_refs = sum(len(files) for files in search_results.values())
        print(f"\n📊 SUMMARY:")
        print(f"   - Total files with Data Cleaning references: {total_files_with_refs}")
        print(f"   - Search patterns executed: {len(search_results)}")
        print(f"   - Files we modified that might affect Data Cleaning: Multiple")
        
        print(f"\n🎯 NEXT STEPS:")
        print("   1. Review the grep results above for clues")
        print("   2. Check if Data Cleaning was in original config")
        print("   3. Test reverting specific changes to isolate the cause")
        print("   4. Focus on base.py and report_config.json changes")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error in exhaustive search: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
