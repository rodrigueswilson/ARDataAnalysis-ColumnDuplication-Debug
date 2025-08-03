#!/usr/bin/env python3
"""
Analyze Debug Output for Daily Counts (ACF_PACF) Sheet

This script analyzes the debug output to understand exactly what's happening
with the Daily Counts (ACF_PACF) sheet creation and zero-fill logic.
"""

import sys
import os

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def analyze_debug_output():
    """Analyze the debug output from the report generation."""
    print("ANALYZING DEBUG OUTPUT FOR Daily Counts (ACF_PACF)")
    print("=" * 60)
    
    print("\n🔍 OBSERVATIONS FROM DEBUG OUTPUT:")
    print("-" * 40)
    
    print("\n1. PIPELINE EXECUTION DEBUG:")
    print("   ✅ Debug messages are appearing for various sheets")
    print("   ✅ [PIPELINE_EXEC_DEBUG] messages are working")
    print("   ❓ Need to verify if Daily Counts (ACF_PACF) specific messages appear")
    
    print("\n2. AGGREGATION DEBUG:")
    print("   ✅ [AGGREGATION_DEBUG] messages should be appearing")
    print("   ❓ Need to check if DAILY_COUNTS_ALL_WITH_ZEROES pipeline is shown")
    
    print("\n3. ZERO-FILL DEBUG:")
    print("   ✅ [ZERO_FILL_DEBUG] messages should show cache key matching")
    print("   ❓ Need to verify if _should_apply_zero_fill is called")
    
    print("\n4. SPECIFIC FINDINGS:")
    print("   ✅ ARIMA processing shows 'Daily Counts (ACF_PACF)' sheet exists")
    print("   ✅ Sheet is being created and processed")
    print("   ❓ But zero-fill logic may not be triggered")
    
    return True

def create_targeted_debug_script():
    """Create a targeted debug script to capture Daily Counts (ACF_PACF) execution."""
    print("\nCREATING TARGETED DEBUG SCRIPT")
    print("=" * 60)
    
    script_content = '''#!/usr/bin/env python3
"""
Targeted Debug Script for Daily Counts (ACF_PACF) Sheet

This script generates a report with focused debug output filtering
to show only Daily Counts (ACF_PACF) related messages.
"""

import subprocess
import sys
import re

def run_report_with_filtered_debug():
    """Run report generation and filter debug output."""
    print("RUNNING REPORT WITH FILTERED DEBUG OUTPUT")
    print("=" * 50)
    
    try:
        # Run the report generation and capture output
        result = subprocess.run(
            [sys.executable, "generate_report.py"],
            capture_output=True,
            text=True,
            cwd="."
        )
        
        output_lines = result.stdout.split('\\n')
        
        # Filter for Daily Counts (ACF_PACF) related messages
        daily_counts_messages = []
        acf_pacf_messages = []
        debug_messages = []
        
        for line in output_lines:
            if "Daily Counts (ACF_PACF)" in line:
                daily_counts_messages.append(line)
            elif "ACF_PACF" in line:
                acf_pacf_messages.append(line)
            elif any(debug_tag in line for debug_tag in [
                "[PIPELINE_EXEC_DEBUG]", 
                "[AGGREGATION_DEBUG]", 
                "[ZERO_FILL_DEBUG]",
                "[ZERO_FILL_PATCH]"
            ]):
                debug_messages.append(line)
        
        print(f"\\n📊 FILTERED DEBUG RESULTS:")
        print(f"Daily Counts (ACF_PACF) messages: {len(daily_counts_messages)}")
        print(f"ACF_PACF related messages: {len(acf_pacf_messages)}")
        print(f"Debug messages: {len(debug_messages)}")
        
        if daily_counts_messages:
            print(f"\\n🎯 DAILY COUNTS (ACF_PACF) MESSAGES:")
            for msg in daily_counts_messages:
                print(f"  {msg}")
        
        if debug_messages:
            print(f"\\n🔍 DEBUG MESSAGES:")
            for msg in debug_messages[:10]:  # Show first 10
                print(f"  {msg}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error running filtered debug: {e}")
        return False

if __name__ == "__main__":
    run_report_with_filtered_debug()
'''
    
    # Write the targeted debug script
    with open('targeted_debug_daily_counts.py', 'w') as f:
        f.write(script_content)
    
    print("✅ Created targeted_debug_daily_counts.py")
    return True

def check_current_excel_output():
    """Check the current Excel output to see the actual data."""
    print("\nCHECKING CURRENT EXCEL OUTPUT")
    print("=" * 60)
    
    try:
        import pandas as pd
        import glob
        
        # Find the latest Excel file
        excel_files = glob.glob("AR_Analysis_Report_*.xlsx")
        if not excel_files:
            print("❌ No Excel files found")
            return False
        
        latest_file = max(excel_files, key=os.path.getctime)
        print(f"📊 Latest Excel file: {latest_file}")
        
        # Read the Daily Counts (ACF_PACF) sheet
        try:
            df = pd.read_excel(latest_file, sheet_name='Daily Counts (ACF_PACF)', header=2)
            print(f"✅ Sheet dimensions: {df.shape}")
            
            # Check for Total_Files column
            if 'Total_Files' in df.columns:
                total_files = df['Total_Files'].sum()
                print(f"📈 Current total files: {total_files}")
                
                # Check date range
                if 'Date' in df.columns:
                    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                    date_range = df['Date'].dropna()
                    if not date_range.empty:
                        print(f"📅 Date range: {date_range.min()} to {date_range.max()}")
                        print(f"📅 Number of dates: {len(date_range)}")
                        
                        # Check for gaps
                        date_range_sorted = date_range.sort_values()
                        expected_days = (date_range_sorted.max() - date_range_sorted.min()).days + 1
                        actual_days = len(date_range_sorted)
                        print(f"📅 Expected days in range: {expected_days}")
                        print(f"📅 Actual days present: {actual_days}")
                        print(f"📅 Missing days: {expected_days - actual_days}")
                
                return True
            else:
                print("❌ Total_Files column not found")
                print(f"Available columns: {list(df.columns)}")
                return False
                
        except Exception as e:
            print(f"❌ Error reading sheet: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Error checking Excel output: {e}")
        return False

def main():
    print("COMPREHENSIVE DEBUG OUTPUT ANALYSIS")
    print("=" * 70)
    
    try:
        # Analyze the debug output
        analyze_debug_output()
        
        # Create targeted debug script
        create_targeted_debug_script()
        
        # Check current Excel output
        check_current_excel_output()
        
        print("\n" + "=" * 70)
        print("ANALYSIS COMPLETE")
        print("=" * 70)
        
        print("\n💡 NEXT STEPS:")
        print("1. Run: python targeted_debug_daily_counts.py")
        print("2. Look for specific Daily Counts (ACF_PACF) debug messages")
        print("3. Verify if DAILY_COUNTS_ALL_WITH_ZEROES pipeline is being used")
        print("4. Check if zero-fill logic is being triggered")
        
        print("\n🎯 KEY QUESTIONS TO ANSWER:")
        print("- Is [PIPELINE_EXEC_DEBUG] showing Daily Counts (ACF_PACF)?")
        print("- Is [AGGREGATION_DEBUG] showing DAILY_COUNTS_ALL_WITH_ZEROES?")
        print("- Is [ZERO_FILL_DEBUG] showing cache key matching?")
        print("- Is [ZERO_FILL_PATCH] being applied?")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
