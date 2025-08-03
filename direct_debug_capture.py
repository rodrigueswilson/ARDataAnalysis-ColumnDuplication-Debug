#!/usr/bin/env python3
"""
Direct Debug Capture for Daily Counts (ACF_PACF) Sheet

This script captures debug output directly and analyzes the execution path.
"""

import sys
import os
import subprocess
import re

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def capture_debug_output():
    """Capture debug output from report generation."""
    print("CAPTURING DEBUG OUTPUT FOR Daily Counts (ACF_PACF)")
    print("=" * 60)
    
    try:
        # Run report generation and capture all output
        print("🔄 Running report generation...")
        result = subprocess.run(
            [sys.executable, "generate_report.py"],
            capture_output=True,
            text=True,
            cwd=".",
            timeout=300  # 5 minute timeout
        )
        
        print(f"✅ Report generation completed")
        print(f"📊 Return code: {result.returncode}")
        print(f"📝 Output lines: {len(result.stdout.split())}")
        
        # Analyze the output
        output_lines = result.stdout.split('\n')
        
        # Look for specific debug patterns
        pipeline_debug = []
        aggregation_debug = []
        zero_fill_debug = []
        daily_counts_refs = []
        
        for i, line in enumerate(output_lines):
            if "[PIPELINE_EXEC_DEBUG]" in line:
                pipeline_debug.append((i, line))
            elif "[AGGREGATION_DEBUG]" in line:
                aggregation_debug.append((i, line))
            elif "[ZERO_FILL_DEBUG]" in line:
                zero_fill_debug.append((i, line))
            elif "Daily Counts (ACF_PACF)" in line:
                daily_counts_refs.append((i, line))
        
        print(f"\n📊 DEBUG MESSAGE COUNTS:")
        print(f"   Pipeline debug messages: {len(pipeline_debug)}")
        print(f"   Aggregation debug messages: {len(aggregation_debug)}")
        print(f"   Zero-fill debug messages: {len(zero_fill_debug)}")
        print(f"   Daily Counts (ACF_PACF) references: {len(daily_counts_refs)}")
        
        # Show Daily Counts (ACF_PACF) specific messages
        if daily_counts_refs:
            print(f"\n🎯 DAILY COUNTS (ACF_PACF) REFERENCES:")
            for line_num, line in daily_counts_refs:
                print(f"   Line {line_num}: {line.strip()}")
        
        # Show pipeline debug messages
        if pipeline_debug:
            print(f"\n🔍 PIPELINE DEBUG MESSAGES (first 5):")
            for line_num, line in pipeline_debug[:5]:
                print(f"   Line {line_num}: {line.strip()}")
        
        # Show aggregation debug messages
        if aggregation_debug:
            print(f"\n🔍 AGGREGATION DEBUG MESSAGES (first 5):")
            for line_num, line in aggregation_debug[:5]:
                print(f"   Line {line_num}: {line.strip()}")
        
        # Show zero-fill debug messages
        if zero_fill_debug:
            print(f"\n🔍 ZERO-FILL DEBUG MESSAGES:")
            for line_num, line in zero_fill_debug:
                print(f"   Line {line_num}: {line.strip()}")
        
        # Look for DAILY_COUNTS_ALL_WITH_ZEROES pipeline references
        pipeline_refs = []
        for i, line in enumerate(output_lines):
            if "DAILY_COUNTS_ALL_WITH_ZEROES" in line:
                pipeline_refs.append((i, line))
        
        if pipeline_refs:
            print(f"\n🎯 DAILY_COUNTS_ALL_WITH_ZEROES REFERENCES:")
            for line_num, line in pipeline_refs:
                print(f"   Line {line_num}: {line.strip()}")
        
        return True
        
    except subprocess.TimeoutExpired:
        print("❌ Report generation timed out")
        return False
    except Exception as e:
        print(f"❌ Error capturing debug output: {e}")
        return False

def analyze_current_excel():
    """Analyze the current Excel file to understand the data."""
    print(f"\n📊 ANALYZING CURRENT EXCEL OUTPUT")
    print("=" * 60)
    
    try:
        import pandas as pd
        import glob
        
        # Find latest Excel file
        excel_files = glob.glob("AR_Analysis_Report_*.xlsx")
        if not excel_files:
            print("❌ No Excel files found")
            return False
        
        latest_file = max(excel_files, key=os.path.getctime)
        print(f"📄 Latest file: {latest_file}")
        
        # Read Daily Counts (ACF_PACF) sheet
        df = pd.read_excel(latest_file, sheet_name='Daily Counts (ACF_PACF)', header=2)
        print(f"📊 Sheet dimensions: {df.shape}")
        
        # Analyze the data
        if 'Total_Files' in df.columns:
            total = df['Total_Files'].sum()
            print(f"📈 Total files: {total}")
            
            # Check for zero values
            zero_count = (df['Total_Files'] == 0).sum()
            non_zero_count = (df['Total_Files'] > 0).sum()
            print(f"📊 Days with files: {non_zero_count}")
            print(f"📊 Days with zero files: {zero_count}")
            
            # Date analysis
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                valid_dates = df['Date'].dropna()
                if not valid_dates.empty:
                    print(f"📅 Date range: {valid_dates.min()} to {valid_dates.max()}")
                    print(f"📅 Total days: {len(valid_dates)}")
                    
                    # Expected vs actual days
                    date_range = pd.date_range(valid_dates.min(), valid_dates.max(), freq='D')
                    expected_days = len(date_range)
                    actual_days = len(valid_dates)
                    print(f"📅 Expected days in range: {expected_days}")
                    print(f"📅 Actual days present: {actual_days}")
                    print(f"📅 Missing days: {expected_days - actual_days}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error analyzing Excel: {e}")
        return False

def main():
    print("DIRECT DEBUG CAPTURE AND ANALYSIS")
    print("=" * 70)
    
    try:
        # Capture debug output
        if not capture_debug_output():
            print("❌ Failed to capture debug output")
            return 1
        
        # Analyze current Excel output
        if not analyze_current_excel():
            print("❌ Failed to analyze Excel output")
            return 1
        
        print("\n" + "=" * 70)
        print("ANALYSIS COMPLETE")
        print("=" * 70)
        
        print("\n💡 KEY FINDINGS:")
        print("- Check if Daily Counts (ACF_PACF) debug messages appeared")
        print("- Verify if DAILY_COUNTS_ALL_WITH_ZEROES pipeline is being used")
        print("- Confirm if zero-fill debug messages are triggered")
        print("- Analyze the actual data structure and totals")
        
        return 0
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
