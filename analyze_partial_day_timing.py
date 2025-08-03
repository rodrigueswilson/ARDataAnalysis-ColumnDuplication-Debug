#!/usr/bin/env python3
"""
Partial Day Timing Analysis

This script analyzes the timing of file captures on partial collection days
to determine if any files were captured after 12:00 PM (noon).

Focus on 2021-09-20 which is configured as a "Partial" day due to 
"Early Dismissal - Professional Development".
"""

import sys
import os
from datetime import datetime, time
import pandas as pd

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_utils import get_db_connection
from ar_utils import get_non_collection_days

def analyze_partial_day_timing():
    """Analyze timing of files captured on partial collection days."""
    print("PARTIAL DAY TIMING ANALYSIS")
    print("=" * 60)
    
    try:
        # Get database connection
        db = get_db_connection()
        collection = db.media_records
        
        # Get partial days from configuration
        non_collection_days = get_non_collection_days()
        partial_days = []
        
        for date_obj, info in non_collection_days.items():
            if info.get('type') == 'Partial':
                partial_days.append({
                    'date': date_obj.strftime('%Y-%m-%d'),
                    'event': info.get('event', 'Unknown'),
                    'date_obj': date_obj
                })
        
        print(f"Found {len(partial_days)} partial collection days:")
        for day in partial_days:
            print(f"  - {day['date']}: {day['event']}")
        
        # Focus on 2021-09-20 first
        target_date = "2021-09-20"
        print(f"\n" + "=" * 60)
        print(f"DETAILED ANALYSIS: {target_date}")
        print("=" * 60)
        
        # Query all files from this date with timestamp information
        query = {
            "ISO_Date": target_date,
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": ["JPG", "MP3"]},
            "is_collection_day": True,
            "Outlier_Status": False
        }
        
        # Get files with timing information
        files = list(collection.find(query, {
            "file_path": 1,
            "file_type": 1,
            "timestamp": 1,
            "ISO_Date": 1,
            "ISO_Time": 1,
            "_id": 0
        }).sort("timestamp", 1))
        
        print(f"Total files on {target_date}: {len(files)}")
        
        if not files:
            print("No files found for this date.")
            return
        
        # Analyze timing patterns
        print(f"\nTIMING ANALYSIS:")
        print("-" * 40)
        
        # Parse timestamps and categorize
        morning_files = []  # Before 12:00 PM
        afternoon_files = []  # After 12:00 PM
        unknown_time_files = []
        
        for file_info in files:
            iso_time = file_info.get('ISO_Time')
            timestamp = file_info.get('timestamp')
            
            if iso_time:
                try:
                    # Parse time (format should be HH:MM:SS)
                    time_parts = iso_time.split(':')
                    if len(time_parts) >= 2:
                        hour = int(time_parts[0])
                        minute = int(time_parts[1])
                        
                        file_time = time(hour, minute)
                        noon = time(12, 0)
                        
                        if file_time < noon:
                            morning_files.append(file_info)
                        else:
                            afternoon_files.append(file_info)
                    else:
                        unknown_time_files.append(file_info)
                except (ValueError, IndexError):
                    unknown_time_files.append(file_info)
            else:
                unknown_time_files.append(file_info)
        
        # Report results
        print(f"Files captured BEFORE 12:00 PM: {len(morning_files)}")
        print(f"Files captured AFTER 12:00 PM:  {len(afternoon_files)}")
        print(f"Files with unknown time:        {len(unknown_time_files)}")
        
        # Show timing distribution
        if morning_files or afternoon_files:
            print(f"\nTIME DISTRIBUTION:")
            print("-" * 40)
            
            # Create hourly breakdown
            hourly_counts = {}
            for file_info in morning_files + afternoon_files:
                iso_time = file_info.get('ISO_Time', '')
                if iso_time and ':' in iso_time:
                    try:
                        hour = int(iso_time.split(':')[0])
                        hourly_counts[hour] = hourly_counts.get(hour, 0) + 1
                    except ValueError:
                        pass
            
            for hour in sorted(hourly_counts.keys()):
                count = hourly_counts[hour]
                time_str = f"{hour:02d}:00"
                marker = "🌅" if hour < 12 else "🌇"
                print(f"  {time_str}: {count:3d} files {marker}")
        
        # Show sample files from afternoon if any
        if afternoon_files:
            print(f"\nSAMPLE AFTERNOON FILES (after 12:00 PM):")
            print("-" * 40)
            for i, file_info in enumerate(afternoon_files[:10]):  # Show first 10
                file_path = file_info.get('file_path', 'Unknown')
                iso_time = file_info.get('ISO_Time', 'Unknown')
                file_type = file_info.get('file_type', 'Unknown')
                
                # Extract filename from path
                filename = file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]
                
                print(f"  {i+1:2d}. {iso_time} - {filename} ({file_type})")
                
                if i == 9 and len(afternoon_files) > 10:
                    print(f"     ... and {len(afternoon_files) - 10} more files")
        
        # Analyze all partial days
        print(f"\n" + "=" * 60)
        print("ALL PARTIAL DAYS SUMMARY")
        print("=" * 60)
        
        for day_info in partial_days:
            date_str = day_info['date']
            event = day_info['event']
            
            # Query files for this date
            query['ISO_Date'] = date_str
            day_files = list(collection.find(query, {
                "ISO_Time": 1,
                "_id": 0
            }))
            
            if not day_files:
                print(f"{date_str}: No files found")
                continue
            
            # Count morning vs afternoon
            morning_count = 0
            afternoon_count = 0
            
            for file_info in day_files:
                iso_time = file_info.get('ISO_Time')
                if iso_time and ':' in iso_time:
                    try:
                        hour = int(iso_time.split(':')[0])
                        if hour < 12:
                            morning_count += 1
                        else:
                            afternoon_count += 1
                    except ValueError:
                        pass
            
            total_files = len(day_files)
            print(f"{date_str}: {total_files:3d} files total")
            print(f"  Event: {event}")
            print(f"  Before 12:00 PM: {morning_count:3d} files")
            print(f"  After 12:00 PM:  {afternoon_count:3d} files")
            
            if afternoon_count > 0:
                print(f"  🚨 ALERT: {afternoon_count} files captured after noon on partial day!")
            else:
                print(f"  ✅ All files captured before noon (as expected)")
            print()
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    try:
        success = analyze_partial_day_timing()
        
        if success:
            print("=" * 60)
            print("ANALYSIS COMPLETE")
            print("=" * 60)
            print("Key Questions Answered:")
            print("1. Are files captured after 12:00 PM on partial days?")
            print("2. What is the timing distribution on partial days?")
            print("3. Does the timing align with 'Early Dismissal' events?")
        else:
            print("Analysis failed - check error messages above")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
