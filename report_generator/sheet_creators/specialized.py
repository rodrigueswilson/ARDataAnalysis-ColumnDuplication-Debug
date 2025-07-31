"""
Specialized Analysis Sheet Creator Module
========================================

This module contains specialized sheet creators for advanced analysis sheets
such as Audio Efficiency Details and MP3 Duration Analysis.
"""

import pandas as pd
from .base import BaseSheetCreator


class SpecializedSheetCreator(BaseSheetCreator):
    """
    Handles creation of specialized analysis sheets with custom formatting
    and multi-table layouts.
    """
    
    def create_audio_efficiency_details_sheet(self, workbook):
        """
        Creates the Audio Efficiency Details sheet with enhanced filtering and positioning.
        
        Args:
            workbook: openpyxl workbook object
        """
        try:
            # Get audio efficiency data
            pipeline = [
                {"$match": {
                    "file_type": "MP3",
                    "is_collection_day": True,
                    "Outlier_Status": False
                }},
                {"$group": {
                    "_id": {
                        "Collection_Period": "$Collection_Period",
                        "School_Year": "$School_Year"
                    },
                    "Total_MP3_Files": {"$sum": 1},
                    "Total_Duration_Seconds": {"$sum": "$Duration_Seconds"},
                    "Avg_Duration_Seconds": {"$avg": "$Duration_Seconds"},
                    "Days_With_MP3": {"$addToSet": "$ISO_Date"}
                }},
                {"$addFields": {
                    "Days_With_MP3": {"$size": "$Days_With_MP3"},
                    "Total_Duration_Hours": {"$divide": ["$Total_Duration_Seconds", 3600]},
                    "Avg_Files_Per_Day": {"$divide": ["$Total_MP3_Files", {"$size": "$Days_With_MP3"}]},
                    "Period_Efficiency": {"$multiply": [
                        {"$divide": ["$Total_Duration_Seconds", 3600]}, 100
                    ]}
                }},
                {"$sort": {
                    "_id.School_Year": 1,
                    "_id.Collection_Period": 1
                }}
            ]
            
            df = self._run_aggregation(pipeline)
            if df.empty:
                print("[WARNING] No data found for Audio Efficiency Details")
                return
            
            # Create worksheet
            ws = workbook.create_sheet("Audio Efficiency Details")
            
            # Title
            ws['A1'] = "AR Data Analysis - Audio Efficiency Details"
            self.formatter.apply_title_style(ws, 'A1')
            
            # Headers
            headers = [
                'Collection Period', 'School Year', 'Total MP3 Files', 'Total Duration (Hours)',
                'Avg Duration per File', 'Days with MP3', 'Avg Files per Day', 'Period Efficiency'
            ]
            
            for col, header in enumerate(headers, 1):
                ws.cell(row=3, column=col, value=header)
            
            self.formatter.apply_header_style(ws, 'A3:H3')
            
            # Helper function to convert seconds to HH:MM:SS format
            def seconds_to_hms(seconds):
                if pd.isna(seconds) or seconds == 0:
                    return "00:00:00"
                hours = int(seconds // 3600)
                minutes = int((seconds % 3600) // 60)
                seconds = int(seconds % 60)
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            
            # Add data rows
            row = 4
            for _, record in df.iterrows():
                period = record.get('Collection_Period', 'Unknown')
                school_year = record.get('School_Year', 'Unknown')
                files = record.get('Total_MP3_Files', 0)
                duration_hours = record.get('Total_Duration_Hours', 0)
                avg_duration = record.get('Avg_Duration_Seconds', 0)
                days_with_mp3 = record.get('Days_With_MP3', 0)
                avg_files_per_day = record.get('Avg_Files_Per_Day', 0)
                period_efficiency = record.get('Period_Efficiency', 0)
                
                ws.cell(row=row, column=1, value=period)
                ws.cell(row=row, column=2, value=school_year)
                ws.cell(row=row, column=3, value=files)
                ws.cell(row=row, column=4, value=round(duration_hours, 2))
                ws.cell(row=row, column=5, value=seconds_to_hms(avg_duration))
                ws.cell(row=row, column=6, value=days_with_mp3)
                ws.cell(row=row, column=7, value=round(avg_files_per_day, 1))
                ws.cell(row=row, column=8, value=round(period_efficiency, 1))
                
                row += 1
            
            # Apply formatting
            self.formatter.apply_data_style(ws, f'A4:H{row-1}')
            self.formatter.auto_adjust_columns(ws)
            
            print("[SUCCESS] Audio Efficiency Details sheet created")
            
        except Exception as e:
            print(f"[ERROR] Failed to create Audio Efficiency Details sheet: {e}")

    def create_mp3_duration_analysis_sheet(self, workbook):
        """
        Creates comprehensive MP3 Duration Analysis sheet with multiple tables.
        
        Args:
            workbook: openpyxl workbook object
        """
        try:
            # Get MP3 duration data by school year
            school_year_pipeline = [
                {"$match": {
                    "file_type": "MP3",
                    "is_collection_day": True,
                    "Outlier_Status": False
                }},
                {"$group": {
                    "_id": "$School_Year",
                    "Total_MP3_Files": {"$sum": 1},
                    "Total_Duration_Seconds": {"$sum": "$Duration_Seconds"},
                    "Avg_Duration_Seconds": {"$avg": "$Duration_Seconds"},
                    "Min_Duration_Seconds": {"$min": "$Duration_Seconds"},
                    "Max_Duration_Seconds": {"$max": "$Duration_Seconds"}
                }},
                {"$addFields": {
                    "Total_Duration_Hours": {"$divide": ["$Total_Duration_Seconds", 3600]}
                }},
                {"$sort": {"_id": 1}}
            ]
            
            # Get MP3 duration data by period
            period_pipeline = [
                {"$match": {
                    "file_type": "MP3",
                    "is_collection_day": True,
                    "Outlier_Status": False
                }},
                {"$group": {
                    "_id": {
                        "Collection_Period": "$Collection_Period",
                        "School_Year": "$School_Year"
                    },
                    "Total_MP3_Files": {"$sum": 1},
                    "Total_Duration_Seconds": {"$sum": "$Duration_Seconds"},
                    "Avg_Duration_Seconds": {"$avg": "$Duration_Seconds"},
                    "Days_With_MP3": {"$addToSet": "$ISO_Date"}
                }},
                {"$addFields": {
                    "Days_With_MP3": {"$size": "$Days_With_MP3"},
                    "Total_Duration_Hours": {"$divide": ["$Total_Duration_Seconds", 3600]},
                    "Avg_Files_Per_Day": {"$divide": ["$Total_MP3_Files", {"$size": "$Days_With_MP3"}]},
                    "Period_Efficiency": {"$multiply": [
                        {"$divide": ["$Total_Duration_Seconds", 3600]}, 100
                    ]}
                }},
                {"$sort": {
                    "_id.School_Year": 1,
                    "_id.Collection_Period": 1
                }}
            ]
            
            # Get MP3 duration data by month
            monthly_pipeline = [
                {"$match": {
                    "file_type": "MP3",
                    "is_collection_day": True,
                    "Outlier_Status": False
                }},
                {"$group": {
                    "_id": {
                        "Month": "$ISO_Month",
                        "School_Year": "$School_Year"
                    },
                    "Total_MP3_Files": {"$sum": 1},
                    "Total_Duration_Seconds": {"$sum": "$Duration_Seconds"},
                    "Avg_Duration_Seconds": {"$avg": "$Duration_Seconds"}
                }},
                {"$addFields": {
                    "Total_Duration_Hours": {"$divide": ["$Total_Duration_Seconds", 3600]}
                }},
                {"$sort": {
                    "_id.School_Year": 1,
                    "_id.Month": 1
                }}
            ]
            
            # Execute pipelines
            school_year_df = self._run_aggregation(school_year_pipeline)
            period_df = self._run_aggregation(period_pipeline)
            monthly_df = self._run_aggregation(monthly_pipeline)
            
            if school_year_df.empty and period_df.empty and monthly_df.empty:
                print("[WARNING] No data found for MP3 Duration Analysis")
                return
            
            # Create worksheet
            ws = workbook.create_sheet("MP3 Duration Analysis")
            
            # Title
            ws['A1'] = "AR Data Analysis - MP3 Duration Analysis"
            self.formatter.apply_title_style(ws, 'A1')
            
            # Helper function to convert seconds to HH:MM:SS format
            def seconds_to_hms(seconds):
                if pd.isna(seconds) or seconds == 0:
                    return "00:00:00"
                hours = int(seconds // 3600)
                minutes = int((seconds % 3600) // 60)
                seconds = int(seconds % 60)
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            
            current_row = 3
            
            # Add School Year Duration Summary table
            if not school_year_df.empty:
                current_row = self._add_duration_summary_table(ws, school_year_df, current_row, seconds_to_hms)
                current_row += 3  # Add spacing
            
            # Add Collection Period Duration table
            if not period_df.empty:
                current_row = self._add_period_duration_table(ws, period_df, current_row, seconds_to_hms)
                current_row += 3  # Add spacing
            
            # Add Monthly Duration Distribution table
            if not monthly_df.empty:
                current_row = self._add_monthly_duration_table(ws, monthly_df, current_row, seconds_to_hms)
            
            # Apply formatting and auto-adjust columns
            self.formatter.auto_adjust_columns(ws)
            
            print("[SUCCESS] MP3 Duration Analysis sheet created")
            
        except Exception as e:
            print(f"[ERROR] Failed to create MP3 Duration Analysis sheet: {e}")

    def _add_duration_summary_table(self, ws, df, start_row, seconds_to_hms):
        """
        Adds the School Year Duration Summary table.
        """
        # Table title
        ws.cell(row=start_row, column=1, value="School Year MP3 Duration Summary")
        current_row = start_row + 2
        
        # Headers
        headers = [
            "School Year", "Total Files", "Total Hours", "Avg Duration", 
            "Min Duration", "Max Duration"
        ]
        
        for col, header in enumerate(headers, 1):
            ws.cell(row=current_row, column=col, value=header)
        current_row += 1
        
        # Data rows
        for _, row in df.iterrows():
            school_year = row.get('School_Year', 'Unknown')  # Flattened field
            files = row.get('Total_MP3_Files', 0)
            hours = row.get('Total_Duration_Hours', 0)
            avg_duration = row.get('Avg_Duration_Seconds', 0)
            min_duration = row.get('Min_Duration_Seconds', 0)
            max_duration = row.get('Max_Duration_Seconds', 0)
            
            ws.cell(row=current_row, column=1, value=school_year)
            ws.cell(row=current_row, column=2, value=files)
            ws.cell(row=current_row, column=3, value=round(hours, 2))
            ws.cell(row=current_row, column=4, value=seconds_to_hms(avg_duration))
            ws.cell(row=current_row, column=5, value=seconds_to_hms(min_duration))
            ws.cell(row=current_row, column=6, value=seconds_to_hms(max_duration))
            
            current_row += 1
        
        return current_row

    def _add_period_duration_table(self, ws, df, start_row, seconds_to_hms):
        """
        Adds the Collection Period Duration table.
        """
        # Table title
        ws.cell(row=start_row, column=1, value="Collection Period MP3 Duration Analysis")
        current_row = start_row + 2
        
        # Headers
        headers = [
            "Period", "School Year", "Files", "Total Hours", "Avg Duration",
            "Days with MP3", "Avg Files/Day", "Efficiency"
        ]
        
        for col, header in enumerate(headers, 1):
            ws.cell(row=current_row, column=col, value=header)
        current_row += 1
        
        # Data rows
        for _, row in df.iterrows():
            # Access flattened fields directly (after _run_aggregation processing)
            period = row.get('Collection_Period', 'Unknown')
            school_year = row.get('School_Year', 'Unknown')
            
            files = row.get('Total_MP3_Files', 0)
            duration_hours = row.get('Total_Duration_Hours', 0)
            avg_duration = row.get('Avg_Duration_Seconds', 0)
            days_with_mp3 = row.get('Days_With_MP3', 0)
            avg_files_per_day = row.get('Avg_Files_Per_Day', 0)
            period_efficiency = row.get('Period_Efficiency', 0)
            
            ws.cell(row=current_row, column=1, value=period)
            ws.cell(row=current_row, column=2, value=school_year)
            ws.cell(row=current_row, column=3, value=files)
            ws.cell(row=current_row, column=4, value=duration_hours)
            ws.cell(row=current_row, column=5, value=seconds_to_hms(avg_duration))
            ws.cell(row=current_row, column=6, value=days_with_mp3)
            ws.cell(row=current_row, column=7, value=avg_files_per_day)
            ws.cell(row=current_row, column=8, value=period_efficiency)
            
            current_row += 1
        
        return current_row
    
    def _add_monthly_duration_table(self, ws, df, start_row, seconds_to_hms):
        """
        Adds the Monthly Duration Distribution table.
        """
        # Table title
        ws.cell(row=start_row, column=1, value="Monthly MP3 Duration Distribution")
        current_row = start_row + 2
        
        # Headers
        headers = [
            "Month", "2021-22 Files", "2021-22 Hours", "2022-23 Files", 
            "2022-23 Hours", "Total Files", "Total Hours", "Avg Duration"
        ]
        
        for col, header in enumerate(headers, 1):
            ws.cell(row=current_row, column=col, value=header)
        current_row += 1
        
        # Process data by month
        month_names = [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ]
        
        # Group data by month
        monthly_data = {}
        # Process monthly data
        for _, row in df.iterrows():
            # Access flattened fields directly (after _run_aggregation processing)
            month_num = row.get('Month', 0)
            school_year = row.get('School_Year', 'Unknown')
            
            if month_num not in monthly_data:
                monthly_data[month_num] = {}
            
            monthly_data[month_num][school_year] = {
                'files': row.get('Total_MP3_Files', 0),
                'hours': row.get('Total_Duration_Hours', 0),
                'avg_duration': row.get('Avg_Duration_Seconds', 0)
            }
        
        # Write monthly rows
        for month_num in sorted(monthly_data.keys()):
            if month_num < 1 or month_num > 12:
                continue
                
            month_name = month_names[month_num - 1]
            month_data = monthly_data[month_num]
            
            # Get data for each school year
            sy_2122 = month_data.get('2021-2022', {'files': 0, 'hours': 0, 'avg_duration': 0})
            sy_2223 = month_data.get('2022-2023', {'files': 0, 'hours': 0, 'avg_duration': 0})
            
            total_files = sy_2122['files'] + sy_2223['files']
            total_hours = sy_2122['hours'] + sy_2223['hours']
            avg_duration = (sy_2122['avg_duration'] + sy_2223['avg_duration']) / 2 if (sy_2122['avg_duration'] > 0 and sy_2223['avg_duration'] > 0) else max(sy_2122['avg_duration'], sy_2223['avg_duration'])
            
            ws.cell(row=current_row, column=1, value=month_name)
            ws.cell(row=current_row, column=2, value=sy_2122['files'])
            ws.cell(row=current_row, column=3, value=sy_2122['hours'])
            ws.cell(row=current_row, column=4, value=sy_2223['files'])
            ws.cell(row=current_row, column=5, value=sy_2223['hours'])
            ws.cell(row=current_row, column=6, value=total_files)
            ws.cell(row=current_row, column=7, value=total_hours)
            ws.cell(row=current_row, column=8, value=seconds_to_hms(avg_duration))
            
            current_row += 1
        
        return current_row
