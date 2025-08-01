"""
Sheet Creators Package
=====================

This package contains modular sheet creator classes for generating different
types of Excel sheets in the AR Data Analysis system.

The package is organized into focused modules:
- base: Core functionality and basic sheets (summary, raw data, data cleaning)
- pipeline: Configuration-driven sheets with ACF/PACF and ARIMA analysis
- specialized: Advanced analysis sheets (audio efficiency, MP3 duration)

Usage:
    from report_generator.sheet_creators import SheetCreator
    
    # Create a unified sheet creator with all capabilities
    creator = SheetCreator(db, formatter)
    
    # Create individual sheets
    creator.create_summary_statistics_sheet(workbook)
    creator.process_pipeline_configurations(workbook)
    creator.create_mp3_duration_analysis_sheet(workbook)
"""

from .base import BaseSheetCreator
from .pipeline import PipelineSheetCreator
from .specialized import SpecializedSheetCreator


class SheetCreator(SpecializedSheetCreator, PipelineSheetCreator):
    """
    Unified sheet creator that combines all sheet creation capabilities.
    
    This class inherits from all specialized sheet creator classes to provide
    a single interface for creating all types of sheets in the system.
    
    Features:
    - Base sheets: Summary statistics, raw data, data cleaning
    - Pipeline sheets: Configuration-driven with ACF/PACF and ARIMA analysis
    - Specialized sheets: Audio efficiency details, MP3 duration analysis
    """
    
    def __init__(self, db, formatter):
        """
        Initialize the unified sheet creator.
        
        Args:
            db: MongoDB database connection
            formatter: ExcelFormatter instance for styling
        """
        # Initialize the base class (all others inherit from it)
        super().__init__(db, formatter)
    
    def _calculate_consecutive_days(self, df, total_collection_days):
        """
        Calculate maximum consecutive days with and without data.
        This method is needed for the day analysis tables in Summary Statistics.
        """
        try:
            if df.empty:
                return 0, total_collection_days
            
            # Get unique dates with data
            dates_with_data = set(df['Date'].dt.date)
            
            # Create a simple consecutive day calculation
            # This is a simplified version - in a full implementation,
            # you'd want to consider the actual collection day calendar
            
            sorted_dates = sorted(dates_with_data)
            if not sorted_dates:
                return 0, total_collection_days
            
            # Calculate consecutive days with data
            max_consecutive_data = 1
            current_consecutive_data = 1
            
            for i in range(1, len(sorted_dates)):
                if (sorted_dates[i] - sorted_dates[i-1]).days == 1:
                    current_consecutive_data += 1
                    max_consecutive_data = max(max_consecutive_data, current_consecutive_data)
                else:
                    current_consecutive_data = 1
            
            # Estimate consecutive days without data (simplified)
            max_consecutive_zero = max(1, total_collection_days - len(dates_with_data))
            
            return max_consecutive_data, max_consecutive_zero
            
        except Exception as e:
            print(f"[WARNING] Error calculating consecutive days: {e}")
            return 0, 0
    
    def create_data_cleaning_sheet(self, workbook):
        """
        Creates the Data Cleaning sheet.
        """
        try:
            # Get outlier data
            pipeline = [
                {"$match": {"file_type": {"$in": ["JPG", "MP3"]}}},
                {"$group": {
                    "_id": {
                        "outlier_status": "$Outlier_Status",
                        "file_type": "$file_type"
                    },
                    "count": {"$sum": 1},
                    "total_size_mb": {"$sum": "$File_Size_MB"},
                    "avg_size_mb": {"$avg": "$File_Size_MB"}
                }},
                {"$sort": {"_id.file_type": 1, "_id.outlier_status": 1}}
            ]
            
            df = self._run_aggregation(pipeline)
            if df.empty:
                print("[WARNING] No data found for Data Cleaning sheet")
                return
            
            # Create worksheet
            ws = workbook.create_sheet("Data Cleaning")
            
            # Title
            ws['A1'] = "AR Data Analysis - Data Quality & Cleaning Report"
            self.formatter.apply_title_style(ws, 'A1')
            
            # Section 1: Outlier Analysis
            ws['A3'] = "Outlier Analysis"
            self.formatter.apply_section_header_style(ws, 'A3')
            
            # Headers for outlier table
            outlier_headers = ['File Type', 'Outlier Status', 'Count', 'Total Size (MB)', 'Avg Size (MB)']
            for col, header in enumerate(outlier_headers, 1):
                ws.cell(row=5, column=col, value=header)
            self.formatter.apply_header_style(ws, 'A5:E5')
            
            # Outlier data
            row = 6
            for _, record in df.iterrows():
                ws.cell(row=row, column=1, value=record.get('file_type', 'Unknown'))
                ws.cell(row=row, column=2, value=record.get('outlier_status', 'Unknown'))
                ws.cell(row=row, column=3, value=record.get('count', 0))
                ws.cell(row=row, column=4, value=round(record.get('total_size_mb', 0), 2))
                ws.cell(row=row, column=5, value=round(record.get('avg_size_mb', 0), 2))
                row += 1
            
            # Apply formatting
            self.formatter.apply_data_style(ws, f'A6:E{row-1}')
            
            # Apply alternating row colors for better readability
            self.formatter.apply_alternating_row_colors(ws, 6, row-1, 1, 5)
            
            # Section 2: Data Quality Metrics
            ws[f'A{row+2}'] = "Data Quality Metrics"
            self.formatter.apply_section_header_style(ws, f'A{row+2}')
            
            # Get quality metrics
            total_records = self.db['media_records'].count_documents({"file_type": {"$in": ["JPG", "MP3"]}})
            outlier_records = self.db['media_records'].count_documents({
                "file_type": {"$in": ["JPG", "MP3"]},
                "Outlier_Status": True
            })
            valid_records = total_records - outlier_records
            
            quality_metrics = [
                ('Total Records', total_records),
                ('Valid Records', valid_records),
                ('Outlier Records', outlier_records),
                ('Data Quality Rate', (valid_records/total_records) if total_records > 0 else "N/A")
            ]
            
            metrics_row = row + 4
            for i, (metric, value) in enumerate(quality_metrics):
                ws.cell(row=metrics_row + i, column=1, value=metric)
                cell = ws.cell(row=metrics_row + i, column=2)
                if metric == 'Data Quality Rate' and isinstance(value, (int, float)):
                    # Store as numeric value and apply percentage formatting
                    cell.value = value  # Already a decimal (0.0 to 1.0)
                    cell.number_format = '0.0%'  # Apply Excel percentage formatting
                else:
                    cell.value = value
            
            # Apply formatting
            self.formatter.apply_data_style(ws, f'A{metrics_row}:B{metrics_row + len(quality_metrics) - 1}')
            self.formatter.auto_adjust_columns(ws)
            
            print("[SUCCESS] Data Cleaning sheet created")
            
        except Exception as e:
            print(f"[ERROR] Failed to create Data Cleaning sheet: {e}")
    
    def _create_period_breakdown_table(self, ws, df, start_row):
        """
        Create the period breakdown table for day analysis.
        This method is needed for the Summary Statistics day analysis tables.
        """
        try:
            # Section title
            ws.cell(row=start_row, column=1, value="Day Analysis - Period Breakdown (Collection Days Only, Non-Outliers)")
            self.formatter.apply_title_style(ws, f'A{start_row}')
            start_row += 2
            
            # Headers
            headers = ['Period', 'Collection Days', 'Days with Data', 'Zero Days', 'Coverage %', 'Avg Files/Day', 'Max Consec. Data', 'Max Consec. Zero']
            for col, header in enumerate(headers, 1):
                ws.cell(row=start_row, column=col, value=header)
            self.formatter.apply_header_style(ws, f'A{start_row}:H{start_row}')
            start_row += 1
            
            # Get all periods
            from utils.calendar import get_all_periods, calculate_collection_days_for_period
            periods = get_all_periods()
            
            # Sort periods logically
            period_order = ['SY 21-22 P1', 'SY 21-22 P2', 'SY 21-22 P3', 'SY 22-23 P1', 'SY 22-23 P2', 'SY 22-23 P3']
            sorted_periods = [p for p in period_order if p in periods]
            
            # Fill period data
            for row_offset, period in enumerate(sorted_periods):
                row = start_row + row_offset
                period_df = df[df['Period'] == period].copy() if 'Period' in df.columns else pd.DataFrame()
                
                collection_days = calculate_collection_days_for_period(period)
                unique_dates = period_df['Date'].dt.date.unique() if not period_df.empty else []
                days_with_data = len(unique_dates)
                days_with_zero = collection_days - days_with_data
                coverage_pct = (days_with_data / collection_days * 100) if collection_days > 0 else 0
                avg_files_per_day = period_df['Total_Files'].mean() if not period_df.empty else 0
                
                # Calculate consecutive days
                consecutive_data, consecutive_zero = self._calculate_consecutive_days(period_df, collection_days)
                
                ws.cell(row=row, column=1, value=period)
                ws.cell(row=row, column=2, value=collection_days)
                ws.cell(row=row, column=3, value=days_with_data)
                ws.cell(row=row, column=4, value=days_with_zero)
                # Store coverage percentage as numeric value with Excel formatting
                cell = ws.cell(row=row, column=5)
                cell.value = coverage_pct / 100  # Convert percentage to decimal for Excel
                cell.number_format = '0.0%'  # Apply Excel percentage formatting
                ws.cell(row=row, column=6, value=round(avg_files_per_day, 1))
                ws.cell(row=row, column=7, value=consecutive_data)
                ws.cell(row=row, column=8, value=consecutive_zero)
            
            # Apply formatting
            end_row = start_row + len(sorted_periods) - 1
            self.formatter.apply_data_style(ws, f'A{start_row}:H{end_row}')
            self.formatter.apply_alternating_row_colors(ws, start_row, end_row, 1, 8)
            
            return end_row + 1
        
        except Exception as e:
            print(f"[ERROR] Failed to create period breakdown table: {e}")
            return start_row + 1
    
    def _create_monthly_breakdown_table(self, ws, df, start_row):
        """
        Create the monthly breakdown table for day analysis.
        """
        import calendar
        import pandas as pd
        from datetime import datetime
        
        # Section title
        ws.cell(row=start_row, column=1, value="Day Analysis - Monthly Breakdown (Collection Days Only, Non-Outliers)")
        self.formatter.apply_title_style(ws, f'A{start_row}')
        start_row += 2
        
        # Headers
        headers = ['Month', 'School Year', 'Collection Days', 'Days with Data', 'Zero Days', 'Coverage %', 'Avg Files/Day']
        for col, header in enumerate(headers, 1):
            ws.cell(row=start_row, column=col, value=header)
        self.formatter.apply_header_style(ws, f'A{start_row}:G{start_row}')
        start_row += 1
        
        # Group by month and school year
        monthly_data = []
        for month in sorted(df['Month'].unique()):
            month_df = df[df['Month'] == month].copy()
            if not month_df.empty:
                school_year = month_df['School_Year'].iloc[0]
                unique_dates = month_df['Date'].dt.date.unique()
                days_with_data = len(unique_dates)
                avg_files_per_day = month_df['Total_Files'].mean()
                
                # Calculate collection days for this month (approximate)
                month_date = pd.to_datetime(month + '-01')
                month_name = month_date.strftime('%b %Y')
                
                # Estimate collection days in month (weekdays excluding holidays)
                year = month_date.year
                month_num = month_date.month
                days_in_month = calendar.monthrange(year, month_num)[1]
                
                # Count weekdays in month (rough estimate)
                weekdays_in_month = 0
                for day in range(1, days_in_month + 1):
                    date_obj = datetime(year, month_num, day)
                    if date_obj.weekday() < 5:  # Monday to Friday
                        weekdays_in_month += 1
                
                # Use actual collection days from data or estimate
                collection_days_estimate = min(weekdays_in_month, 23)  # Max ~23 collection days per month
                days_with_zero = max(0, collection_days_estimate - days_with_data)
                coverage_pct = (days_with_data / collection_days_estimate * 100) if collection_days_estimate > 0 else 0
                
                monthly_data.append({
                    'month': month_name,
                    'school_year': school_year,
                    'collection_days': collection_days_estimate,
                    'days_with_data': days_with_data,
                    'days_with_zero': days_with_zero,
                    'coverage_pct': coverage_pct,
                    'avg_files_per_day': avg_files_per_day
                })
        
        # Fill monthly data
        for row_offset, month_data in enumerate(monthly_data):
            row = start_row + row_offset
            ws.cell(row=row, column=1, value=month_data['month'])
            ws.cell(row=row, column=2, value=month_data['school_year'])
            ws.cell(row=row, column=3, value=month_data['collection_days'])
            ws.cell(row=row, column=4, value=month_data['days_with_data'])
            ws.cell(row=row, column=5, value=month_data['days_with_zero'])
            # Store coverage percentage as numeric value with Excel formatting
            cell = ws.cell(row=row, column=6)
            cell.value = month_data['coverage_pct'] / 100  # Convert percentage to decimal for Excel
            cell.number_format = '0.0%'  # Apply Excel percentage formatting
            ws.cell(row=row, column=7, value=round(month_data['avg_files_per_day'], 1))
        
        # Apply formatting
        end_row = start_row + len(monthly_data) - 1
        self.formatter.apply_data_style(ws, f'A{start_row}:G{end_row}')
        self.formatter.apply_alternating_row_colors(ws, start_row, end_row, 1, 7)
        
        return end_row + 1
    
    def get_available_methods(self):
        """
        Returns a list of all available sheet creation methods.
        
        Returns:
            dict: Dictionary of method categories and their methods
        """
        return {
            'base_sheets': [
                'create_summary_statistics_sheet',
                'create_raw_data_sheet', 
                'create_data_cleaning_sheet'
            ],
            'pipeline_sheets': [
                'process_pipeline_configurations'
            ],
            'specialized_sheets': [
                'create_audio_efficiency_details_sheet',
                'create_mp3_duration_analysis_sheet'
            ],
            'utility_methods': [
                '_fill_missing_collection_days',
                '_run_aggregation',
                '_should_apply_forecasting',
                '_apply_arima_forecasting'
            ]
        }
    
    def create_all_sheets(self, workbook):
        """
        Creates all configured sheets in the workbook.
        
        Args:
            workbook: openpyxl workbook object
        """
        try:
            print("[INFO] Starting comprehensive sheet creation...")
            
            # Create base sheets
            print("[INFO] Creating base sheets...")
            self.create_summary_statistics_sheet(workbook)
            # Pass a deep copy of the data to prevent modifications from leaking between sheets
            self.create_raw_data_sheet(workbook, data.copy(deep=True))
            self.create_data_cleaning_sheet(workbook)
            
            # Create pipeline-driven sheets
            print("[INFO] Creating pipeline-driven sheets...")
            self.process_pipeline_configurations(workbook)
            
            # Create specialized sheets
            print("[INFO] Creating specialized analysis sheets...")
            self.create_audio_efficiency_details_sheet(workbook)
            self.create_mp3_duration_analysis_sheet(workbook)
            
            print("[SUCCESS] All sheets created successfully")
            
        except Exception as e:
            print(f"[ERROR] Failed to create all sheets: {e}")
            raise


# Export the main class for backward compatibility
__all__ = ['SheetCreator', 'BaseSheetCreator', 'PipelineSheetCreator', 'SpecializedSheetCreator']
