"""
Base Sheet Creator Module
========================

This module contains the base SheetCreator class with common functionality
and core sheet creation methods for summary, raw data, and data cleaning sheets.
"""

import json
import pandas as pd
from pathlib import Path
from datetime import datetime

# Local imports
from utils import (
    get_school_calendar, get_non_collection_days, precompute_collection_days,
    add_acf_pacf_analysis, infer_sheet_type, reorder_with_acf_pacf
)
from pipelines import PIPELINES  # Now using modular pipelines/ package

# Import db_utils conditionally to avoid import errors
try:
    from db_utils import get_db_connection
except ImportError:
    get_db_connection = None


class BaseSheetCreator:
    """
    Base class for sheet creation with common functionality and core sheet methods.
    """
    
    def __init__(self, db, formatter):
        """
        Initialize the sheet creator.
        
        Args:
            db: MongoDB database connection
            formatter: ExcelFormatter instance for styling
        """
        self.db = db
        self.formatter = formatter
        # Global pipeline cache to prevent duplicate executions
        self._pipeline_cache = {}
    
    def _fill_missing_collection_days(self, df, pipeline_name):
        """
        Fills in missing collection days with zero counts for complete time series.
        This is critical for ACF/PACF/ARIMA analysis which requires continuous time series.
        
        Args:
            df (pandas.DataFrame): DataFrame from daily aggregation
            pipeline_name (str): Name of the pipeline to determine if zero-fill is needed
            
        Returns:
            pandas.DataFrame: DataFrame with missing collection days filled with zeros
        """
        if not ('DAILY' in pipeline_name.upper() and 
                ('WITH_ZEROES' in pipeline_name.upper() or 'COLLECTION_ONLY' in pipeline_name.upper())):
            return df
        
        try:
            school_calendar = get_school_calendar()
            non_collection_days = get_non_collection_days()
            collection_day_map = precompute_collection_days(school_calendar, non_collection_days)
            
            all_collection_days = []
            for date_obj, info in collection_day_map.items():
                all_collection_days.append({'_id': date_obj.strftime('%Y-%m-%d')})
            
            all_days_df = pd.DataFrame(all_collection_days)
            
            # Merge with existing data, keeping actual data and filling missing with zeros
            merged_df = pd.merge(all_days_df, df, on='_id', how='left').fillna(0)
            
            # Ensure correct data types after merge
            for col in ['Total_Files', 'MP3_Files', 'JPG_Files']:
                if col in merged_df.columns:
                    merged_df[col] = merged_df[col].astype(int)
            if 'Total_Size_MB' in merged_df.columns:
                 merged_df['Total_Size_MB'] = merged_df['Total_Size_MB'].astype(float)

            return merged_df.sort_values('_id').reset_index(drop=True)
            
        except Exception as e:
            print(f"[WARNING] Zero-fill failed, returning original data: {e}")
            return df
    
    def _run_aggregation_cached(self, cache_key, pipeline, use_base_filter=True, collection_name='media_records'):
        """
        Runs a MongoDB aggregation pipeline with caching to prevent duplicate executions.
        """
        # Check cache first
        if cache_key in self._pipeline_cache:
            print(f"[CACHE HIT] BaseSheetCreator: Reusing cached result for {cache_key}")
            return self._pipeline_cache[cache_key].copy()
        
        # Execute pipeline and cache result
        print(f"[CACHE MISS] BaseSheetCreator: Executing and caching {cache_key}")
        result = self._run_aggregation_original(pipeline, use_base_filter, collection_name)
        self._pipeline_cache[cache_key] = result.copy()
        return result
    
    def _run_aggregation_original(self, pipeline, use_base_filter=True, collection_name='media_records'):
        """
        Original pipeline execution method (renamed to avoid conflicts).
        """
        try:
            collection = self.db[collection_name]
            full_pipeline = pipeline
            if use_base_filter:
                base_filter = {"$match": {"file_type": {"$in": ["JPG", "MP3"]}}}
                full_pipeline = [base_filter] + pipeline
            
            cursor = collection.aggregate(full_pipeline, allowDiskUse=True)
            df = pd.DataFrame(list(cursor))
            
            if '_id' in df.columns and df.shape[0] > 0 and isinstance(df['_id'].iloc[0], dict):
                # Flatten the _id dictionary into separate columns
                id_df = pd.json_normalize(df['_id'])
                df = pd.concat([df.drop('_id', axis=1), id_df], axis=1)
            
            return df
        except Exception as e:
            print(f"[ERROR] Aggregation failed: {e}")
            return pd.DataFrame()
    
    def _run_aggregation(self, pipeline, use_base_filter=True, collection_name='media_records'):
        """
        Runs a MongoDB aggregation pipeline and returns a DataFrame.
        Now uses caching to prevent duplicate executions.
        """
        # Create cache key from pipeline and parameters
        cache_key = f"base_{str(pipeline)}_{use_base_filter}_{collection_name}"
        return self._run_aggregation_cached(cache_key, pipeline, use_base_filter, collection_name)

    def create_summary_statistics_sheet(self, workbook):
        """
        Creates the detailed Summary Statistics sheet.
        """
        try:
            # Get the data for each school year
            pipeline = [
                {"$match": {"file_type": {"$in": ["JPG", "MP3"]}}},
                {"$group": {
                    "_id": "$ISO_Date",
                    "Total_Files": {"$sum": 1},
                    "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
                    "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
                    "Total_Size_MB": {"$sum": "$File_Size_MB"}
                }},
                {"$sort": {"_id": 1}}
            ]
            
            df = self._run_aggregation(pipeline)
            if df.empty:
                print("[WARNING] No data found for Summary Statistics")
                return
            
            # Convert _id to datetime for filtering
            df['Date'] = pd.to_datetime(df['_id'])
            
            # Define school year periods
            periods = {
                '2021-2022': (datetime(2021, 8, 1), datetime(2022, 6, 30)),
                '2022-2023': (datetime(2022, 8, 1), datetime(2023, 6, 30)),
                'Overall': (datetime(2021, 8, 1), datetime(2023, 6, 30))
            }
            
            # Calculate statistics for each period
            def calc_stats_for_period(df, school_year=None):
                """Calculate statistics for a specific period"""
                if school_year and school_year in periods and school_year != 'Overall':
                    start_date, end_date = periods[school_year]
                    period_df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)].copy()
                else:
                    period_df = df.copy()
                
                if period_df.empty:
                    return {}
                
                # Get non-collection days and convert to strings for filtering
                non_collection_days = get_non_collection_days()
                non_collection_day_strings = []
                for day in non_collection_days:
                    if hasattr(day, 'strftime'):
                        non_collection_day_strings.append(day.strftime('%Y-%m-%d'))
                    else:
                        non_collection_day_strings.append(str(day))
                
                # Filter out non-collection days for statistics calculation
                collection_days_df = period_df[~period_df['_id'].isin(non_collection_day_strings)]
                
                if collection_days_df.empty:
                    return {}
                
                # Calculate collection days using the same logic as Period Counts
                if school_year and school_year != 'Overall':
                    from utils.calendar import calculate_collection_days_for_period
                    collection_days = calculate_collection_days_for_period(school_year)
                else:
                    # For overall, sum both periods
                    from utils.calendar import calculate_collection_days_for_period
                    collection_days = (calculate_collection_days_for_period('2021-2022') + 
                                     calculate_collection_days_for_period('2022-2023'))
                
                stats = {
                    'total_files': collection_days_df['Total_Files'].sum(),
                    'mp3_files': collection_days_df['MP3_Files'].sum(),
                    'jpg_files': collection_days_df['JPG_Files'].sum(),
                    'total_size_mb': collection_days_df['Total_Size_MB'].sum(),
                    'collection_days': collection_days,
                    'mean_files_per_day': collection_days_df['Total_Files'].mean(),
                    'median_files_per_day': collection_days_df['Total_Files'].median(),
                    'std_files_per_day': collection_days_df['Total_Files'].std(),
                    'min_files_per_day': collection_days_df['Total_Files'].min(),
                    'max_files_per_day': collection_days_df['Total_Files'].max(),
                    'range_files_per_day': collection_days_df['Total_Files'].max() - collection_days_df['Total_Files'].min()
                }
                
                return stats
            
            # Calculate statistics for each period
            stats_2122 = calc_stats_for_period(df, '2021-2022')
            stats_2223 = calc_stats_for_period(df, '2022-2023')
            stats_overall = calc_stats_for_period(df, 'Overall')
            
            # Create the worksheet
            ws = workbook.create_sheet("Summary Statistics")
            
            # Title
            ws['A1'] = "AR Data Analysis - Summary Statistics"
            self.formatter.apply_title_style(ws, 'A1')
            
            # Headers
            headers = ['Metric', '2021-2022', '2022-2023', 'Overall']
            for col, header in enumerate(headers, 1):
                ws.cell(row=3, column=col, value=header)
            self.formatter.apply_header_style(ws, f'A3:D3')
            
            # Data rows
            metrics = [
                ('Total Files', 'total_files'),
                ('MP3 Files', 'mp3_files'),
                ('JPG Files', 'jpg_files'),
                ('Total Size (MB)', 'total_size_mb'),
                ('Collection Days', 'collection_days'),
                ('Mean Files per Day', 'mean_files_per_day'),
                ('Median Files per Day', 'median_files_per_day'),
                ('Std Dev Files per Day', 'std_files_per_day'),
                ('Min Files per Day', 'min_files_per_day'),
                ('Max Files per Day', 'max_files_per_day'),
                ('Range Files per Day', 'range_files_per_day')
            ]
            
            for row, (label, key) in enumerate(metrics, 4):
                ws.cell(row=row, column=1, value=label)
                
                # Format values appropriately
                for col, stats in enumerate([stats_2122, stats_2223, stats_overall], 2):
                    if key in stats:
                        value = stats[key]
                        if key in ['mean_files_per_day', 'median_files_per_day', 'std_files_per_day']:
                            ws.cell(row=row, column=col, value=round(value, 1))
                        elif key == 'total_size_mb':
                            ws.cell(row=row, column=col, value=round(value, 2))
                        else:
                            ws.cell(row=row, column=col, value=value)
                    else:
                        ws.cell(row=row, column=col, value='N/A')
            
            # Apply formatting
            self.formatter.apply_data_style(ws, f'A4:D{3 + len(metrics)}')
            self.formatter.auto_adjust_columns(ws)
            
            print("[SUCCESS] Summary Statistics sheet created")
            
        except Exception as e:
            print(f"[ERROR] Failed to create Summary Statistics sheet: {e}")

    def create_raw_data_sheet(self, workbook):
        """
        Creates the Raw Data sheet.
        """
        try:
            # Get all records
            collection = self.db['media_records']
            cursor = collection.find({"file_type": {"$in": ["JPG", "MP3"]}}).limit(10000)
            df = pd.DataFrame(list(cursor))
            
            if df.empty:
                print("[WARNING] No data found for Raw Data sheet")
                return
            
            # Create worksheet
            ws = workbook.create_sheet("Raw Data")
            
            # Add title
            ws['A1'] = "AR Data Analysis - Raw Data Sample"
            self.formatter.apply_title_style(ws, 'A1')
            
            # Add data starting from row 3
            for col, column_name in enumerate(df.columns, 1):
                ws.cell(row=3, column=col, value=str(column_name))
            
            # Apply header formatting
            self.formatter.apply_header_style(ws, f'A3:{chr(64 + len(df.columns))}3')
            
            # Add data rows
            for row_idx, (_, row) in enumerate(df.iterrows(), 4):
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=str(value))
            
            # Apply data formatting
            self.formatter.apply_data_style(ws, f'A4:{chr(64 + len(df.columns))}{3 + len(df)}')
            self.formatter.auto_adjust_columns(ws)
            
            print(f"[SUCCESS] Raw Data sheet created with {len(df)} records")
            
        except Exception as e:
            print(f"[ERROR] Failed to create Raw Data sheet: {e}")

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
                ('Data Quality Rate', f"{(valid_records/total_records*100):.1f}%" if total_records > 0 else "N/A")
            ]
            
            metrics_row = row + 4
            for i, (metric, value) in enumerate(quality_metrics):
                ws.cell(row=metrics_row + i, column=1, value=metric)
                ws.cell(row=metrics_row + i, column=2, value=value)
            
            # Apply formatting
            self.formatter.apply_data_style(ws, f'A{metrics_row}:B{metrics_row + len(quality_metrics) - 1}')
            self.formatter.auto_adjust_columns(ws)
            
            print("[SUCCESS] Data Cleaning sheet created")
            
        except Exception as e:
            print(f"[ERROR] Failed to create Data Cleaning sheet: {e}")
