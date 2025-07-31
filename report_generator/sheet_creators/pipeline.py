"""
Pipeline-Driven Sheet Creator Module
===================================

This module handles the creation of Excel sheets based on pipeline configurations
from report_config.json. It processes all configured pipelines and creates
sheets with ACF/PACF analysis and ARIMA forecasting capabilities.
"""

import json
import pandas as pd
from pathlib import Path
from openpyxl.utils import get_column_letter

# Local imports
from utils import (
    add_acf_pacf_analysis, infer_sheet_type, reorder_with_acf_pacf,
    reorder_with_forecast_columns
)
from pipelines import PIPELINES  # Now using modular pipelines/ package
from .base import BaseSheetCreator


class PipelineSheetCreator(BaseSheetCreator):
    """
    Handles creation of sheets based on pipeline configurations with advanced
    time series analysis capabilities.
    """
    
    def process_pipeline_configurations(self, workbook):
        """
        Processes all pipeline configurations from report_config.json and creates sheets.
        Each sheet fetches its own fresh pipeline data to prevent column duplication.
        
        Args:
            workbook: openpyxl workbook object
        """
        try:
            # Load report configuration
            config_path = Path(__file__).parent.parent.parent / "report_config.json"
            with open(config_path, 'r') as f:
                config = json.load(f)
            
            # Process each sheet configuration
            for sheet_config in config.get('sheets', []):
                if not sheet_config.get('enabled', True):
                    print(f"[SKIP] Sheet '{sheet_config['name']}' is disabled")
                    continue
                
                try:
                    # Each sheet gets its own fresh pipeline data to prevent column accumulation
                    self._create_pipeline_sheet(workbook, sheet_config)
                except Exception as e:
                    print(f"[ERROR] Failed to create sheet '{sheet_config['name']}': {e}")
            
            print(f"[SUCCESS] Processed {len([s for s in config.get('sheets', []) if s.get('enabled', True)])} pipeline sheets")
            
        except Exception as e:
            print(f"[ERROR] Failed to process pipeline configurations: {e}")
    
    def _create_pipeline_sheet(self, workbook, sheet_config, data=None):
        """
        Creates a single sheet based on pipeline configuration.
        
        Args:
            workbook: openpyxl workbook object
            sheet_config: Sheet configuration dictionary
            data: Optional pre-fetched data (for backward compatibility)
        """
        sheet_name = sheet_config['name']
        pipeline_name = sheet_config['pipeline']
        
        # Each sheet must use its own fresh pipeline data to prevent column accumulation
        print(f"[INFO] Creating sheet: {sheet_name}")
        print(f"    - Using pipeline: {pipeline_name}")
        
        # Get fresh data from the specific pipeline for this sheet
        from pipelines import PIPELINES
        if pipeline_name not in PIPELINES:
            print(f"[ERROR] Pipeline '{pipeline_name}' not found")
            return
        
        pipeline = PIPELINES[pipeline_name]
        df = self._run_aggregation(pipeline)
        print(f"    - Fresh pipeline data: {len(df)} rows × {len(df.columns)} columns")
        
        if df.empty:
            print(f"[WARNING] No data returned for pipeline '{pipeline_name}'")
            return
        
        # Apply zero-fill for daily pipelines if needed
        df = self._fill_missing_collection_days(df, pipeline_name)
        
        # Determine sheet type and apply appropriate analysis
        sheet_type = infer_sheet_type(sheet_name)
        
        # Apply ACF/PACF analysis for time series sheets
        if sheet_type in ['daily', 'weekly', 'biweekly', 'monthly', 'period']:
            print(f"[INFO] Adding ACF/PACF analysis for {sheet_type} sheet")
            original_columns = df.columns.tolist()
            
            # Match legacy behavior: only analyze Total_Files metric
            if 'Total_Files' in df.columns:
                print(f"    - Analyzing metric: Total_Files (legacy behavior)")
                print(f"    - Before ACF/PACF: {len(df.columns)} columns")
                print(f"    - Existing ACF columns: {[col for col in df.columns if 'ACF_Lag_' in str(col)]}")
                
                # Apply ACF/PACF analysis - get new columns and concatenate properly
                acf_pacf_results = add_acf_pacf_analysis(df, value_col='Total_Files', sheet_type=sheet_type)
                if not acf_pacf_results.empty:
                    # Check for overlapping columns before concatenation
                    overlapping_cols = set(df.columns) & set(acf_pacf_results.columns)
                    if overlapping_cols:
                        print(f"    - WARNING: Overlapping columns detected: {overlapping_cols}")
                        # Remove overlapping columns from ACF/PACF results to prevent duplication
                        acf_pacf_results = acf_pacf_results.drop(columns=list(overlapping_cols))
                    
                    df = pd.concat([df, acf_pacf_results], axis=1)
                    
                    # Ensure no duplicate columns after concatenation
                    original_col_count = len(df.columns)
                    seen_columns = set()
                    unique_columns = []
                    for col in df.columns:
                        if col not in seen_columns:
                            unique_columns.append(col)
                            seen_columns.add(col)
                    
                    if len(unique_columns) != original_col_count:
                        print(f"    - DEDUPLICATION: Removed {original_col_count - len(unique_columns)} duplicate columns")
                        df = df[unique_columns]
                
                print(f"    - After ACF/PACF: {len(df.columns)} columns")
                print(f"    - New ACF columns: {[col for col in df.columns if 'ACF_Lag_' in str(col)]}")
            else:
                print(f"    - Skipping ACF/PACF analysis: Total_Files column not found")
                
            df = reorder_with_acf_pacf(df, original_columns)
            
            # Apply ARIMA forecasting if enabled
            if self._should_apply_forecasting(sheet_config, sheet_type):
                print(f"[INFO] Adding ARIMA forecasting for {sheet_type} sheet")
                # Store columns before forecasting for reordering
                columns_before_forecast = df.columns.tolist()
                df = self._apply_arima_forecasting(df, sheet_type)
                df = reorder_with_forecast_columns(df, columns_before_forecast)
        
        # CRITICAL FIX: Ensure no duplicate columns before writing to Excel
        print(f"    - Before deduplication: {len(df.columns)} columns")
        original_col_count = len(df.columns)
        
        # Remove duplicate columns while preserving order
        seen_columns = set()
        unique_columns = []
        for col in df.columns:
            if col not in seen_columns:
                unique_columns.append(col)
                seen_columns.add(col)
        
        if len(unique_columns) != original_col_count:
            print(f"    - DEDUPLICATION: Removed {original_col_count - len(unique_columns)} duplicate columns")
            df = df[unique_columns]
            print(f"    - After deduplication: {len(df.columns)} columns")
        else:
            print(f"    - No duplicate columns found")
        
        # Create the worksheet
        ws = workbook.create_sheet(sheet_name)
        
        # Add title
        ws['A1'] = f"AR Data Analysis - {sheet_name}"
        self.formatter.apply_title_style(ws, 'A1')
        
        # Add headers
        for col, column_name in enumerate(df.columns, 1):
            ws.cell(row=3, column=col, value=str(column_name))
        
        # Apply header formatting with special styling for ACF/PACF columns
        last_col_letter = get_column_letter(len(df.columns))
        header_range = f'A3:{last_col_letter}3'
        self.formatter.apply_header_style(ws, header_range)
        
        # Apply special formatting for ACF/PACF columns
        self._apply_acf_pacf_header_formatting(ws, df.columns, 3)
        
        # Add data rows
        for row_idx, (_, row) in enumerate(df.iterrows(), 4):
            for col_idx, value in enumerate(row, 1):
                # Handle different data types appropriately
                if pd.isna(value):
                    cell_value = ""
                elif isinstance(value, (int, float)):
                    cell_value = value
                else:
                    cell_value = str(value)
                
                ws.cell(row=row_idx, column=col_idx, value=cell_value)
        
        # Apply data formatting with optimization for large datasets
        last_col_letter = get_column_letter(len(df.columns))
        data_range = f'A4:{last_col_letter}{3 + len(df)}'
        
        # Calculate total cells to determine if we need optimized approach
        total_cells = len(df) * len(df.columns)
        
        if total_cells > 1000:
            print(f"[INFO] Large dataset detected ({len(df)} rows × {len(df.columns)} cols = {total_cells} cells)")
            print(f"[INFO] Using optimized formatting approach")
            
            # For large datasets, apply minimal formatting to avoid performance issues
            self._apply_minimal_data_formatting(ws, 4, 3 + len(df), len(df.columns))
        else:
            # For smaller datasets, use full formatting
            self.formatter.apply_data_style(ws, data_range)
        
        # Apply special formatting for ACF/PACF data
        self._apply_acf_pacf_data_formatting(ws, df.columns, 4, len(df))
        
        # Auto-adjust column widths
        self.formatter.auto_adjust_columns(ws)
        
        # Add ACF/PACF charts if this is a time series sheet with ACF/PACF data
        if sheet_type in ['daily', 'weekly', 'biweekly', 'monthly', 'period']:
            chart_added = self._add_acf_pacf_charts(ws, 4, 3 + len(df), sheet_type)
            if chart_added:
                print(f"[SUCCESS] Added ACF/PACF charts to sheet '{sheet_name}'")
        
        print(f"[SUCCESS] Created sheet '{sheet_name}' with {len(df)} rows")
    
    def _add_acf_pacf_charts(self, worksheet, data_start_row, data_end_row, sheet_type):
        """
        Add ACF/PACF charts directly to the worksheet during sheet creation.
        
        Args:
            worksheet: openpyxl worksheet object
            data_start_row: First row of data (after headers)
            data_end_row: Last row of data
            sheet_type: Type of sheet (daily, weekly, etc.)
            
        Returns:
            bool: True if charts were added successfully, False otherwise
        """
        try:
            # Import chart functions from acf_pacf_charts
            from acf_pacf_charts import add_acf_pacf_chart, add_chart_summary_info
            
            # Add ACF/PACF chart directly to the worksheet
            chart = add_acf_pacf_chart(worksheet, data_start_row, data_end_row, sheet_type)
            
            if chart:
                # Count ACF/PACF columns for summary - use more permissive column detection
                headers = [str(cell.value) for cell in worksheet[1] if cell.value is not None]
                
                # More flexible detection of ACF/PACF columns
                acf_cols = [h for h in headers if ('ACF' in h or 'Acf' in h or 'acf' in h) and ('Lag' in h or 'lag' in h) and not ('Significant' in h or 'significant' in h)]
                pacf_cols = [h for h in headers if ('PACF' in h or 'Pacf' in h or 'pacf' in h) and ('Lag' in h or 'lag' in h) and not ('Significant' in h or 'significant' in h)]
                total_lags = len(set([h.split('_Lag')[1].split('_')[0] for h in acf_cols + pacf_cols if '_Lag' in str(h)]))
                computed_lags = len(acf_cols)
                
                # Add summary information
                add_chart_summary_info(worksheet, data_end_row, sheet_type, total_lags, computed_lags)
                
                return True
            else:
                print(f"[INFO] No ACF/PACF charts added to {worksheet.title} - no suitable data found")
                return False
                
        except Exception as e:
            print(f"[WARNING] Could not add ACF/PACF charts to {worksheet.title}: {e}")
            return False
    
    def _should_apply_forecasting(self, sheet_config, sheet_type):
        """
        Determines if ARIMA forecasting should be applied to this sheet.
        
        Args:
            sheet_config: Sheet configuration dictionary
            sheet_type: Inferred sheet type (daily, weekly, etc.)
            
        Returns:
            bool: True if forecasting should be applied
        """
        try:
            # Load forecast configuration
            config_path = Path(__file__).parent.parent.parent / "report_config.json"
            with open(config_path, 'r') as f:
                config = json.load(f)
            
            forecast_options = config.get('forecast_options', {})
            
            # Check if forecasting is globally enabled
            if not forecast_options.get('enabled', False):
                return False
            
            # Check if this time scale is enabled for forecasting
            enabled_scales = forecast_options.get('time_scales', [])
            return sheet_type in enabled_scales
            
        except Exception as e:
            print(f"[WARNING] Could not determine forecasting settings: {e}")
            return False
    
    def _apply_arima_forecasting(self, df, sheet_type):
        """
        Applies ARIMA forecasting to the DataFrame.
        
        Args:
            df: DataFrame with time series data
            sheet_type: Type of time series (daily, weekly, etc.)
            
        Returns:
            DataFrame with forecast columns added
        """
        # TEMPORARILY DISABLED FOR MIGRATION TESTING
        print(f"[INFO] ARIMA forecasting temporarily disabled for migration testing")
        
        # Add disabled forecast columns to maintain column structure
        target_metrics = ['Total_Files']  # Default metric
        for metric in target_metrics:
            if metric in df.columns:
                df[f'{metric}_Forecast'] = "<ARIMA Disabled>"
                df[f'{metric}_Forecast_Lower'] = "<ARIMA Disabled>"
                df[f'{metric}_Forecast_Upper'] = "<ARIMA Disabled>"
        
        return df
    
    def _apply_acf_pacf_header_formatting(self, ws, columns, header_row):
        """
        Applies special formatting to ACF/PACF column headers.
        
        Args:
            ws: Worksheet object
            columns: List of column names
            header_row: Row number containing headers
        """
        try:
            from openpyxl.styles import PatternFill
            
            # Define ACF/PACF header fill (light blue)
            acf_pacf_fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
            
            for col_idx, column_name in enumerate(columns, 1):
                if any(pattern in str(column_name) for pattern in ['ACF_', 'PACF_']):
                    cell = ws.cell(row=header_row, column=col_idx)
                    cell.fill = acf_pacf_fill
                    
        except Exception as e:
            print(f"[WARNING] Could not apply ACF/PACF header formatting: {e}")
    
    def _apply_acf_pacf_data_formatting(self, ws, columns, start_row, num_rows):
        """
        Applies special formatting to ACF/PACF data cells.
        
        Args:
            ws: Worksheet object
            columns: List of column names
            start_row: First data row
            num_rows: Number of data rows
        """
        try:
            from openpyxl.styles import PatternFill
            
            # Define ACF/PACF data fill (very light blue)
            acf_pacf_fill = PatternFill(start_color="F5F9FF", end_color="F5F9FF", fill_type="solid")
            
            for col_idx, column_name in enumerate(columns, 1):
                if any(pattern in str(column_name) for pattern in ['ACF_', 'PACF_']):
                    for row_idx in range(start_row, start_row + num_rows):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.fill = acf_pacf_fill
                        
        except Exception as e:
            print(f"[WARNING] Could not apply ACF/PACF data formatting: {e}")
    
    def _apply_minimal_data_formatting(self, ws, start_row, end_row, num_cols):
        """
        Applies minimal formatting optimized for large datasets.
        
        Args:
            ws: Worksheet object
            start_row: First data row
            end_row: Last data row (exclusive)
            num_cols: Number of columns
        """
        try:
            from openpyxl.styles import Alignment
            
            print(f"[INFO] Applying minimal formatting to range: A{start_row}:{get_column_letter(num_cols)}{end_row-1}")
            
            # Only set column widths and basic alignment, avoid cell-by-cell styling
            data_alignment = Alignment(horizontal='left', vertical='center')
            
            # Set column widths efficiently
            for col_num in range(1, num_cols + 1):
                col_letter = get_column_letter(col_num)
                if col_letter not in ws.column_dimensions:
                    ws.column_dimensions[col_letter].width = 12
            
            # Apply alignment only to a sample of cells to establish the pattern
            # openpyxl will often inherit this for similar cells
            sample_rows = min(10, end_row - start_row)  # Format first 10 rows as template
            
            for row_num in range(start_row, start_row + sample_rows):
                for col_num in range(1, min(num_cols + 1, 6)):  # First 5 columns as template
                    try:
                        cell = ws.cell(row=row_num, column=col_num)
                        cell.alignment = data_alignment
                    except Exception:
                        continue
            
            print(f"[INFO] Minimal formatting applied successfully")
            
        except Exception as e:
            print(f"[WARNING] Could not apply minimal data formatting: {e}")
            # Fallback to no formatting rather than crash
            pass
