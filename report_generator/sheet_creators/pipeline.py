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

# Local imports - CRITICAL: Use explicit imports to avoid namespace collision
# The utils package also exports add_acf_pacf_analysis (problematic version)
# We must ensure ONLY the ar_utils.py version (correct) is used
from ar_utils import add_acf_pacf_analysis, reorder_with_acf_pacf, infer_sheet_type
from column_cleanup_utils import cleanup_duplicate_acf_pacf_columns
from utils.formatting import reorder_with_forecast_columns  # Explicit submodule import
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
                    # CRITICAL FIX: Clear pipeline cache before each sheet to prevent shared state contamination
                    print(f"[DEBUG] Clearing pipeline cache before creating sheet: {sheet_config['name']}")
                    self._pipeline_cache.clear()
                    
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
        
        # CRITICAL FIX: Use non-cached aggregation to prevent DataFrame contamination
        # The caching mechanism was storing DataFrames that had been mutated with ACF/PACF columns
        # This caused each subsequent sheet to inherit previously added columns
        print(f"[DEBUG] Using fresh (non-cached) pipeline execution to prevent column contamination")
        df = self._run_aggregation_original(pipeline)
        print(f"    - Fresh pipeline data: {len(df)} rows × {len(df.columns)} columns")
        
        # ADDITIONAL SAFETY: Verify no ACF/PACF columns exist in fresh data
        acf_pacf_patterns = ['ACF_Lag_', 'PACF_Lag_', '_Significant']
        existing_acf_pacf_cols = [col for col in df.columns 
                                if any(pattern in str(col) for pattern in acf_pacf_patterns)]
        
        if existing_acf_pacf_cols:
            print(f"[CRITICAL ERROR] Fresh pipeline data already contains ACF/PACF columns: {existing_acf_pacf_cols}")
            print(f"[CRITICAL ERROR] This indicates contamination in the MongoDB aggregation pipeline itself")
            # Remove them as emergency fallback
            df = df.drop(columns=existing_acf_pacf_cols)
            print(f"[EMERGENCY FIX] Removed contaminated columns, proceeding with clean data")
        
        if df.empty:
            print(f"[WARNING] No data returned for pipeline '{pipeline_name}'")
            return
        
        # Apply zero-fill for daily pipelines if needed
        df = self._fill_missing_collection_days(df, pipeline_name)
        
        # Determine sheet type and apply appropriate analysis
        sheet_type = infer_sheet_type(sheet_name)
        
        # Apply ACF/PACF analysis based on configuration (not just sheet type)
        from chart_config_helper import should_add_acf_pacf_columns
        
        if should_add_acf_pacf_columns(sheet_name):
            print(f"[INFO] Adding ACF/PACF analysis for {sheet_type} sheet")
            original_columns = df.columns.tolist()
            
            # Match legacy behavior: only analyze Total_Files metric
            if 'Total_Files' in df.columns:
                print(f"    - Analyzing metric: Total_Files (legacy behavior)")
                print(f"    - Before ACF/PACF: {len(df.columns)} columns")
                print(f"    - Existing ACF columns: {[col for col in df.columns if 'ACF_Lag_' in str(col)]}")
                
                # Apply ACF/PACF analysis - FIXED: Proper column deduplication
                # CRITICAL: Ensure we start with a completely fresh DataFrame copy
                df_for_analysis = df.copy(deep=True)
                
                # DEBUG: Log DataFrame state before ACF/PACF analysis
                print(f"    - [DEBUG] DataFrame before ACF/PACF analysis:")
                print(f"      - Shape: {df.shape}")
                print(f"      - Columns: {list(df.columns)}")
                print(f"      - ACF columns already present: {[col for col in df.columns if 'ACF_Lag_' in str(col)]}")
                print(f"      - PACF columns already present: {[col for col in df.columns if 'PACF_Lag_' in str(col)]}")
                
                # Get ACF/PACF results (returns ONLY new columns)
                print(f"    - [DEBUG] Calling add_acf_pacf_analysis...")
                acf_pacf_results = add_acf_pacf_analysis(df_for_analysis, value_col='Total_Files', sheet_type=sheet_type)
                
                # DEBUG: Log ACF/PACF results
                print(f"    - [DEBUG] ACF/PACF analysis results:")
                print(f"      - Shape: {acf_pacf_results.shape}")
                print(f"      - Columns: {list(acf_pacf_results.columns)}")
                
                # CRITICAL DEBUG: Check for PACF columns specifically
                acf_cols_generated = [col for col in acf_pacf_results.columns if 'ACF_Lag_' in col and '_Significant' not in col]
                pacf_cols_generated = [col for col in acf_pacf_results.columns if 'PACF_Lag_' in col and '_Significant' not in col]
                print(f"      - ACF columns generated: {acf_cols_generated}")
                print(f"      - PACF columns generated: {pacf_cols_generated}")
                
                if not pacf_cols_generated:
                    print(f"    - [CRITICAL ERROR] NO PACF COLUMNS GENERATED BY add_acf_pacf_analysis!")
                    print(f"    - [CRITICAL ERROR] This indicates the PACF calculation is failing in the main pipeline")
                    
                if not acf_pacf_results.empty:
                    print(f"    - ACF/PACF analysis returned {len(acf_pacf_results.columns)} new columns")
                    
                    # VERIFICATION: Confirm no ACF/PACF columns exist before adding new ones
                    # With the fix above, this should never trigger, but kept as safety check
                    acf_pacf_patterns = ['ACF_Lag_', 'PACF_Lag_', '_Significant']
                    existing_acf_pacf_cols = [col for col in df.columns 
                                            if any(pattern in str(col) for pattern in acf_pacf_patterns)]
                    
                    if existing_acf_pacf_cols:
                        print(f"    - WARNING: Found unexpected ACF/PACF columns (should not happen with fix): {existing_acf_pacf_cols}")
                        print(f"    - REMOVING them as safety measure")
                        df = df.drop(columns=existing_acf_pacf_cols)
                    
                    # Now safely concatenate the new ACF/PACF columns
                    # Perform concatenation
                    df = pd.concat([df, acf_pacf_results], axis=1)
                    
                    # CRITICAL DEBUG: Check PACF columns after concatenation
                    acf_cols_after_concat = [col for col in df.columns if 'ACF_Lag_' in col and '_Significant' not in col]
                    pacf_cols_after_concat = [col for col in df.columns if 'PACF_Lag_' in col and '_Significant' not in col]
                    print(f"    - [DEBUG] After concatenation:")
                    print(f"      - ACF columns: {acf_cols_after_concat}")
                    print(f"      - PACF columns: {pacf_cols_after_concat}")
                    
                    if not pacf_cols_after_concat:
                        print(f"    - [CRITICAL ERROR] PACF COLUMNS LOST DURING CONCATENATION!")
                    
                    # Check for duplicates immediately after concatenation
                    if len(df.columns) != len(set(df.columns)):
                        print(f"    - WARNING: Duplicate columns detected after concatenation, removing...")
                        df = df.loc[:, ~df.columns.duplicated()]
                    
                    # Clean up any duplicate ACF/PACF columns from inconsistent naming
                    df = cleanup_duplicate_acf_pacf_columns(df, 'Total_Files')
                    
                    print(f"    - After ACF/PACF: {len(df.columns)} columns")
                    print(f"    - New ACF columns: {[col for col in df.columns if 'ACF_Lag_' in str(col)]}")
                    print(f"    - New PACF columns: {[col for col in df.columns if 'PACF_Lag_' in str(col)]}")
                    
                    # CRITICAL DEBUG: Final PACF check after all processing
                    final_acf_cols = [col for col in df.columns if 'ACF_Lag_' in col and '_Significant' not in col]
                    final_pacf_cols = [col for col in df.columns if 'PACF_Lag_' in col and '_Significant' not in col]
                    print(f"    - [FINAL DEBUG] After all processing:")
                    print(f"      - Final ACF columns: {final_acf_cols}")
                    print(f"      - Final PACF columns: {final_pacf_cols}")
                    
                    if not final_pacf_cols:
                        print(f"    - [CRITICAL ERROR] PACF COLUMNS COMPLETELY MISSING FROM FINAL DATAFRAME!")
                        print(f"    - [CRITICAL ERROR] This will result in missing PACF columns in Excel output")
            else:
                print(f"    - Skipping ACF/PACF analysis: Total_Files column not found")
                
            # Ensure no duplicate columns before reordering
            if len(df.columns) != len(set(df.columns)):
                print(f"    - WARNING: Duplicate columns detected before reordering, removing...")
                df = df.loc[:, ~df.columns.duplicated(keep='first')]
                
            df = reorder_with_acf_pacf(df, original_columns)
            
            # Apply ARIMA forecasting based on configuration (not just sheet type)
            from chart_config_helper import should_add_arima_columns
            
            if should_add_arima_columns(sheet_name):
                print(f"[INFO] Adding ARIMA forecasting for {sheet_type} sheet (configuration-driven)")
                # Store columns before forecasting for reordering
                columns_before_forecast = df.columns.tolist()
                df = self._apply_arima_forecasting(df, sheet_type)
                df = reorder_with_forecast_columns(df, columns_before_forecast)
            else:
                print(f"[INFO] Skipping ARIMA forecasting for {sheet_name} (not enabled in configuration)")
        
        # CRITICAL FIX: Smart deduplication that preserves PACF columns
        print(f"    - [DEBUG] Before deduplication: {len(df.columns)} columns")
        print(f"    - [DEBUG] Columns before dedup: {list(df.columns)}")
        original_col_count = len(df.columns)
        
        # Check for duplicates first
        if len(df.columns) != len(set(df.columns)):
            print(f"    - [CRITICAL] Duplicate columns detected before Excel export!")
            duplicate_cols = [col for col in df.columns if list(df.columns).count(col) > 1]
            print(f"    - [CRITICAL] Duplicate columns: {set(duplicate_cols)}")
            
            # SMART DEDUPLICATION: Remove true duplicates while preserving unique columns
            # Use pandas built-in deduplication but with proper validation
            columns_before = list(df.columns)
            df = df.loc[:, ~df.columns.duplicated(keep='first')]
            columns_after = list(df.columns)
            
            # Log which columns were removed
            removed_columns = [col for col in columns_before if col not in columns_after or columns_before.count(col) > columns_after.count(col)]
            for col in set(removed_columns):
                count_before = columns_before.count(col)
                count_after = columns_after.count(col)
                if count_before > count_after:
                    print(f"    - [DEDUP] Removed {count_before - count_after} duplicate(s) of: {col}")
            
            print(f"    - [FIX] After smart deduplication: {len(df.columns)} columns")
            print(f"    - [FIX] Removed {original_col_count - len(df.columns)} duplicate columns")
            print(f"    - [DEBUG] Columns after dedup: {list(df.columns)}")
            
            # VALIDATION: Ensure PACF columns are still present
            pacf_cols_after = [col for col in df.columns if 'PACF_Lag_' in col and '_Significant' not in col]
            acf_cols_after = [col for col in df.columns if 'ACF_Lag_' in col and '_Significant' not in col]
            print(f"    - [VALIDATION] ACF columns after dedup: {acf_cols_after}")
            print(f"    - [VALIDATION] PACF columns after dedup: {pacf_cols_after}")
            
            if not pacf_cols_after and acf_cols_after:
                print(f"    - [CRITICAL ERROR] PACF columns lost during deduplication!")
                print(f"    - [CRITICAL ERROR] This will cause chart overlap issue!")
            elif pacf_cols_after:
                print(f"    - [SUCCESS] PACF columns preserved after deduplication")
        else:
            print(f"    - [OK] No duplicate columns found before Excel export")
        
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
            
            # Apply alternating row colors for better readability
            self.formatter.apply_alternating_row_colors(ws, 4, 3 + len(df), 1, len(df.columns))
        else:
            # For smaller datasets, use full formatting
            self.formatter.apply_data_style(ws, data_range)
            
            # Apply alternating row colors for better readability
            self.formatter.apply_alternating_row_colors(ws, 4, 3 + len(df), 1, len(df.columns))
        
        # Apply special formatting for ACF/PACF data
        self._apply_acf_pacf_data_formatting(ws, df.columns, 4, len(df))
        
        # Auto-adjust column widths
        self.formatter.auto_adjust_columns(ws)
        
        # Add ACF/PACF charts based on configuration
        from chart_config_helper import should_add_chart
        if sheet_type in ['daily', 'weekly', 'biweekly', 'monthly', 'period'] and should_add_chart(sheet_name, 'acf_pacf'):
            chart_added = self._add_acf_pacf_charts(ws, 4, 3 + len(df), sheet_type)
            if chart_added:
                print(f"[SUCCESS] Added ACF/PACF charts to sheet '{sheet_name}' (config-driven)")
        elif sheet_type in ['daily', 'weekly', 'biweekly', 'monthly', 'period']:
            print(f"[INFO] Skipping ACF/PACF charts for '{sheet_name}' - not enabled in configuration")
        
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
        # ARIMA forecasting is now enabled - use ar_utils implementation
        from ar_utils import add_arima_forecast_columns
        return add_arima_forecast_columns(df, "Total_Files", "daily")
    
    def _apply_acf_pacf_header_formatting(self, ws, columns, header_row):
        """
        Applies special formatting to ACF/PACF column headers with enhanced readability.
        
        Args:
            ws: Worksheet object
            columns: List of column names
            header_row: Row number containing headers
        """
        try:
            from openpyxl.styles import PatternFill, Font
            
            # Enhanced ACF/PACF header styling for better readability and contrast
            # Darker background for better contrast
            acf_pacf_fill = PatternFill(start_color="B8D4F0", end_color="B8D4F0", fill_type="solid")
            
            # Dark blue font for excellent contrast against light blue background
            acf_pacf_font = Font(color="003366", bold=True, size=10)
            
            for col_idx, column_name in enumerate(columns, 1):
                if any(pattern in str(column_name) for pattern in ['ACF_', 'PACF_']):
                    cell = ws.cell(row=header_row, column=col_idx)
                    cell.fill = acf_pacf_fill
                    cell.font = acf_pacf_font
                    
        except Exception as e:
            print(f"[WARNING] Could not apply ACF/PACF header formatting: {e}")
    
    def _apply_acf_pacf_data_formatting(self, ws, columns, start_row, num_rows):
        """
        Applies special formatting to ACF/PACF data cells with enhanced readability.
        
        Args:
            ws: Worksheet object
            columns: List of column names
            start_row: First data row
            num_rows: Number of data rows
        """
        try:
            from openpyxl.styles import PatternFill, Font
            
            # Enhanced ACF/PACF data styling for better readability
            # Subtle background that complements the header styling
            acf_pacf_fill = PatternFill(start_color="F0F6FC", end_color="F0F6FC", fill_type="solid")
            
            # Slightly darker font for better readability
            acf_pacf_font = Font(color="1F4E79", size=9)
            
            for col_idx, column_name in enumerate(columns, 1):
                if any(pattern in str(column_name) for pattern in ['ACF_', 'PACF_']):
                    for row_idx in range(start_row, start_row + num_rows):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.fill = acf_pacf_fill
                        cell.font = acf_pacf_font
                        
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

