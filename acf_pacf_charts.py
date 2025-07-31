#!/usr/bin/env python3
"""
ACF/PACF Chart Visualization Utilities

This module provides functions to create professional ACF/PACF charts in Excel
using openpyxl, enhancing the interpretability of time series analysis results.
"""

import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.series import Series, SeriesLabel
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.styles import Font, PatternFill, Alignment
import numpy as np
import math


def add_acf_pacf_chart(worksheet, data_start_row, data_end_row, sheet_type="daily"):
    """
    Add an ACF/PACF line chart to an Excel worksheet.
    
    Args:
        worksheet: openpyxl worksheet object
        data_start_row: First row of data (after headers)
        data_end_row: Last row of data
        sheet_type: Type of sheet (daily, weekly, monthly, etc.)
    
    Returns:
        Chart object that was added to the worksheet
    """
    
    # Find ACF and PACF columns (updated pattern for Total_Files_ACF_Lag_X format)
    headers = [cell.value for cell in worksheet[3]]  # Headers are in row 3
    acf_cols = [(i+1, h) for i, h in enumerate(headers) if str(h).endswith('_ACF_Lag_1') or str(h).endswith('_ACF_Lag_7') or str(h).endswith('_ACF_Lag_14') or '_ACF_Lag_' in str(h)]
    pacf_cols = [(i+1, h) for i, h in enumerate(headers) if str(h).endswith('_PACF_Lag_1') or str(h).endswith('_PACF_Lag_7') or str(h).endswith('_PACF_Lag_14') or '_PACF_Lag_' in str(h)]
    
    if not acf_cols and not pacf_cols:
        print(f"No ACF/PACF columns found in {worksheet.title}")
        return None
    
    # Create line chart
    chart = LineChart()
    chart.title = f"ACF_PACF Analysis - {sheet_type.title()} Total Files"
    chart.style = 10  # Professional style
    chart.height = 10  # Reasonable height
    chart.width = 15   # Good width for readability
    
    # Configure axes
    chart.y_axis.title = "Correlation Coefficient"
    chart.x_axis.title = "Lag"
    chart.y_axis.scaling.min = -1.0
    chart.y_axis.scaling.max = 1.0
    
    # Work with actual column structure: Total_Files_ACF_Lag_X, Total_Files_PACF_Lag_X, etc.
    # Find the first ACF and PACF columns for charting
    acf_col_idx = None
    pacf_col_idx = None
    
    for i, header in enumerate(headers):
        if '_ACF_Lag_' in str(header) and '_Significant' not in str(header) and acf_col_idx is None:
            acf_col_idx = i + 1  # Convert to 1-based indexing
        elif '_PACF_Lag_' in str(header) and '_Significant' not in str(header) and pacf_col_idx is None:
            pacf_col_idx = i + 1  # Convert to 1-based indexing
    
    if acf_col_idx is None and pacf_col_idx is None:
        print(f"No _ACF_Lag_ or _PACF_Lag_ columns found in {worksheet.title}")
        print(f"Available headers: {[str(h) for h in headers if 'ACF' in str(h) or 'PACF' in str(h)]}")
        return None
    
    # Create ACF/PACF correlation plot data
    # Extract ACF/PACF values and create lag-based chart data
    
    # Get the actual ACF/PACF values from the first data row (they're constant across all rows)
    acf_data = []
    pacf_data = []
    lag_numbers = []
    
    # Extract lag numbers and values from column headers (updated for Total_Files_ACF_Lag_X format)
    for col_idx, header in enumerate(headers):
        if '_ACF_Lag_' in str(header) and '_Significant' not in str(header):
            # Extract lag number, handling cases like 'Total_Files_ACF_Lag_1' -> '1'
            lag_part = str(header).split('_Lag_')[1]
            if lag_part.isdigit():
                lag_num = int(lag_part)
                acf_value = worksheet.cell(row=data_start_row, column=col_idx + 1).value
                if isinstance(acf_value, (int, float)):
                    lag_numbers.append(lag_num)
                    acf_data.append(acf_value)
        elif '_PACF_Lag_' in str(header) and '_Significant' not in str(header):
            # Extract lag number, handling cases like 'Total_Files_PACF_Lag_1' -> '1'
            lag_part = str(header).split('_Lag_')[1]
            if lag_part.isdigit():
                lag_num = int(lag_part)
                pacf_value = worksheet.cell(row=data_start_row, column=col_idx + 1).value
                if isinstance(pacf_value, (int, float)):
                    pacf_data.append(pacf_value)
    
    # Create temporary data area for chart (below main data)
    chart_data_start_row = data_end_row + 5
    
    # Write lag numbers as x-axis
    for i, lag in enumerate(sorted(set(lag_numbers))):
        worksheet.cell(row=chart_data_start_row + i, column=1, value=lag)
    
    # Write ACF values
    if acf_data:
        worksheet.cell(row=chart_data_start_row - 1, column=2, value="ACF")
        for i, val in enumerate(acf_data):
            worksheet.cell(row=chart_data_start_row + i, column=2, value=val)
    
    # Write PACF values
    if pacf_data:
        worksheet.cell(row=chart_data_start_row - 1, column=3, value="PACF")
        for i, val in enumerate(pacf_data):
            worksheet.cell(row=chart_data_start_row + i, column=3, value=val)
    
    # Create chart references using the temporary data
    chart_data_end_row = chart_data_start_row + len(lag_numbers) - 1
    x_axis_data = Reference(worksheet, min_col=1, min_row=chart_data_start_row, max_row=chart_data_end_row)
    
    chart_series = []
    if acf_data:
        acf_values = Reference(worksheet, min_col=2, min_row=chart_data_start_row, max_row=chart_data_end_row)
        chart_series.append(('ACF', acf_values, "4F81BD"))  # Blue
    
    if pacf_data:
        pacf_values = Reference(worksheet, min_col=3, min_row=chart_data_start_row, max_row=chart_data_end_row)
        chart_series.append(('PACF', pacf_values, "C0504D"))  # Red
    
    # Calculate confidence intervals for reference
    n_observations = max(data_end_row - data_start_row + 1, 10)  # Minimum 10 for safety
    confidence_level = 1.96 / math.sqrt(n_observations)  # 95% confidence interval
    
    print(f"[CHART] Chart data for {worksheet.title}: {len(chart_series)} series, n~{n_observations}, CI=+/-{confidence_level:.3f}")
    
    if not chart_series:
        print(f"[WARNING] No valid ACF/PACF values found in {worksheet.title}")
        
        # Add user-friendly message for insufficient data
        insufficient_data_row = data_end_row + 3
        worksheet.cell(row=insufficient_data_row, column=1, value="[CHART] ACF/PACF Chart Information")
        worksheet.cell(row=insufficient_data_row, column=1).font = openpyxl.styles.Font(bold=True, size=12)
        
        if "Period" in worksheet.title:
            message = "ACF/PACF chart omitted due to insufficient data (only 6 periods)"
        else:
            message = "ACF/PACF chart omitted due to insufficient statistical data"
            
        worksheet.cell(row=insufficient_data_row + 1, column=1, value=message)
        worksheet.cell(row=insufficient_data_row + 1, column=1).font = openpyxl.styles.Font(italic=True)
        
        print(f"[OK] Added explanatory message to {worksheet.title}")
        return None

    # Add data series to chart with proper styling
    for series_name, series_values, color in chart_series:
        chart.add_data(series_values, titles_from_data=False)
        if len(chart.series) > 0:
            current_series = chart.series[-1]
            # Use SeriesLabel for proper OpenPyXL compatibility
            from openpyxl.chart.series import SeriesLabel
            current_series.title = SeriesLabel(v=series_name)
            
            # Apply styling to the series
            if hasattr(current_series, 'graphicalProperties'):
                current_series.graphicalProperties.line.solidFill = color
                current_series.graphicalProperties.line.width = 25000
                current_series.smooth = True  # Smooth lines for better aesthetics
    
    # Set categories (x-axis labels) - use row numbers as lag indicators
    chart.set_categories(x_axis_data)
    
    # Enhanced chart styling (ChatGPT recommendations)
    chart.legend.position = 'b'  # Bottom legend positioning
    # Note: Gridlines require ChartLines objects in openpyxl, simplified for compatibility
    
    # Position chart to the right of the data
    chart_position = f"F{data_start_row}"
    worksheet.add_chart(chart, chart_position)
    
    print(f"[SUCCESS] Added ACF/PACF chart to {worksheet.title} at {chart_position}")
    return chart


def add_arima_forecast_chart(worksheet, data_start_row, data_end_row, sheet_type="daily"):
    """
    Add an ARIMA forecast chart to an Excel worksheet if forecast columns exist.
    
    Args:
        worksheet: openpyxl worksheet object
        data_start_row: First row of data (after headers)
        data_end_row: Last row of data
        sheet_type: Type of sheet (daily, weekly, monthly, etc.)
    
    Returns:
        Chart object that was added to the worksheet, or None if no forecast data
    """
    
    # Find Total_Files and forecast columns
    headers = [cell.value for cell in worksheet[1]]
    total_files_col = None
    forecast_col = None
    
    for i, header in enumerate(headers):
        if str(header) == 'Total_Files':
            total_files_col = i + 1
        elif str(header) == 'Total_Files_Forecast':
            forecast_col = i + 1
    
    if not total_files_col or not forecast_col:
        print(f"No ARIMA forecast data found in {worksheet.title}")
        return None
    
    # Check if forecast column contains numeric data (not error messages)
    sample_forecast = worksheet.cell(row=data_start_row, column=forecast_col).value
    if not isinstance(sample_forecast, (int, float)):
        print(f"ARIMA forecast contains non-numeric data in {worksheet.title}: {sample_forecast}")
        return None
    
    # Create line chart for ARIMA forecast
    chart = LineChart()
    chart.title = f"ARIMA Forecast vs. Actual - {sheet_type.title()} Total Files"
    chart.style = 10
    chart.height = 10
    chart.width = 15
    
    # Configure axes
    chart.y_axis.title = "File Count"
    chart.x_axis.title = "Time Period"
    
    # Add historical data series
    historical_data = Reference(worksheet, min_col=total_files_col, min_row=1, max_row=data_end_row)
    chart.add_data(historical_data, titles_from_data=True)
    
    # Add forecast data series
    forecast_data = Reference(worksheet, min_col=forecast_col, min_row=1, max_row=data_end_row)
    chart.add_data(forecast_data, titles_from_data=True)
    
    # Position chart to the right of ACF/PACF chart
    chart_position = f"Q{data_end_row + 5}"
    worksheet.add_chart(chart, chart_position)
    
    print(f"Added ARIMA forecast chart to {worksheet.title} at {chart_position}")
    return chart


def add_chart_summary_info(worksheet, data_end_row, sheet_type, total_lags, computed_lags):
    """
    Add a compact summary panel below the chart data.
    
    Args:
        worksheet: openpyxl worksheet object
        data_end_row: Last row of main data
        sheet_type: Type of sheet (daily, weekly, etc.)
        total_lags: Total number of requested lags
        computed_lags: Number of successfully computed lags
    """
    
    summary_start_row = data_end_row + 8  # Below chart data
    
    # Add summary headers with formatting
    summary_data = [
        ("[CHART] ACF/PACF Analysis Summary", ""),
        ("Lag Range Used:", f"{computed_lags} of {total_lags} requested"),
        ("Computed on Column:", "Total_Files"),
        ("Analysis Type:", f"{sheet_type.title()} Time Series"),
        ("Status:", "[OK] Complete" if computed_lags > 0 else "[EMOJI] Limited Data")
    ]
    
    for i, (label, value) in enumerate(summary_data):
        worksheet.cell(row=summary_start_row + i, column=1, value=label)
        worksheet.cell(row=summary_start_row + i, column=2, value=value)
        
        # Bold the first row (title)
        if i == 0:
            worksheet.cell(row=summary_start_row + i, column=1).font = openpyxl.styles.Font(bold=True)
    
    print(f"[SUCCESS] Added summary panel to {worksheet.title}")


def create_acf_pacf_dashboard_sheet(workbook, acf_pacf_sheets):
    """
    Create a dedicated dashboard sheet with all ACF/PACF charts for comparison.
    
    Args:
        workbook: openpyxl workbook object
        acf_pacf_sheets: List of sheet names containing ACF/PACF data
    
    Returns:
        Dashboard worksheet object
    """
    
    # Create new sheet
    dashboard = workbook.create_sheet(title="ACF_PACF_Dashboard")
    
    # Add title
    dashboard.cell(row=1, column=1, value="[TREND] ACF/PACF Analysis Dashboard")
    dashboard.cell(row=1, column=1).font = openpyxl.styles.Font(size=16, bold=True)
    
    dashboard.cell(row=2, column=1, value="Comparative view of autocorrelation patterns across all time scales")
    dashboard.cell(row=2, column=1).font = openpyxl.styles.Font(italic=True)
    
    # Add charts from each ACF/PACF sheet
    chart_row = 4
    for i, sheet_name in enumerate(acf_pacf_sheets):
        if sheet_name in workbook.sheetnames:
            source_sheet = workbook[sheet_name]
            
            # Extract chart data and create summary chart
            # This would be a simplified version focusing on key patterns
            
            dashboard.cell(row=chart_row, column=1, value=f"[CHART] {sheet_name}")
            dashboard.cell(row=chart_row, column=1).font = openpyxl.styles.Font(bold=True)
            
            chart_row += 15  # Space for each chart
    
    print(f"[OK] Created ACF/PACF Dashboard sheet")
    return dashboard


def enhance_acf_pacf_visualization(workbook):
    """
    Main function to enhance all ACF/PACF sheets with charts and summaries.
    
    Args:
        workbook: openpyxl workbook object
    
    Returns:
        List of sheets that were enhanced
    """
    
    print("enhance_acf_pacf_visualization called")
    print(f"Workbook type: {type(workbook)}")
    enhanced_sheets = []
    acf_pacf_sheet_names = []
    
    print(f"Scanning {len(workbook.sheetnames)} sheets for ACF/PACF sheets")
    for sheet_name in workbook.sheetnames:
        # Updated detection logic to match actual sheet names
        if any(pattern in sheet_name for pattern in ["ACF_PACF", "(ACF_PACF)"]):
            print(f"Found ACF/PACF sheet: {sheet_name}")
            acf_pacf_sheet_names.append(sheet_name)
            worksheet = workbook[sheet_name]
            
            # Determine sheet type
            if "Daily" in sheet_name:
                sheet_type = "daily"
            elif "Weekly" in sheet_name:
                sheet_type = "weekly"
            elif "Monthly" in sheet_name:
                sheet_type = "monthly"
            elif "Biweekly" in sheet_name:
                sheet_type = "biweekly"
            elif "Period" in sheet_name:
                sheet_type = "period"
            else:
                sheet_type = "unknown"
            
            # Find data range
            data_start_row = 2  # After headers
            data_end_row = worksheet.max_row
            
            # Count ACF/PACF columns (updated for Total_Files_ACF_Lag_X format)
            headers = [cell.value for cell in worksheet[3]]  # Headers are in row 3
            acf_cols = [h for h in headers if h and '_ACF_Lag_' in str(h)]
            pacf_cols = [h for h in headers if h and '_PACF_Lag_' in str(h)]
            total_lags = len(set([h.split('_Lag_')[1] for h in acf_cols + pacf_cols if '_Lag_' in str(h)]))
            
            print(f"Found {len(acf_cols)} ACF columns and {len(pacf_cols)} PACF columns")
            if acf_cols or pacf_cols:
                print(f"Sample ACF/PACF columns: {(acf_cols + pacf_cols)[:3]}")
            
            # Add ACF/PACF chart
            chart = add_acf_pacf_chart(worksheet, data_start_row, data_end_row, sheet_type)
            if chart:
                # Add ARIMA forecast chart
                arima_chart = add_arima_forecast_chart(worksheet, data_start_row, data_end_row, sheet_type)
                
                # Add summary info
                computed_lags = len(acf_cols)  # Approximate
                add_chart_summary_info(worksheet, data_end_row, sheet_type, total_lags, computed_lags)
                enhanced_sheets.append(sheet_name)
    
    # Create dashboard sheet
    if acf_pacf_sheet_names:
        create_acf_pacf_dashboard_sheet(workbook, acf_pacf_sheet_names)
    
    print(f"[OK] Enhanced {len(enhanced_sheets)} sheets with ACF/PACF visualizations")
    return enhanced_sheets


def add_forecast_summary_info(worksheet, data_end_row, sheet_type, forecast_quality=None, model_order=None):
    """
{{ ... }}
    
    Args:
        worksheet: openpyxl worksheet object
        data_end_row: Last row of main data
        sheet_type: Type of sheet (daily, weekly, etc.)
        forecast_quality: Quality assessment (Good/Warning/Error)
        model_order: ARIMA model order tuple (p,d,q)
    """
    
    # Position summary below chart area
    summary_start_row = data_end_row + 20  # Below chart
    
    # Summary data
    summary_data = [
        ("[TREND] ARIMA Forecast Summary", ""),
        ("Time Scale:", sheet_type.title()),
        ("Forecast Quality:", forecast_quality or "Unknown"),
        ("Model Order:", f"ARIMA{model_order}" if model_order else "Unknown"),
        ("Forecast Horizon:", f"{_get_forecast_horizon(sheet_type)} periods")
    ]
    
    # Add summary to worksheet
    for i, (label, value) in enumerate(summary_data):
        worksheet.cell(row=summary_start_row + i, column=1, value=label)
        worksheet.cell(row=summary_start_row + i, column=2, value=value)
        
        # Bold the first row (title)
        if i == 0:
            worksheet.cell(row=summary_start_row + i, column=1).font = Font(bold=True)
    
    print(f"[OK] Added ARIMA forecast summary to {worksheet.title}")


def _get_forecast_horizon(sheet_type):
    """
    Get the forecast horizon for a given sheet type.
    
    Args:
        sheet_type: Type of sheet (daily, weekly, etc.)
        
    Returns:
        int: Forecast horizon in periods
    """
    horizons = {
        'daily': 14,
        'weekly': 6,
        'biweekly': 4,
        'monthly': 3,
        'period': 2
    }
    return horizons.get(sheet_type, 14)


def enhance_arima_forecast_visualization(workbook):
    """
    Add ARIMA forecast charts to all sheets containing the necessary forecast data.
    This logic is based on the proven manual_arima_fix.py approach.
    
    Args:
        workbook: openpyxl workbook object
    
    Returns:
        List of sheets that were enhanced with forecast charts
    """
    
    print("[TREND] Enhancing report with ARIMA forecast charts...")
    charts_added = 0
    enhanced_sheets = []
    
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        
        if ws.max_row <= 1:
            continue
            
        headers = [str(cell.value) for cell in ws[1]]
        total_files_col, forecast_col = None, None
        
        # Find required columns
        for i, header in enumerate(headers, 1):
            if header == 'Total_Files':
                total_files_col = i
            elif header == 'Total_Files_Forecast':
                forecast_col = i
        
        # Proceed only if both columns are found
        if total_files_col and forecast_col:
            # Check if the forecast data is numeric
            first_forecast = ws.cell(row=2, column=forecast_col).value
            if isinstance(first_forecast, (int, float)):
                print(f"  -> Adding ARIMA chart to: {sheet_name}")
                
                # Create chart with proven logic
                chart = LineChart()
                chart.title = f"ARIMA Forecast vs. Actual - {sheet_name}"
                chart.style = 12
                chart.height = 10
                chart.width = 18
                chart.y_axis.title = "Total Files"
                chart.x_axis.title = "Time"
                
                # Define data series for historical and forecast values
                historical_data = Reference(ws, min_col=total_files_col, min_row=1, max_row=ws.max_row)
                forecast_data = Reference(ws, min_col=forecast_col, min_row=1, max_row=ws.max_row)
                
                chart.add_data(historical_data, titles_from_data=True)
                chart.add_data(forecast_data, titles_from_data=True)
                
                # Style the series for clarity
                if len(chart.series) > 0:  # Actual data
                    chart.series[0].graphicalProperties.line.solidFill = "4F81BD"  # Blue
                    chart.series[0].graphicalProperties.line.width = 25000
                if len(chart.series) > 1:  # Forecast data
                    chart.series[1].graphicalProperties.line.solidFill = "C0504D"  # Red
                    chart.series[1].graphicalProperties.line.dashStyle = "dash"
                    chart.series[1].graphicalProperties.line.width = 25000
                
                # Position chart below existing data
                chart_position = f'A{ws.max_row + 5}'
                ws.add_chart(chart, chart_position)
                charts_added += 1
                enhanced_sheets.append(sheet_name)
            else:
                print(f"  -> Skipping {sheet_name}: Forecast data is not numeric ('{first_forecast}')")
    
    if charts_added > 0:
        print(f"[OK] Successfully added {charts_added} ARIMA forecast charts.")
    else:
        print("[EMOJI] No new ARIMA charts were added. Check if forecast columns contain numeric data.")
    
    return enhanced_sheets


def _detect_sheet_type(sheet_name: str) -> str:
    """
    Detect the sheet type from the sheet name.
    
    Args:
        sheet_name: Name of the Excel sheet
        
    Returns:
        str: Sheet type (daily, weekly, monthly, biweekly, period, unknown)
    """
    lower = sheet_name.lower()
    if "daily" in lower:
        return "daily"
    if "weekly" in lower:
        return "weekly"
    if "monthly" in lower:
        return "monthly"
    if "biweekly" in lower:
        return "biweekly"
    if "period" in lower:
        return "period"
    return "unknown"
