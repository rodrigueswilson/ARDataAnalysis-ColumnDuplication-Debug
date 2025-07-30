"""
Excel Formatting and Styling Utilities
======================================

This module contains all Excel formatting functions, styling utilities,
and worksheet appearance management for the report generator.
"""

from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

class ExcelFormatter:
    """
    Handles all Excel formatting operations including colors, fonts, borders,
    and cell styling for consistent report appearance.
    """
    
    def __init__(self):
        """Initialize formatter with default color scheme."""
        self.color_scheme = {
            'header_bg': '1F497D', 
            'header_font': 'FFFFFF',
            'alt_row_bg': 'DDEBF7', 
            'total_row_bg': '4F81BD',
            'total_row_font': 'FFFFFF', 
            'chart_series1': '4F81BD',
            'chart_series2': 'C0504D'
        }
    
    def format_sheet(self, ws):
        """
        Applies consistent formatting to a worksheet with alternating row colors 
        and proper column widths.
        
        Args:
            ws: openpyxl worksheet object
        """
        if ws.max_row <= 1:
            return  # Skip empty sheets
        
        # Header formatting
        header_font = Font(bold=True, color=self.color_scheme['header_font'])
        header_fill = PatternFill(start_color=self.color_scheme['header_bg'], 
                                 end_color=self.color_scheme['header_bg'], 
                                 fill_type='solid')
        header_alignment = Alignment(horizontal='center', vertical='center')
        
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
        
        # Alternating row colors (skip header)
        alt_fill = PatternFill(start_color=self.color_scheme['alt_row_bg'], 
                              end_color=self.color_scheme['alt_row_bg'], 
                              fill_type='solid')
        
        for row_num in range(2, ws.max_row + 1):
            if row_num % 2 == 0:  # Even rows get alternate color
                for cell in ws[row_num]:
                    if cell.fill.start_color.index == '00000000':  # Only if not already filled
                        cell.fill = alt_fill
        
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            ws.column_dimensions[column_letter].width = adjusted_width
    
    def add_total_row(self, ws, df):
        """
        Adds a formatted total row to the worksheet.
        
        Args:
            ws: openpyxl worksheet object
            df: pandas DataFrame with the data
        """
        if df.empty:
            return
        
        total_row_num = ws.max_row + 1
        
        # Total row formatting
        total_font = Font(bold=True, color=self.color_scheme['total_row_font'])
        total_fill = PatternFill(start_color=self.color_scheme['total_row_bg'], 
                                end_color=self.color_scheme['total_row_bg'], 
                                fill_type='solid')
        
        # Add "TOTAL" label in first column
        ws.cell(row=total_row_num, column=1, value="TOTAL")
        ws.cell(row=total_row_num, column=1).font = total_font
        ws.cell(row=total_row_num, column=1).fill = total_fill
        
        # Calculate totals for numeric columns
        for col_idx, column_name in enumerate(df.columns, start=2):
            cell = ws.cell(row=total_row_num, column=col_idx)
            cell.font = total_font
            cell.fill = total_fill
            
            try:
                # Check if column is numeric and can be summed
                if pd.api.types.is_numeric_dtype(df[column_name]):
                    total_value = df[column_name].sum()
                    if pd.isna(total_value):
                        cell.value = ""
                    else:
                        cell.value = total_value
                else:
                    cell.value = ""
            except Exception as e:
                # Fallback for any unexpected data type issues
                cell.value = ""
    
    def add_audio_characteristics_total_row(self, ws, df):
        """
        Adds a custom total row for Audio Note Characteristics that only sums 'Per Year' rows.
        
        Args:
            ws: openpyxl worksheet object  
            df: pandas DataFrame with the data
        """
        if df.empty:
            return
        
        # Filter for 'Per Year' rows only
        per_year_df = df[df.iloc[:, 0].str.contains('Per Year', na=False)]
        
        if per_year_df.empty:
            return
        
        total_row_num = ws.max_row + 1
        
        # Total row formatting
        total_font = Font(bold=True, color=self.color_scheme['total_row_font'])
        total_fill = PatternFill(start_color=self.color_scheme['total_row_bg'], 
                                end_color=self.color_scheme['total_row_bg'], 
                                fill_type='solid')
        
        # Add "TOTAL (Per Year Only)" label
        ws.cell(row=total_row_num, column=1, value="TOTAL (Per Year Only)")
        ws.cell(row=total_row_num, column=1).font = total_font
        ws.cell(row=total_row_num, column=1).fill = total_fill
        
        # Calculate totals for numeric columns from per_year rows only
        for col_idx, column_name in enumerate(df.columns, start=2):
            cell = ws.cell(row=total_row_num, column=col_idx)
            cell.font = total_font
            cell.fill = total_fill
            
            if per_year_df[column_name].dtype in ['int64', 'float64']:
                try:
                    total_value = per_year_df[column_name].sum()
                    cell.value = total_value
                except:
                    cell.value = ""
            else:
                cell.value = ""
    
    def add_top_days_table(self, ws, df, start_row):
        """
        Adds a table of top efficiency days to the worksheet at the specified starting row.
        
        Args:
            ws: The worksheet to add the table to
            df: DataFrame containing the top efficiency data
            start_row: The row number to start adding the table at
        """
        if df.empty:
            return
        
        # Add table header
        header_font = Font(bold=True, color=self.color_scheme['header_font'])
        header_fill = PatternFill(start_color=self.color_scheme['header_bg'], 
                                 end_color=self.color_scheme['header_bg'], 
                                 fill_type='solid')
        
        ws.cell(row=start_row, column=1, value="Top 10 Days by Efficiency")
        ws.cell(row=start_row, column=1).font = header_font
        ws.cell(row=start_row, column=1).fill = header_fill
        
        # Add column headers
        headers = ["Date", "Total Files", "Audio Minutes", "Files per Audio Minute"]
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=start_row + 1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
        
        # Add data rows
        for row_idx, (_, row) in enumerate(df.head(10).iterrows(), start=start_row + 2):
            ws.cell(row=row_idx, column=1, value=row['_id'])
            ws.cell(row=row_idx, column=2, value=row['total_files_day'])
            ws.cell(row=row_idx, column=3, value=round(row['total_duration_minutes'], 1))
            ws.cell(row=row_idx, column=4, value=round(row['efficiency_files_per_audio_minute'], 2))
            
            # Apply alternating row colors
            if row_idx % 2 == 0:
                alt_fill = PatternFill(start_color=self.color_scheme['alt_row_bg'], 
                                     end_color=self.color_scheme['alt_row_bg'], 
                                     fill_type='solid')
                for col in range(1, 5):
                    ws.cell(row=row_idx, column=col).fill = alt_fill
