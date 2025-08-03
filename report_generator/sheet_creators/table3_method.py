"""
Temporary file containing the _create_filter_impact_summary method to be added to __init__.py
"""

def _create_filter_impact_summary(self, ws, start_row, totals):
    """
    Create Table 3: Filter Impact Summary
    
    Args:
        ws: Worksheet to add the table to
        start_row: Row to start the table at
        totals: Dictionary containing the totals for each category
        
    Returns:
        int: The next row after the table
    """
    try:
        print(f"[INFO] Creating Table 3: Filter Impact Summary at row {start_row}")
        
        # Table 3 header
        table3_start_row = start_row
        ws.cell(row=table3_start_row, column=1, value="Table 3: Filter Impact Summary")
        self.formatter.apply_section_header_style(ws, f'A{table3_start_row}')
        
        # Headers for summary table (human-readable)
        summary_headers = ['Filter Category', 'Number of Files', '% of All Files', 'TOTAL']
        summary_header_row = table3_start_row + 2
        
        for col, header in enumerate(summary_headers, 1):
            ws.cell(row=summary_header_row, column=col, value=header)
        self.formatter.apply_header_style(ws, f'A{summary_header_row}:D{summary_header_row}')
        
        # Calculate summary data (clear and descriptive labels)
        total_files = totals['total_files']
        summary_data = [
            ('School Outliers (unusual files from school days)', totals['school_outliers'], 
             (totals['school_outliers'] / total_files * 100) if total_files > 0 else 0),
            ('Non-School Normal (regular files from non-school days)', totals['non_school_normal'], 
             (totals['non_school_normal'] / total_files * 100) if total_files > 0 else 0),
            ('School Normal (final research dataset)', totals['school_normal'], 
             (totals['school_normal'] / total_files * 100) if total_files > 0 else 0),
            ('Non-School Outliers (unusual files from non-school days)', totals['non_school_outliers'], 
             (totals['non_school_outliers'] / total_files * 100) if total_files > 0 else 0)
        ]
        
        # Fill summary data
        summary_row = summary_header_row + 1
        
        for criterion, count, percentage in summary_data:
            ws.cell(row=summary_row, column=1, value=criterion)
            ws.cell(row=summary_row, column=2, value=count)
            
            # Format percentage
            pct_cell = ws.cell(row=summary_row, column=3, value=percentage / 100)
            pct_cell.number_format = '0.0%'
            
            # Add TOTAL column (same as Number of Files for each row)
            ws.cell(row=summary_row, column=4, value=count)
            
            summary_row += 1
        
        # Add formatted total row at the specific position
        table3_total_values = {
            1: 'TOTAL',
            2: totals['total_files'],
            3: 1.0,  # 100%
            4: totals['total_files']  # Total column should match the total files
        }
        
        # Add total row at the correct position
        self.formatter.add_total_row_at_position(ws, summary_row, table3_total_values)
        
        # Format total percentage (should be 100%)
        total_pct_cell = ws.cell(row=summary_row, column=3)
        total_pct_cell.number_format = '0.0%'
        
        summary_row += 1
        
        # Apply formatting to Table 3 (Filter Impact Summary)
        self.formatter.apply_data_style(ws, f'A{summary_header_row + 1}:D{summary_row - 1}')
        self.formatter.apply_alternating_row_colors(ws, summary_header_row + 1, summary_row - 1, 1, 4)
        
        print(f"[SUCCESS] Created Table 3: Filter Impact Summary")
        return summary_row
        
    except Exception as e:
        print(f"[ERROR] Failed to create Table 3: {e}")
        import traceback
        traceback.print_exc()
        return start_row  # Return the original start row if there's an error
