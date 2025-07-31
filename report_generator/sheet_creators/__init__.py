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
