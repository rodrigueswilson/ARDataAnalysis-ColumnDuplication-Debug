"""
Pipeline Utilities for AR Data Analysis

This module provides utility functions for pipeline definitions to ensure
consistent filtering and data processing across the system.
"""

class PipelineFilterUtils:
    """
    Utility class for pipeline filters to ensure consistency across the system.
    Provides static methods that return standard MongoDB match stages.
    """
    
    @staticmethod
    def get_base_filter():
        """
        Returns the base filter that should be applied to most pipelines.
        Filters out records with School_Year="N/A" and restricts to JPG/MP3 files.
        
        Returns:
            dict: MongoDB $match stage dictionary
        """
        return {
            "$match": {
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]}
            }
        }
    
    @staticmethod
    def get_collection_day_filter():
        """
        Returns a filter for collection days only.
        
        Returns:
            dict: MongoDB $match stage dictionary
        """
        return {
            "$match": {
                "is_collection_day": True
            }
        }
    
    @staticmethod
    def get_non_outlier_filter():
        """
        Returns a filter for non-outlier files only.
        
        Returns:
            dict: MongoDB $match stage dictionary
        """
        return {
            "$match": {
                "Outlier_Status": False
            }
        }
    
    @staticmethod
    def get_both_filters():
        """
        Returns a combined filter for both collection days and non-outlier files.
        This is the standard research dataset filter.
        
        Returns:
            dict: MongoDB $match stage dictionary
        """
        return {
            "$match": {
                "School_Year": {"$ne": "N/A"},
                "file_type": {"$in": ["JPG", "MP3"]},
                "is_collection_day": True,
                "Outlier_Status": False
            }
        }
    
    @staticmethod
    def get_mp3_filter():
        """
        Returns a filter for MP3 files only.
        
        Returns:
            dict: MongoDB $match stage dictionary
        """
        return {
            "$match": {
                "file_type": "MP3"
            }
        }
    
    @staticmethod
    def get_jpg_filter():
        """
        Returns a filter for JPG files only.
        
        Returns:
            dict: MongoDB $match stage dictionary
        """
        return {
            "$match": {
                "file_type": "JPG"
            }
        }


# Factory function to create DataCleaningUtils-compatible pipelines
def create_pipeline_with_filters(base_stages, filters=None):
    """
    Creates a pipeline by combining filter stages with base aggregation stages.
    
    Args:
        base_stages (list): The core aggregation stages (e.g., $group, $project)
        filters (list, optional): List of filter stages to apply before base_stages.
            If None, applies the default base filter.
    
    Returns:
        list: Complete MongoDB pipeline with filters applied
    """
    if filters is None:
        filters = [PipelineFilterUtils.get_base_filter()]
    
    return filters + base_stages
