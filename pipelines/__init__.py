"""
Pipeline Registry for AR Data Analysis

This module serves as the central registry for all MongoDB aggregation pipelines.
It imports pipelines from various themed modules and consolidates them into a
single PIPELINES dictionary for easy access throughout the application.

Module Structure:
- daily_counts.py: Daily aggregation pipelines
- weekly_counts.py: Weekly aggregation pipelines  
- activity_analysis.py: Activity and pattern analysis pipelines
- camera_usage.py: Camera usage analysis pipelines
- file_analysis.py: File size and distribution analysis
- dashboard_data.py: Dashboard-specific data pipelines
- mp3_analysis.py: MP3 duration analysis pipelines
- time_series.py: Time series analysis pipelines (existing)
"""

# Import all pipeline modules
from .daily_counts import DAILY_PIPELINES
from .weekly_counts import WEEKLY_PIPELINES
from .biweekly_counts import BIWEEKLY_PIPELINES
from .activity_analysis import ACTIVITY_PIPELINES
from .camera_usage import CAMERA_PIPELINES
from .file_analysis import FILE_PIPELINES
from .dashboard_data import DASHBOARD_PIPELINES
from .mp3_analysis import MP3_PIPELINES
from .time_series import TIME_SERIES_PIPELINES

# Consolidate all pipeline dictionaries into a single registry
PIPELINES = {}

# Add all themed pipeline modules
PIPELINES.update(DAILY_PIPELINES)
PIPELINES.update(WEEKLY_PIPELINES)
PIPELINES.update(BIWEEKLY_PIPELINES)
PIPELINES.update(ACTIVITY_PIPELINES)
PIPELINES.update(CAMERA_PIPELINES)
PIPELINES.update(FILE_PIPELINES)
PIPELINES.update(DASHBOARD_PIPELINES)
PIPELINES.update(MP3_PIPELINES)
PIPELINES.update(TIME_SERIES_PIPELINES)

# Export the main registry
__all__ = ['PIPELINES']
