"""
Data Enrichment Utilities for AR Data Analysis

This module provides utilities for enriching file data with contextual information
including calendar details, activity schedules, and metadata derivation.

Key Functions:
- Contextual information derivation
- Activity schedule mapping
- Metadata enrichment for files
- Collection day validation
"""

import datetime
from typing import Dict, Any, Optional
from .config import get_activity_schedule
from .calendar import is_collection_day, get_period_for_date, get_school_year_for_date


def get_contextual_info(
    dt_obj: datetime.datetime,
    calendar: dict,
    non_collection_days: dict,
    schedule: list,
    collection_day_map: dict,
    audio_props: Optional[Dict[str, Any]] = None,
    is_outlier: bool = False
) -> Dict[str, Any]:
    """
    Derives all contextual information for a given file based on its timestamp.

    This function acts as a central hub for data enrichment, pulling together
    calendar details, holiday information, and activity schedules to create a
    comprehensive metadata profile for each file.

    Args:
        dt_obj: The datetime object representing when the file was created.
        calendar: The school calendar configuration dictionary.
        non_collection_days: Dictionary of non-collection days to check against.
        schedule: List containing the daily activity schedule.
        collection_day_map: Precomputed map of valid collection days.
        audio_props: Optional dictionary with audio-specific properties.
        is_outlier: Boolean indicating if this file is marked as an outlier.

    Returns:
        A dictionary containing all derived contextual information.
    """
    date_obj = dt_obj.date()
    
    # Basic temporal information
    context = {
        'ISO_Date': date_obj.isoformat(),
        'ISO_Year': dt_obj.year,
        'ISO_Week': dt_obj.isocalendar()[1],
        'ISO_YearWeek': f"{dt_obj.year}-W{dt_obj.isocalendar()[1]:02d}",
        'Day_of_Week': dt_obj.strftime('%A'),
        'Month': dt_obj.month,
        'Year': dt_obj.year,
        'Hour': dt_obj.hour,
        'Minute': dt_obj.minute,
        'Time_of_Day': dt_obj.strftime('%H:%M:%S'),
        'Outlier_Status': is_outlier
    }

    # Collection day validation
    context['is_collection_day'] = is_collection_day(date_obj, collection_day_map)
    
    # School year and period information
    if context['is_collection_day']:
        context['School_Year'] = get_school_year_for_date(date_obj, collection_day_map) or 'N/A'
        context['Collection_Period'] = get_period_for_date(date_obj, collection_day_map) or 'N/A'
    else:
        context['School_Year'] = 'N/A'
        context['Collection_Period'] = 'N/A'

    # Non-collection day information
    if date_obj in non_collection_days:
        non_collection_info = non_collection_days[date_obj]
        context['Non_Collection_Reason'] = non_collection_info.get('reason', 'Unknown')
        context['Non_Collection_Type'] = non_collection_info.get('type', 'Unknown')
    else:
        context['Non_Collection_Reason'] = None
        context['Non_Collection_Type'] = None

    # Activity schedule mapping
    context['Scheduled_Activity'] = _get_scheduled_activity(dt_obj, schedule)

    # Audio-specific properties
    if audio_props:
        context.update(audio_props)

    return context


def _get_scheduled_activity(dt_obj: datetime.datetime, schedule: list) -> str:
    """
    Determines the scheduled activity for a given datetime based on the activity schedule.

    Args:
        dt_obj: The datetime object to check.
        schedule: List containing the daily activity schedule.

    Returns:
        The name of the scheduled activity, or 'Unscheduled' if none found.
    """
    time_str = dt_obj.strftime('%H:%M')
    day_of_week = dt_obj.strftime('%A')

    for activity in schedule:
        # Check if this activity applies to the current day
        if day_of_week in activity.get('days', []):
            start_time = activity.get('start_time', '')
            end_time = activity.get('end_time', '')
            
            if start_time <= time_str <= end_time:
                return activity.get('name', 'Unknown Activity')

    return 'Unscheduled'


def enrich_file_metadata(
    file_path: str,
    timestamp: datetime.datetime,
    file_type: str,
    file_size_mb: float,
    collection_day_map: dict,
    audio_props: Optional[Dict[str, Any]] = None,
    is_outlier: bool = False
) -> Dict[str, Any]:
    """
    Creates enriched metadata for a file with all contextual information.

    Args:
        file_path: Path to the file
        timestamp: File creation timestamp
        file_type: Type of file (MP3, JPG, etc.)
        file_size_mb: File size in megabytes
        collection_day_map: Precomputed collection day map
        audio_props: Optional audio-specific properties
        is_outlier: Whether the file is marked as an outlier

    Returns:
        Dictionary containing complete file metadata
    """
    from .config import get_school_calendar, get_non_collection_days, get_activity_schedule
    
    # Get configuration data
    calendar = get_school_calendar()
    non_collection_days = get_non_collection_days()
    schedule = get_activity_schedule()

    # Get contextual information
    context = get_contextual_info(
        timestamp, calendar, non_collection_days, schedule,
        collection_day_map, audio_props, is_outlier
    )

    # Combine with basic file information
    metadata = {
        'File_Path': file_path,
        'File_Name': file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1],
        'File_Type': file_type,
        'File_Size_MB': file_size_mb,
        'Creation_DateTime': timestamp.isoformat(),
        **context
    }

    return metadata


def validate_enriched_data(metadata: Dict[str, Any]) -> Dict[str, Any]:
    """
    Validates and cleans enriched metadata.

    Args:
        metadata: Dictionary containing file metadata

    Returns:
        Validated and cleaned metadata dictionary
    """
    # Ensure required fields are present
    required_fields = [
        'File_Path', 'File_Type', 'File_Size_MB', 'ISO_Date',
        'School_Year', 'Collection_Period', 'is_collection_day'
    ]

    for field in required_fields:
        if field not in metadata:
            metadata[field] = 'N/A' if field in ['School_Year', 'Collection_Period'] else None

    # Validate data types
    if 'File_Size_MB' in metadata:
        try:
            metadata['File_Size_MB'] = float(metadata['File_Size_MB'])
        except (ValueError, TypeError):
            metadata['File_Size_MB'] = 0.0

    if 'Outlier_Status' in metadata:
        metadata['Outlier_Status'] = bool(metadata['Outlier_Status'])

    if 'is_collection_day' in metadata:
        metadata['is_collection_day'] = bool(metadata['is_collection_day'])

    return metadata


def get_file_category(file_type: str, file_size_mb: float, duration_seconds: Optional[float] = None) -> str:
    """
    Categorizes a file based on its properties.

    Args:
        file_type: Type of file (MP3, JPG, etc.)
        file_size_mb: File size in megabytes
        duration_seconds: Duration in seconds (for audio files)

    Returns:
        File category string
    """
    if file_type.upper() == 'MP3':
        if duration_seconds is not None:
            if duration_seconds < 30:
                return 'Short Audio'
            elif duration_seconds < 300:  # 5 minutes
                return 'Medium Audio'
            else:
                return 'Long Audio'
        else:
            return 'Audio'
    
    elif file_type.upper() == 'JPG':
        if file_size_mb < 1.0:
            return 'Small Image'
        elif file_size_mb < 5.0:
            return 'Medium Image'
        else:
            return 'Large Image'
    
    else:
        return 'Other'


def summarize_enrichment_stats(metadata_list: list) -> Dict[str, Any]:
    """
    Generates summary statistics for a list of enriched metadata.

    Args:
        metadata_list: List of metadata dictionaries

    Returns:
        Dictionary containing summary statistics
    """
    if not metadata_list:
        return {}

    total_files = len(metadata_list)
    
    # Count by file type
    file_types = {}
    collection_days = 0
    outliers = 0
    school_years = set()
    periods = set()

    for metadata in metadata_list:
        # File type counts
        file_type = metadata.get('File_Type', 'Unknown')
        file_types[file_type] = file_types.get(file_type, 0) + 1

        # Collection day counts
        if metadata.get('is_collection_day', False):
            collection_days += 1

        # Outlier counts
        if metadata.get('Outlier_Status', False):
            outliers += 1

        # School years and periods
        school_year = metadata.get('School_Year')
        if school_year and school_year != 'N/A':
            school_years.add(school_year)

        period = metadata.get('Collection_Period')
        if period and period != 'N/A':
            periods.add(period)

    return {
        'total_files': total_files,
        'file_types': file_types,
        'collection_day_files': collection_days,
        'non_collection_day_files': total_files - collection_days,
        'outlier_files': outliers,
        'school_years': sorted(list(school_years)),
        'periods': sorted(list(periods)),
        'collection_day_percentage': (collection_days / total_files * 100) if total_files > 0 else 0,
        'outlier_percentage': (outliers / total_files * 100) if total_files > 0 else 0
    }
