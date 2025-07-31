"""
Calendar and Date Utilities for AR Data Analysis

This module handles calendar-related operations including collection day calculations,
period management, and date validation for the AR Data Analysis system.

Key Functions:
- Collection day computation and validation
- Period date range calculations
- Calendar-based data enrichment
- Holiday and non-collection day handling
"""

import datetime
from typing import Dict, Any, Optional
from .config import get_school_calendar, get_non_collection_days


def calculate_collection_days_for_period(period_name: str) -> int:
    """
    Calculate the total number of valid collection days for a given period.
    
    This function counts all weekdays within the period's date range,
    excluding holidays, breaks, and other non-collection days as defined
    in the configuration.
    
    Args:
        period_name: The period identifier (e.g., 'SY 21-22 P1')
        
    Returns:
        int: Total number of valid collection days in the period
    """
    try:
        school_calendar = get_school_calendar()
        non_collection_days = get_non_collection_days()
        
        # Find the period in the school calendar
        period_dates = None
        for year_data in school_calendar.values():
            if period_name in year_data.get('periods', {}):
                period_dates = year_data['periods'][period_name]
                break
        
        if not period_dates or len(period_dates) != 2:
            print(f"[WARNING] Period '{period_name}' not found or invalid in school calendar")
            return 0
        
        start_date, end_date = period_dates
        
        # Count valid collection days
        collection_days = 0
        current_date = start_date
        
        while current_date <= end_date:
            # Check if it's a weekday (Monday=0, Sunday=6)
            if current_date.weekday() < 5:  # Monday through Friday
                # Check if it's not a non-collection day
                if current_date not in non_collection_days:
                    collection_days += 1
            
            current_date += datetime.timedelta(days=1)
        
        return collection_days
        
    except Exception as e:
        print(f"[ERROR] Error calculating collection days for period '{period_name}': {e}")
        return 0


def precompute_collection_days(school_calendar: dict, non_collection_days: dict) -> Dict[datetime.date, Dict[str, Any]]:
    """
    Precomputes a map of all valid data collection days for efficient lookups.

    This function iterates through the date ranges defined in the school calendar
    and builds a dictionary where keys are `datetime.date` objects and values
    are dictionaries containing metadata about each collection day.

    Args:
        school_calendar: The school calendar configuration dictionary.
        non_collection_days: Dictionary of non-collection days to exclude.

    Returns:
        A dictionary mapping each valid collection `datetime.date` to a dictionary
        of its properties, such as `School_Year` and `Collection_Period`.
    """
    collection_day_map = {}

    for school_year, year_data in school_calendar.items():
        start_date = year_data['start_date']
        end_date = year_data['end_date']
        periods = year_data.get('periods', {})

        # Iterate through each day in the school year
        current_date = start_date
        while current_date <= end_date:
            # Check if it's a weekday (Monday=0, Sunday=6)
            if current_date.weekday() < 5:  # Monday through Friday
                # Check if it's not a non-collection day
                if current_date not in non_collection_days:
                    # Determine which period this date belongs to
                    collection_period = None
                    for period_name, period_dates in periods.items():
                        if len(period_dates) == 2:
                            period_start, period_end = period_dates
                            if period_start <= current_date <= period_end:
                                collection_period = period_name
                                break

                    # Add to the collection day map
                    collection_day_map[current_date] = {
                        'School_Year': school_year,
                        'Collection_Period': collection_period or 'N/A'
                    }

            current_date += datetime.timedelta(days=1)

    return collection_day_map


def is_collection_day(date_obj: datetime.date, collection_day_map: dict) -> bool:
    """
    Determines if a given date is a collection day based on the precomputed map.
    
    Args:
        date_obj: The date to check
        collection_day_map: Precomputed map of collection days
        
    Returns:
        True if the date is a valid collection day, False otherwise
    """
    return date_obj in collection_day_map


def get_period_for_date(date_obj: datetime.date, collection_day_map: dict) -> Optional[str]:
    """
    Gets the collection period for a given date.
    
    Args:
        date_obj: The date to look up
        collection_day_map: Precomputed map of collection days
        
    Returns:
        Collection period name or None if not a collection day
    """
    day_info = collection_day_map.get(date_obj)
    return day_info.get('Collection_Period') if day_info else None


def get_school_year_for_date(date_obj: datetime.date, collection_day_map: dict) -> Optional[str]:
    """
    Gets the school year for a given date.
    
    Args:
        date_obj: The date to look up
        collection_day_map: Precomputed map of collection days
        
    Returns:
        School year name or None if not a collection day
    """
    day_info = collection_day_map.get(date_obj)
    return day_info.get('School_Year') if day_info else None


def get_date_range_for_period(period_name: str) -> Optional[tuple]:
    """
    Gets the start and end dates for a given period.
    
    Args:
        period_name: The period identifier (e.g., 'SY 21-22 P1')
        
    Returns:
        Tuple of (start_date, end_date) or None if period not found
    """
    school_calendar = get_school_calendar()
    
    for year_data in school_calendar.values():
        if period_name in year_data.get('periods', {}):
            period_dates = year_data['periods'][period_name]
            if len(period_dates) == 2:
                return tuple(period_dates)
    
    return None


def get_all_periods() -> list:
    """
    Gets a list of all available periods across all school years.
    
    Returns:
        List of period names
    """
    school_calendar = get_school_calendar()
    periods = []
    
    for year_data in school_calendar.values():
        periods.extend(year_data.get('periods', {}).keys())
    
    return periods


def get_periods_for_school_year(school_year: str) -> list:
    """
    Gets all periods for a specific school year.
    
    Args:
        school_year: The school year identifier
        
    Returns:
        List of period names for the specified school year
    """
    school_calendar = get_school_calendar()
    year_data = school_calendar.get(school_year, {})
    return list(year_data.get('periods', {}).keys())
