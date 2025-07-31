"""
Configuration Management Utilities for AR Data Analysis

This module handles loading, parsing, and accessing configuration data from config.yaml.
It provides centralized configuration management with proper error handling and
data type conversions.

Key Functions:
- Configuration loading and parsing
- Date string to datetime.date conversion
- Configuration data accessors
- Error handling for missing or invalid configurations
"""

import datetime
import yaml
from pathlib import Path
from typing import Dict, Any


def _load_config() -> Dict[str, Any]:
    """
    Loads and parses the config.yaml file, converting date strings to objects.
    
    Returns:
        Dictionary containing the parsed configuration data
        
    Raises:
        FileNotFoundError: If config.yaml is not found
        yaml.YAMLError: If config.yaml contains invalid YAML
        TypeError: If date conversion fails
    """
    config_path = Path(__file__).parent.parent / "config.yaml"
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file not found at: {config_path}")

    with open(config_path, 'r') as f:
        config = yaml.safe_load(f)

    # --- Post-process configuration data ---
    # Convert date strings in school_calendar to datetime.date objects
    for year, details in config.get('school_calendar', {}).items():
        if 'start_date' in details and isinstance(details['start_date'], str):
            details['start_date'] = datetime.date.fromisoformat(details['start_date'])
        if 'end_date' in details and isinstance(details['end_date'], str):
            details['end_date'] = datetime.date.fromisoformat(details['end_date'])
        for period, dates in details.get('periods', {}).items():
            if isinstance(dates, list) and all(isinstance(d, str) for d in dates):
                details['periods'][period] = [datetime.date.fromisoformat(d) for d in dates]

    # Convert date string keys in non_collection_days to datetime.date objects
    if 'non_collection_days' in config:
        config['non_collection_days'] = {
            datetime.date.fromisoformat(date_str): info
            for date_str, info in config['non_collection_days'].items()
        }

    return config


# Load configuration once when the module is imported
CONFIG = {}
try:
    CONFIG = _load_config()
except (FileNotFoundError, yaml.YAMLError, TypeError) as e:
    print(f"[ERROR] Critical Error: Could not load or parse config.yaml. Details: {e}")
    # Allow module to load but dependent functions will fail gracefully.


def get_school_calendar() -> Dict[str, Any]:
    """
    Returns the school calendar configuration from the loaded CONFIG.
    
    Returns:
        Dictionary containing school calendar data with converted date objects
    """
    return CONFIG.get('school_calendar', {})


def get_non_collection_days() -> Dict[datetime.date, Dict[str, Any]]:
    """
    Returns a dictionary of non-collection days from the loaded CONFIG.
    
    Returns:
        Dictionary mapping datetime.date objects to non-collection day info
    """
    return CONFIG.get('non_collection_days', {})


def get_activity_schedule() -> list:
    """
    Returns the daily activity schedule from the loaded CONFIG.
    
    Returns:
        List containing the daily activity schedule configuration
    """
    return CONFIG.get('daily_activity_schedule', [])


def get_config() -> Dict[str, Any]:
    """
    Returns the complete loaded configuration.
    
    Returns:
        Complete configuration dictionary
    """
    return CONFIG


def reload_config() -> bool:
    """
    Reloads the configuration from disk.
    
    Returns:
        True if reload was successful, False otherwise
    """
    global CONFIG
    try:
        CONFIG = _load_config()
        return True
    except (FileNotFoundError, yaml.YAMLError, TypeError) as e:
        print(f"[ERROR] Failed to reload configuration: {e}")
        return False


def is_config_loaded() -> bool:
    """
    Checks if configuration has been successfully loaded.
    
    Returns:
        True if configuration is loaded and not empty, False otherwise
    """
    return bool(CONFIG)
