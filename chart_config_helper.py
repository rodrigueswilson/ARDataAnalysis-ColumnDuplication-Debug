"""
Chart Configuration Helper

This module provides configuration-driven chart placement logic to replace
string-based sheet name filtering. It reads chart configuration from
report_config.json and provides helper functions for determining which
charts should be added to which sheets.

This is part of the chart placement cleanup initiative to improve
maintainability and eliminate fragile string matching logic.
"""

import json
import os
from typing import Dict, List, Optional, Set


class ChartConfigHelper:
    """Helper class for configuration-driven chart placement logic."""
    
    def __init__(self, config_path: str = None):
        """
        Initialize the chart configuration helper.
        
        Args:
            config_path: Path to report_config.json. If None, uses default location.
        """
        if config_path is None:
            config_path = os.path.join(os.path.dirname(__file__), 'report_config.json')
        
        self.config_path = config_path
        self.config = self._load_config()
        self.sheet_configs = self._build_sheet_lookup()
    
    def _load_config(self) -> Dict:
        """Load and parse the report configuration file."""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            raise RuntimeError(f"Failed to load chart configuration from {self.config_path}: {e}")
    
    def _build_sheet_lookup(self) -> Dict[str, Dict]:
        """Build a lookup dictionary mapping sheet names to their configurations."""
        lookup = {}
        
        # Process all sheets from the configuration
        for sheet in self.config.get('sheets', []):
            sheet_name = sheet.get('sheet_name')
            if sheet_name:
                lookup[sheet_name] = sheet
        
        return lookup
    
    def should_add_chart(self, sheet_name: str, chart_type: str) -> bool:
        """
        Determine if a specific chart type should be added to a sheet.
        
        Args:
            sheet_name: Name of the Excel sheet
            chart_type: Type of chart ('acf_pacf', 'arima_forecast', 'dashboard')
        
        Returns:
            True if the chart should be added, False otherwise
        """
        # Get sheet configuration
        sheet_config = self.sheet_configs.get(sheet_name)
        if not sheet_config:
            # Fallback to legacy string matching for sheets not in config
            return self._legacy_string_matching(sheet_name, chart_type)
        
        # Check if sheet is enabled
        if not sheet_config.get('enabled', True):
            return False
        
        # Get chart configuration
        charts_config = sheet_config.get('charts', {})
        return charts_config.get(chart_type, False)
    
    def should_add_acf_pacf_columns(self, sheet_name: str) -> bool:
        """
        Determine if ACF/PACF columns should be added to a sheet.
        
        Args:
            sheet_name: Name of the Excel sheet
        
        Returns:
            True if ACF/PACF columns should be added, False otherwise
        """
        # Get sheet configuration
        sheet_config = self.sheet_configs.get(sheet_name)
        if not sheet_config:
            # Fallback to legacy logic for sheets not in config
            return self._legacy_column_matching(sheet_name, 'acf_pacf')
        
        # Check if sheet is enabled
        if not sheet_config.get('enabled', True):
            return False
        
        # Get columns configuration (fallback to charts config for backward compatibility)
        columns_config = sheet_config.get('columns', {})
        if not columns_config:
            # If no columns config, fall back to charts config for ACF/PACF
            charts_config = sheet_config.get('charts', {})
            return charts_config.get('acf_pacf', False)
        
        return columns_config.get('acf_pacf', False)
    
    def should_add_arima_columns(self, sheet_name: str) -> bool:
        """
        Determine if ARIMA forecast columns should be added to a sheet.
        
        Args:
            sheet_name: Name of the Excel sheet
        
        Returns:
            True if ARIMA columns should be added, False otherwise
        """
        # Get sheet configuration
        sheet_config = self.sheet_configs.get(sheet_name)
        if not sheet_config:
            # Fallback to legacy logic for sheets not in config
            return self._legacy_column_matching(sheet_name, 'arima')
        
        # Check if sheet is enabled
        if not sheet_config.get('enabled', True):
            return False
        
        # Get columns configuration (fallback to charts config for backward compatibility)
        columns_config = sheet_config.get('columns', {})
        if not columns_config:
            # If no columns config, fall back to charts config for ARIMA
            charts_config = sheet_config.get('charts', {})
            return charts_config.get('arima_forecast', False)
        
        return columns_config.get('arima', False)
    
    def _legacy_string_matching(self, sheet_name: str, chart_type: str) -> bool:
        """
        Fallback to legacy string matching for sheets not in configuration.
        This maintains backward compatibility during the transition period.
        """
        if chart_type == 'acf_pacf':
            return "(ACF_PACF)" in sheet_name or "ACF_PACF" in sheet_name
        elif chart_type == 'arima_forecast':
            return "(ACF_PACF)" in sheet_name
        elif chart_type == 'dashboard':
            return "Dashboard" in sheet_name
        
        return False
    
    def _legacy_column_matching(self, sheet_name: str, column_type: str) -> bool:
        """
        Fallback to legacy logic for column placement in sheets not in configuration.
        This maintains backward compatibility during the transition period.
        
        Args:
            sheet_name: Name of the Excel sheet
            column_type: Type of columns ('acf_pacf', 'arima')
        
        Returns:
            True if columns should be added based on legacy logic, False otherwise
        """
        if column_type == 'acf_pacf':
            # ACF/PACF columns should only be in ACF_PACF sheets
            return "(ACF_PACF)" in sheet_name or "ACF_PACF" in sheet_name
        elif column_type == 'arima':
            # ARIMA columns should only be in ACF_PACF sheets with enabled time scales
            if "(ACF_PACF)" in sheet_name or "ACF_PACF" in sheet_name:
                # Check if this is a daily or weekly sheet (enabled time scales)
                time_scale = self._detect_time_scale_legacy(sheet_name)
                enabled_scales = self.get_enabled_time_scales()
                return time_scale in enabled_scales
            return False
        
        return False
    
    def get_sheet_time_scale(self, sheet_name: str) -> Optional[str]:
        """
        Get the time scale for a sheet from configuration.
        
        Args:
            sheet_name: Name of the Excel sheet
        
        Returns:
            Time scale string ('daily', 'weekly', etc.) or None if not configured
        """
        sheet_config = self.sheet_configs.get(sheet_name)
        if sheet_config:
            return sheet_config.get('time_scale')
        
        # Fallback to legacy detection
        return self._detect_time_scale_legacy(sheet_name)
    
    def _detect_time_scale_legacy(self, sheet_name: str) -> Optional[str]:
        """Legacy time scale detection based on sheet name parsing."""
        lower = sheet_name.lower()
        if "daily" in lower:
            return "daily"
        elif "weekly" in lower:
            return "weekly"
        elif "biweekly" in lower:
            return "biweekly"
        elif "monthly" in lower:
            return "monthly"
        elif "period" in lower:
            return "period"
        
        return None
    
    def get_enabled_time_scales(self) -> Set[str]:
        """Get the set of enabled time scales from forecast configuration."""
        forecast_config = self.config.get('forecast_options', {})
        return set(forecast_config.get('time_scales', ['daily', 'weekly']))
    
    def get_sheets_with_chart_type(self, chart_type: str) -> List[str]:
        """
        Get all sheet names that should have a specific chart type.
        
        Args:
            chart_type: Type of chart ('acf_pacf', 'arima_forecast', 'dashboard')
        
        Returns:
            List of sheet names that should have this chart type
        """
        sheets = []
        
        for sheet_name, sheet_config in self.sheet_configs.items():
            if (sheet_config.get('enabled', True) and 
                sheet_config.get('charts', {}).get(chart_type, False)):
                sheets.append(sheet_name)
        
        return sheets
    
    def validate_configuration(self) -> List[str]:
        """
        Validate the chart configuration and return any issues found.
        
        Returns:
            List of validation error messages (empty if no issues)
        """
        issues = []
        
        for sheet_name, sheet_config in self.sheet_configs.items():
            # Check if charts configuration exists
            if 'charts' not in sheet_config:
                issues.append(f"Sheet '{sheet_name}' missing 'charts' configuration")
                continue
            
            charts_config = sheet_config['charts']
            
            # Validate chart type fields
            expected_chart_types = {'acf_pacf', 'arima_forecast', 'dashboard'}
            for chart_type in expected_chart_types:
                if chart_type not in charts_config:
                    issues.append(f"Sheet '{sheet_name}' missing '{chart_type}' chart configuration")
                elif not isinstance(charts_config[chart_type], bool):
                    issues.append(f"Sheet '{sheet_name}' chart '{chart_type}' must be boolean")
            
            # Validate time_scale
            if 'time_scale' not in sheet_config:
                issues.append(f"Sheet '{sheet_name}' missing 'time_scale' field")
            
            # Check ARIMA forecast consistency
            if (charts_config.get('arima_forecast', False) and 
                sheet_config.get('time_scale') not in self.get_enabled_time_scales()):
                issues.append(f"Sheet '{sheet_name}' has ARIMA forecast enabled but time scale not in enabled list")
        
        return issues


# Global instance for easy access
_chart_config_helper = None

def get_chart_config_helper() -> ChartConfigHelper:
    """Get the global chart configuration helper instance."""
    global _chart_config_helper
    if _chart_config_helper is None:
        _chart_config_helper = ChartConfigHelper()
    return _chart_config_helper

def should_add_chart(sheet_name: str, chart_type: str) -> bool:
    """
    Convenience function to check if a chart should be added to a sheet.
    
    Args:
        sheet_name: Name of the Excel sheet
        chart_type: Type of chart ('acf_pacf', 'arima_forecast', 'dashboard')
    
    Returns:
        True if the chart should be added, False otherwise
    """
    return get_chart_config_helper().should_add_chart(sheet_name, chart_type)

def get_sheet_time_scale(sheet_name: str) -> Optional[str]:
    """
    Convenience function to get the time scale for a sheet.
    
    Args:
        sheet_name: Name of the Excel sheet
    
    Returns:
        Time scale string or None if not configured
    """
    return get_chart_config_helper().get_sheet_time_scale(sheet_name)

def should_add_acf_pacf_columns(sheet_name: str) -> bool:
    """
    Convenience function to check if ACF/PACF columns should be added to a sheet.
    
    Args:
        sheet_name: Name of the Excel sheet
    
    Returns:
        True if ACF/PACF columns should be added, False otherwise
    """
    return get_chart_config_helper().should_add_acf_pacf_columns(sheet_name)

def should_add_arima_columns(sheet_name: str) -> bool:
    """
    Convenience function to check if ARIMA columns should be added to a sheet.
    
    Args:
        sheet_name: Name of the Excel sheet
    
    Returns:
        True if ARIMA columns should be added, False otherwise
    """
    return get_chart_config_helper().should_add_arima_columns(sheet_name)
