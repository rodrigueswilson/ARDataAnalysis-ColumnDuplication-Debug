"""
Sheet Factory for AR Data Analysis

This module implements the Factory pattern for creating Excel sheets based on
configuration. It provides centralized sheet creation logic, validation,
and management capabilities for the unified sheet management system.

Key Features:
- Configuration-driven sheet creation
- Sheet validation and dependency resolution
- Enable/disable sheet management
- Category-based organization
- Order management and conflict resolution
"""

import json
import logging
from typing import Dict, List, Optional, Tuple, Any
from pathlib import Path

class SheetFactory:
    """
    Factory class for creating and managing Excel sheets based on configuration.
    
    This class provides centralized logic for:
    - Loading and validating sheet configurations
    - Creating sheets based on pipeline data
    - Managing sheet dependencies and ordering
    - Enabling/disabling sheets dynamically
    """
    
    def __init__(self, config_path: str = "report_config.json"):
        """
        Initialize the SheetFactory with configuration.
        
        Args:
            config_path: Path to the report configuration JSON file
        """
        self.config_path = config_path
        self.config = self._load_config()
        self.logger = logging.getLogger(__name__)
        
        # Validate configuration on initialization
        self._validate_config()
    
    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from JSON file."""
        try:
            with open(self.config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            self.logger.error(f"Configuration file not found: {self.config_path}")
            raise
        except json.JSONDecodeError as e:
            self.logger.error(f"Invalid JSON in configuration file: {e}")
            raise
    
    def _validate_config(self) -> None:
        """Validate the loaded configuration."""
        required_sections = ['pipeline_modules', 'sheet_categories', 'sheets']
        for section in required_sections:
            if section not in self.config:
                raise ValueError(f"Missing required configuration section: {section}")
        
        # Validate sheet configurations
        validation_rules = self.config.get('validation_rules', {})
        if validation_rules.get('require_enabled_module', True):
            self._validate_sheet_modules()
        
        if validation_rules.get('validate_category', True):
            self._validate_sheet_categories()
    
    def _validate_sheet_modules(self) -> None:
        """Validate that all sheet modules are defined and enabled."""
        enabled_modules = {
            name for name, module_config in self.config['pipeline_modules'].items()
            if module_config.get('enabled', True)
        }
        
        for sheet in self.config['sheets']:
            if sheet.get('enabled', True):
                module = sheet.get('module')
                if module and module not in enabled_modules:
                    self.logger.warning(
                        f"Sheet '{sheet['sheet_name']}' references disabled module '{module}'"
                    )
    
    def _validate_sheet_categories(self) -> None:
        """Validate that all sheet categories are defined."""
        valid_categories = set(self.config['sheet_categories'].keys())
        
        for sheet in self.config['sheets']:
            category = sheet.get('category')
            if category and category not in valid_categories:
                raise ValueError(
                    f"Sheet '{sheet['sheet_name']}' has invalid category '{category}'"
                )
    
    def get_enabled_sheets(self) -> List[Dict[str, Any]]:
        """
        Get list of enabled sheets sorted by category and order.
        
        Returns:
            List of enabled sheet configurations
        """
        enabled_sheets = [
            sheet for sheet in self.config['sheets']
            if sheet.get('enabled', True)
        ]
        
        # Sort by category order first, then by sheet order
        category_orders = {
            name: config.get('order', 999)
            for name, config in self.config['sheet_categories'].items()
        }
        
        def sort_key(sheet):
            category = sheet.get('category', 'unknown')
            category_order = category_orders.get(category, 999)
            sheet_order = sheet.get('order', 999)
            return (category_order, sheet_order)
        
        return sorted(enabled_sheets, key=sort_key)
    
    def get_sheets_by_category(self, category: str) -> List[Dict[str, Any]]:
        """
        Get all enabled sheets in a specific category.
        
        Args:
            category: Category name to filter by
            
        Returns:
            List of sheets in the specified category
        """
        return [
            sheet for sheet in self.get_enabled_sheets()
            if sheet.get('category') == category
        ]
    
    def get_sheet_by_pipeline(self, pipeline_name: str) -> Optional[Dict[str, Any]]:
        """
        Get sheet configuration by pipeline name.
        
        Args:
            pipeline_name: Name of the pipeline
            
        Returns:
            Sheet configuration or None if not found
        """
        for sheet in self.config['sheets']:
            if sheet.get('pipeline') == pipeline_name:
                return sheet
        return None
    
    def enable_sheet(self, sheet_name: str) -> bool:
        """
        Enable a sheet by name.
        
        Args:
            sheet_name: Name of the sheet to enable
            
        Returns:
            True if successful, False otherwise
        """
        for sheet in self.config['sheets']:
            if sheet.get('sheet_name') == sheet_name:
                sheet['enabled'] = True
                self._save_config()
                return True
        return False
    
    def disable_sheet(self, sheet_name: str) -> bool:
        """
        Disable a sheet by name.
        
        Args:
            sheet_name: Name of the sheet to disable
            
        Returns:
            True if successful, False otherwise
        """
        for sheet in self.config['sheets']:
            if sheet.get('sheet_name') == sheet_name:
                sheet['enabled'] = False
                self._save_config()
                return True
        return False
    
    def reorder_sheet(self, sheet_name: str, new_order: int) -> bool:
        """
        Change the order of a sheet.
        
        Args:
            sheet_name: Name of the sheet to reorder
            new_order: New order value
            
        Returns:
            True if successful, False otherwise
        """
        for sheet in self.config['sheets']:
            if sheet.get('sheet_name') == sheet_name:
                sheet['order'] = new_order
                self._save_config()
                return True
        return False
    
    def add_sheet(self, sheet_config: Dict[str, Any]) -> bool:
        """
        Add a new sheet configuration.
        
        Args:
            sheet_config: Complete sheet configuration dictionary
            
        Returns:
            True if successful, False otherwise
        """
        # Validate required fields
        required_fields = ['pipeline', 'sheet_name', 'category', 'module']
        for field in required_fields:
            if field not in sheet_config:
                self.logger.error(f"Missing required field '{field}' in sheet config")
                return False
        
        # Set defaults
        sheet_config.setdefault('enabled', True)
        sheet_config.setdefault('order', 999)
        sheet_config.setdefault('description', '')
        
        # Add to configuration
        self.config['sheets'].append(sheet_config)
        self._save_config()
        return True
    
    def remove_sheet(self, sheet_name: str) -> bool:
        """
        Remove a sheet configuration.
        
        Args:
            sheet_name: Name of the sheet to remove
            
        Returns:
            True if successful, False otherwise
        """
        original_count = len(self.config['sheets'])
        self.config['sheets'] = [
            sheet for sheet in self.config['sheets']
            if sheet.get('sheet_name') != sheet_name
        ]
        
        if len(self.config['sheets']) < original_count:
            self._save_config()
            return True
        return False
    
    def validate_dependencies(self) -> List[str]:
        """
        Validate sheet dependencies and return list of issues.
        
        Returns:
            List of validation error messages
        """
        issues = []
        
        # Check for missing pipelines (would need pipeline registry access)
        # This is a placeholder for future dependency validation
        
        # Check for duplicate orders within categories
        if self.config.get('validation_rules', {}).get('enforce_unique_orders', False):
            category_orders = {}
            for sheet in self.get_enabled_sheets():
                category = sheet.get('category', 'unknown')
                order = sheet.get('order', 999)
                
                if category not in category_orders:
                    category_orders[category] = {}
                
                if order in category_orders[category]:
                    issues.append(
                        f"Duplicate order {order} in category '{category}': "
                        f"'{sheet['sheet_name']}' and '{category_orders[category][order]}'"
                    )
                else:
                    category_orders[category][order] = sheet['sheet_name']
        
        return issues
    
    def _save_config(self) -> None:
        """Save configuration back to file."""
        if self.config.get('sheet_management', {}).get('backup_on_changes', True):
            self._backup_config()
        
        try:
            with open(self.config_path, 'w') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            self.logger.error(f"Failed to save configuration: {e}")
            raise
    
    def _backup_config(self) -> None:
        """Create a backup of the current configuration."""
        backup_path = f"{self.config_path}.backup"
        try:
            with open(backup_path, 'w') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            self.logger.warning(f"Failed to create config backup: {e}")
    
    def get_summary(self) -> Dict[str, Any]:
        """
        Get a summary of the current configuration.
        
        Returns:
            Dictionary with configuration summary
        """
        enabled_sheets = self.get_enabled_sheets()
        
        category_counts = {}
        for sheet in enabled_sheets:
            category = sheet.get('category', 'unknown')
            category_counts[category] = category_counts.get(category, 0) + 1
        
        return {
            'total_sheets': len(self.config['sheets']),
            'enabled_sheets': len(enabled_sheets),
            'disabled_sheets': len(self.config['sheets']) - len(enabled_sheets),
            'categories': len(self.config['sheet_categories']),
            'modules': len(self.config['pipeline_modules']),
            'sheets_by_category': category_counts,
            'validation_issues': len(self.validate_dependencies())
        }
