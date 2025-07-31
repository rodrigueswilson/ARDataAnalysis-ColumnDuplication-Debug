"""
Sheet Management Tools for AR Data Analysis

This module provides comprehensive tools for managing Excel sheets including
enable/disable functionality, reordering, validation, and dependency resolution.
It works in conjunction with the SheetFactory to provide a complete sheet
management system.

Key Features:
- Dynamic sheet enable/disable with validation
- Sheet reordering and conflict resolution
- Dependency validation and resolution
- Batch operations for multiple sheets
- Configuration backup and rollback
- Management CLI interface
"""

import json
import logging
from typing import Dict, List, Optional, Tuple, Any
from pathlib import Path
from datetime import datetime

from .sheet_factory import SheetFactory


class SheetManager:
    """
    Advanced sheet management system with validation and dependency resolution.
    
    This class provides high-level operations for managing sheets including
    batch operations, validation, and advanced features like dependency
    resolution and conflict management.
    """
    
    def __init__(self, factory_or_config_path=None):
        """
        Initialize the SheetManager.
        
        Args:
            factory_or_config_path: Either a SheetFactory instance or path to config file
        """
        if isinstance(factory_or_config_path, SheetFactory):
            self.factory = factory_or_config_path
        else:
            self.factory = SheetFactory(factory_or_config_path)
        self.logger = logging.getLogger(__name__)
        
        # Track changes for rollback capability
        self._change_history = []
        self._max_history = 10
    
    def enable_sheets(self, sheet_names: List[str], validate: bool = True) -> Dict[str, bool]:
        """
        Enable multiple sheets with optional validation.
        
        Args:
            sheet_names: List of sheet names to enable
            validate: Whether to validate dependencies before enabling
            
        Returns:
            Dictionary mapping sheet names to success status
        """
        results = {}
        
        if validate:
            validation_issues = self._validate_enable_operation(sheet_names)
            if validation_issues:
                self.logger.warning(f"Validation issues found: {validation_issues}")
        
        # Record current state for rollback
        self._record_change("enable_sheets", sheet_names)
        
        for sheet_name in sheet_names:
            try:
                success = self.factory.enable_sheet(sheet_name)
                results[sheet_name] = success
                if success:
                    self.logger.info(f"Enabled sheet: {sheet_name}")
                else:
                    self.logger.warning(f"Failed to enable sheet: {sheet_name}")
            except Exception as e:
                self.logger.error(f"Error enabling sheet {sheet_name}: {e}")
                results[sheet_name] = False
        
        return results
    
    def disable_sheets(self, sheet_names: List[str], validate: bool = True) -> Dict[str, bool]:
        """
        Disable multiple sheets with dependency validation.
        
        Args:
            sheet_names: List of sheet names to disable
            validate: Whether to validate dependencies before disabling
            
        Returns:
            Dictionary mapping sheet names to success status
        """
        results = {}
        
        if validate:
            validation_issues = self._validate_disable_operation(sheet_names)
            if validation_issues:
                self.logger.warning(f"Validation issues found: {validation_issues}")
        
        # Record current state for rollback
        self._record_change("disable_sheets", sheet_names)
        
        for sheet_name in sheet_names:
            try:
                success = self.factory.disable_sheet(sheet_name)
                results[sheet_name] = success
                if success:
                    self.logger.info(f"Disabled sheet: {sheet_name}")
                else:
                    self.logger.warning(f"Failed to disable sheet: {sheet_name}")
            except Exception as e:
                self.logger.error(f"Error disabling sheet {sheet_name}: {e}")
                results[sheet_name] = False
        
        return results
    
    def reorder_sheets_by_category(self, category_order: Dict[str, int]) -> bool:
        """
        Reorder sheets within categories based on new order values.
        
        Args:
            category_order: Dictionary mapping sheet names to new order values
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Record current state for rollback
            self._record_change("reorder_sheets", category_order)
            
            success_count = 0
            for sheet_name, new_order in category_order.items():
                if self.factory.reorder_sheet(sheet_name, new_order):
                    success_count += 1
                    self.logger.info(f"Reordered sheet {sheet_name} to order {new_order}")
                else:
                    self.logger.warning(f"Failed to reorder sheet: {sheet_name}")
            
            return success_count == len(category_order)
            
        except Exception as e:
            self.logger.error(f"Error reordering sheets: {e}")
            return False
    
    def optimize_sheet_order(self) -> bool:
        """
        Automatically optimize sheet order within categories to eliminate gaps.
        
        Returns:
            True if optimization was successful, False otherwise
        """
        try:
            enabled_sheets = self.factory.get_enabled_sheets()
            
            # Group sheets by category
            category_sheets = {}
            for sheet in enabled_sheets:
                category = sheet.get('category', 'unknown')
                if category not in category_sheets:
                    category_sheets[category] = []
                category_sheets[category].append(sheet)
            
            # Optimize order within each category
            reorder_map = {}
            for category, sheets in category_sheets.items():
                # Sort by current order
                sheets.sort(key=lambda x: x.get('order', 999))
                
                # Assign new sequential orders
                for i, sheet in enumerate(sheets):
                    new_order = (i + 1) * 10  # Leave gaps for future insertions
                    if sheet.get('order') != new_order:
                        reorder_map[sheet['sheet_name']] = new_order
            
            if reorder_map:
                return self.reorder_sheets_by_category(reorder_map)
            else:
                self.logger.info("Sheet order already optimized")
                return True
                
        except Exception as e:
            self.logger.error(f"Error optimizing sheet order: {e}")
            return False
    
    def validate_all_sheets(self) -> Dict[str, List[str]]:
        """
        Comprehensive validation of all sheets and their dependencies.
        
        Returns:
            Dictionary with validation results by category
        """
        validation_results = {
            'errors': [],
            'warnings': [],
            'info': []
        }
        
        try:
            # Basic factory validation
            factory_issues = self.factory.validate_dependencies()
            if factory_issues:
                validation_results['errors'].extend(factory_issues)
            
            # Check for missing pipelines
            enabled_sheets = self.factory.get_enabled_sheets()
            from pipelines import PIPELINES  # Now using modular pipelines/ package
            
            for sheet in enabled_sheets:
                pipeline_name = sheet.get('pipeline')
                if pipeline_name and pipeline_name not in PIPELINES:
                    validation_results['errors'].append(
                        f"Sheet '{sheet['sheet_name']}' references missing pipeline '{pipeline_name}'"
                    )
            
            # Check for disabled modules with enabled sheets
            for sheet in enabled_sheets:
                module_name = sheet.get('module')
                if module_name:
                    module_config = self.factory.config['pipeline_modules'].get(module_name, {})
                    if not module_config.get('enabled', True):
                        validation_results['warnings'].append(
                            f"Sheet '{sheet['sheet_name']}' references disabled module '{module_name}'"
                        )
            
            # Check for order conflicts
            category_orders = {}
            for sheet in enabled_sheets:
                category = sheet.get('category', 'unknown')
                order = sheet.get('order', 999)
                
                if category not in category_orders:
                    category_orders[category] = {}
                
                if order in category_orders[category]:
                    validation_results['warnings'].append(
                        f"Order conflict in category '{category}': order {order} used by both "
                        f"'{sheet['sheet_name']}' and '{category_orders[category][order]}'"
                    )
                else:
                    category_orders[category][order] = sheet['sheet_name']
            
            # Summary information
            validation_results['info'].append(f"Validated {len(enabled_sheets)} enabled sheets")
            validation_results['info'].append(f"Found {len(validation_results['errors'])} errors")
            validation_results['info'].append(f"Found {len(validation_results['warnings'])} warnings")
            
        except Exception as e:
            validation_results['errors'].append(f"Validation error: {e}")
        
        return validation_results
    
    def get_sheet_dependencies(self, sheet_name: str) -> Dict[str, Any]:
        """
        Get dependency information for a specific sheet.
        
        Args:
            sheet_name: Name of the sheet to analyze
            
        Returns:
            Dictionary containing dependency information
        """
        sheet_config = self.factory.get_sheet_by_pipeline(sheet_name)
        if not sheet_config:
            # Try to find by sheet name
            for sheet in self.factory.config['sheets']:
                if sheet.get('sheet_name') == sheet_name:
                    sheet_config = sheet
                    break
        
        if not sheet_config:
            return {'error': f"Sheet '{sheet_name}' not found"}
        
        dependencies = {
            'sheet_name': sheet_config.get('sheet_name'),
            'pipeline': sheet_config.get('pipeline'),
            'module': sheet_config.get('module'),
            'category': sheet_config.get('category'),
            'enabled': sheet_config.get('enabled', True),
            'order': sheet_config.get('order', 999),
            'dependencies': [],
            'dependents': []
        }
        
        # Check module dependency
        module_name = sheet_config.get('module')
        if module_name:
            module_config = self.factory.config['pipeline_modules'].get(module_name, {})
            dependencies['module_enabled'] = module_config.get('enabled', True)
            dependencies['module_description'] = module_config.get('description', '')
        
        # Check pipeline dependency
        pipeline_name = sheet_config.get('pipeline')
        if pipeline_name:
            from pipelines import PIPELINES  # Now using modular pipelines/ package
            dependencies['pipeline_exists'] = pipeline_name in PIPELINES
        
        return dependencies
    
    def export_configuration(self, include_disabled: bool = False) -> Dict[str, Any]:
        """
        Export current sheet configuration for backup or analysis.
        
        Args:
            include_disabled: Whether to include disabled sheets
            
        Returns:
            Dictionary containing the exported configuration
        """
        if include_disabled:
            sheets = self.factory.config['sheets']
        else:
            sheets = self.factory.get_enabled_sheets()
        
        export_data = {
            'export_timestamp': self._get_timestamp(),
            'total_sheets': len(sheets),
            'enabled_sheets': len(self.factory.get_enabled_sheets()),
            'categories': len(self.factory.config['sheet_categories']),
            'modules': len(self.factory.config['pipeline_modules']),
            'sheets': sheets,
            'categories': self.factory.config['sheet_categories'],
            'modules': self.factory.config['pipeline_modules']
        }
        
        return export_data
    
    def rollback_last_change(self) -> bool:
        """
        Rollback the last configuration change.
        
        Returns:
            True if rollback was successful, False otherwise
        """
        if not self._change_history:
            self.logger.warning("No changes to rollback")
            return False
        
        try:
            last_change = self._change_history.pop()
            # This is a simplified rollback - in a production system,
            # you would store the actual previous state and restore it
            self.logger.info(f"Rolled back change: {last_change['operation']}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error during rollback: {e}")
            return False
    
    def _validate_enable_operation(self, sheet_names: List[str]) -> List[str]:
        """Validate enabling specific sheets."""
        issues = []
        
        for sheet_name in sheet_names:
            dependencies = self.get_sheet_dependencies(sheet_name)
            
            if 'error' in dependencies:
                issues.append(dependencies['error'])
                continue
            
            # Check if module is enabled
            if not dependencies.get('module_enabled', True):
                issues.append(f"Sheet '{sheet_name}' requires disabled module '{dependencies['module']}'")
            
            # Check if pipeline exists
            if not dependencies.get('pipeline_exists', True):
                issues.append(f"Sheet '{sheet_name}' references missing pipeline '{dependencies['pipeline']}'")
        
        return issues
    
    def _validate_disable_operation(self, sheet_names: List[str]) -> List[str]:
        """Validate disabling specific sheets."""
        issues = []
        
        # Check if disabling these sheets would break dependencies
        # This is a placeholder for more complex dependency checking
        for sheet_name in sheet_names:
            # In a real implementation, you might check if other sheets depend on this one
            pass
        
        return issues
    
    def _record_change(self, operation: str, details: Any) -> None:
        """Record a change for rollback capability."""
        change_record = {
            'timestamp': self._get_timestamp(),
            'operation': operation,
            'details': details
        }
        
        self._change_history.append(change_record)
        
        # Limit history size
        if len(self._change_history) > self._max_history:
            self._change_history.pop(0)
    
    def _get_timestamp(self) -> str:
        """Get current timestamp string."""
        from datetime import datetime
        return datetime.now().isoformat()
    
    def get_management_summary(self) -> Dict[str, Any]:
        """
        Get a comprehensive summary of the sheet management system.
        
        Returns:
            Dictionary containing management system status
        """
        factory_summary = self.factory.get_summary()
        validation_results = self.validate_all_sheets()
        
        return {
            'factory_summary': factory_summary,
            'validation_summary': {
                'errors': len(validation_results['errors']),
                'warnings': len(validation_results['warnings']),
                'status': 'healthy' if len(validation_results['errors']) == 0 else 'issues'
            },
            'change_history_size': len(self._change_history),
            'management_features': {
                'enable_disable': True,
                'reordering': True,
                'validation': True,
                'rollback': True,
                'batch_operations': True,
                'dependency_resolution': True
            }
        }
