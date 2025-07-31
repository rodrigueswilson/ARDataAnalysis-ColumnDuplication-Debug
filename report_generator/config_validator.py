"""
Configuration Validator Module
==============================

This module provides comprehensive validation for report configurations,
including dependency resolution, consistency checks, and validation rules
for the AR Data Analysis system.
"""

import json
from pathlib import Path
from typing import Dict, List, Set, Tuple, Any, Optional
from pipelines import PIPELINES  # Now using modular pipelines/ package


class ConfigurationValidator:
    """
    Validates report configurations and resolves dependencies between sheets,
    pipelines, and modules.
    """
    
    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize the configuration validator.
        
        Args:
            config_path: Path to the report configuration file
        """
        if config_path is None:
            config_path = Path(__file__).parent.parent / "report_config.json"
        
        self.config_path = Path(config_path)
        self.config = None
        self.validation_errors = []
        self.validation_warnings = []
        
    def load_config(self) -> bool:
        """
        Load the configuration file.
        
        Returns:
            bool: True if loaded successfully, False otherwise
        """
        try:
            with open(self.config_path, 'r') as f:
                self.config = json.load(f)
            return True
        except Exception as e:
            self.validation_errors.append(f"Failed to load configuration: {e}")
            return False
    
    def validate_all(self) -> Dict[str, Any]:
        """
        Perform comprehensive validation of the configuration.
        
        Returns:
            dict: Validation results with errors, warnings, and summary
        """
        if not self.load_config():
            return self._create_validation_result(False)
        
        self.validation_errors = []
        self.validation_warnings = []
        
        # Run all validation checks
        self._validate_structure()
        self._validate_sheets()
        self._validate_pipelines()
        self._validate_dependencies()
        self._validate_forecast_options()
        self._validate_pipeline_modules()
        self._validate_consistency()
        
        return self._create_validation_result(len(self.validation_errors) == 0)
    
    def _validate_structure(self):
        """Validate the basic structure of the configuration."""
        required_sections = [
            'sheets', 'forecast_options', 'pipeline_modules',
            'sheet_categories', 'management_options', 'validation_rules'
        ]
        
        for section in required_sections:
            if section not in self.config:
                self.validation_errors.append(f"Missing required section: {section}")
    
    def _validate_sheets(self):
        """Validate sheet configurations."""
        if 'sheets' not in self.config:
            return
        
        sheets = self.config['sheets']
        sheet_names = set()
        sheet_orders = {}
        
        for i, sheet in enumerate(sheets):
            # Check required fields
            required_fields = ['name', 'pipeline', 'enabled', 'category', 'order']
            for field in required_fields:
                if field not in sheet:
                    self.validation_errors.append(
                        f"Sheet {i}: Missing required field '{field}'"
                    )
            
            # Check for duplicate names
            name = sheet.get('name', f'sheet_{i}')
            if name in sheet_names:
                self.validation_errors.append(f"Duplicate sheet name: {name}")
            sheet_names.add(name)
            
            # Check for duplicate orders within categories
            category = sheet.get('category', 'unknown')
            order = sheet.get('order', 0)
            
            if category not in sheet_orders:
                sheet_orders[category] = set()
            
            if order in sheet_orders[category]:
                self.validation_warnings.append(
                    f"Duplicate order {order} in category {category}"
                )
            sheet_orders[category].add(order)
            
            # Validate pipeline exists
            pipeline = sheet.get('pipeline')
            if pipeline and pipeline not in PIPELINES:
                self.validation_errors.append(
                    f"Sheet '{name}': Pipeline '{pipeline}' not found"
                )
            
            # Validate category exists
            if 'sheet_categories' in self.config:
                valid_categories = list(self.config['sheet_categories'].keys())
                if category not in valid_categories:
                    self.validation_errors.append(
                        f"Sheet '{name}': Invalid category '{category}'"
                    )
    
    def _validate_pipelines(self):
        """Validate that all referenced pipelines exist."""
        if 'sheets' not in self.config:
            return
        
        referenced_pipelines = set()
        for sheet in self.config['sheets']:
            pipeline = sheet.get('pipeline')
            if pipeline:
                referenced_pipelines.add(pipeline)
        
        missing_pipelines = referenced_pipelines - set(PIPELINES.keys())
        for pipeline in missing_pipelines:
            self.validation_errors.append(f"Referenced pipeline not found: {pipeline}")
        
        # Check for unused pipelines
        unused_pipelines = set(PIPELINES.keys()) - referenced_pipelines
        if unused_pipelines:
            self.validation_warnings.append(
                f"Unused pipelines: {', '.join(sorted(unused_pipelines))}"
            )
    
    def _validate_dependencies(self):
        """Validate sheet dependencies and resolve dependency chains."""
        if 'sheets' not in self.config:
            return
        
        # Build dependency graph
        dependencies = {}
        sheet_names = set()
        
        for sheet in self.config['sheets']:
            name = sheet.get('name')
            if name:
                sheet_names.add(name)
                deps = sheet.get('dependencies', [])
                dependencies[name] = deps
        
        # Check for missing dependencies
        for sheet_name, deps in dependencies.items():
            for dep in deps:
                if dep not in sheet_names:
                    self.validation_errors.append(
                        f"Sheet '{sheet_name}': Dependency '{dep}' not found"
                    )
        
        # Check for circular dependencies
        circular_deps = self._find_circular_dependencies(dependencies)
        for cycle in circular_deps:
            self.validation_errors.append(
                f"Circular dependency detected: {' -> '.join(cycle)}"
            )
    
    def _validate_forecast_options(self):
        """Validate forecast configuration options."""
        if 'forecast_options' not in self.config:
            return
        
        forecast_options = self.config['forecast_options']
        
        # Check required fields
        required_fields = ['enabled', 'time_scales', 'target_metrics']
        for field in required_fields:
            if field not in forecast_options:
                self.validation_errors.append(
                    f"Forecast options: Missing required field '{field}'"
                )
        
        # Validate time scales
        valid_time_scales = ['daily', 'weekly', 'biweekly', 'monthly', 'period']
        time_scales = forecast_options.get('time_scales', [])
        
        for scale in time_scales:
            if scale not in valid_time_scales:
                self.validation_errors.append(
                    f"Forecast options: Invalid time scale '{scale}'"
                )
        
        # Validate target metrics
        target_metrics = forecast_options.get('target_metrics', [])
        if not target_metrics:
            self.validation_warnings.append(
                "Forecast options: No target metrics specified"
            )
    
    def _validate_pipeline_modules(self):
        """Validate pipeline module configurations."""
        if 'pipeline_modules' not in self.config:
            return
        
        pipeline_modules = self.config['pipeline_modules']
        
        for module_name, module_config in pipeline_modules.items():
            # Check required fields
            required_fields = ['enabled', 'description']
            for field in required_fields:
                if field not in module_config:
                    self.validation_errors.append(
                        f"Pipeline module '{module_name}': Missing field '{field}'"
                    )
    
    def _validate_consistency(self):
        """Validate consistency across different configuration sections."""
        if not all(section in self.config for section in ['sheets', 'sheet_categories']):
            return
        
        # Check that all sheet categories are used
        defined_categories = set(self.config['sheet_categories'].keys())
        used_categories = set(sheet.get('category') for sheet in self.config['sheets'])
        
        unused_categories = defined_categories - used_categories
        if unused_categories:
            self.validation_warnings.append(
                f"Unused sheet categories: {', '.join(sorted(unused_categories))}"
            )
        
        undefined_categories = used_categories - defined_categories
        for category in undefined_categories:
            if category:  # Skip None values
                self.validation_errors.append(
                    f"Sheet category '{category}' used but not defined"
                )
    
    def _find_circular_dependencies(self, dependencies: Dict[str, List[str]]) -> List[List[str]]:
        """
        Find circular dependencies using depth-first search.
        
        Args:
            dependencies: Dictionary mapping sheet names to their dependencies
            
        Returns:
            List of circular dependency chains
        """
        circular_deps = []
        visited = set()
        rec_stack = set()
        
        def dfs(node: str, path: List[str]) -> bool:
            if node in rec_stack:
                # Found a cycle
                cycle_start = path.index(node)
                circular_deps.append(path[cycle_start:] + [node])
                return True
            
            if node in visited:
                return False
            
            visited.add(node)
            rec_stack.add(node)
            
            for dep in dependencies.get(node, []):
                if dfs(dep, path + [node]):
                    return True
            
            rec_stack.remove(node)
            return False
        
        for node in dependencies:
            if node not in visited:
                dfs(node, [])
        
        return circular_deps
    
    def _create_validation_result(self, success: bool) -> Dict[str, Any]:
        """Create a validation result dictionary."""
        return {
            'success': success,
            'errors': self.validation_errors,
            'warnings': self.validation_warnings,
            'summary': {
                'total_errors': len(self.validation_errors),
                'total_warnings': len(self.validation_warnings),
                'sheets_validated': len(self.config.get('sheets', [])) if self.config else 0,
                'pipelines_available': len(PIPELINES),
                'categories_defined': len(self.config.get('sheet_categories', {})) if self.config else 0
            }
        }
    
    def get_dependency_order(self) -> List[str]:
        """
        Get the correct order for sheet creation based on dependencies.
        
        Returns:
            List of sheet names in dependency order
        """
        if not self.config or 'sheets' not in self.config:
            return []
        
        # Build dependency graph
        dependencies = {}
        all_sheets = set()
        
        for sheet in self.config['sheets']:
            name = sheet.get('name')
            if name:
                all_sheets.add(name)
                dependencies[name] = sheet.get('dependencies', [])
        
        # Topological sort
        result = []
        visited = set()
        temp_visited = set()
        
        def visit(node: str):
            if node in temp_visited:
                # Circular dependency - skip this node
                return
            if node in visited:
                return
            
            temp_visited.add(node)
            for dep in dependencies.get(node, []):
                if dep in all_sheets:
                    visit(dep)
            temp_visited.remove(node)
            visited.add(node)
            result.append(node)
        
        for sheet in all_sheets:
            if sheet not in visited:
                visit(sheet)
        
        return result
    
    def validate_and_fix(self) -> Dict[str, Any]:
        """
        Validate configuration and attempt to fix common issues.
        
        Returns:
            dict: Validation results with fixes applied
        """
        result = self.validate_all()
        
        if not result['success']:
            # Attempt to fix common issues
            fixes_applied = []
            
            # Fix missing required sections
            if 'sheets' not in self.config:
                self.config['sheets'] = []
                fixes_applied.append("Added missing 'sheets' section")
            
            # Fix duplicate sheet orders
            self._fix_duplicate_orders()
            fixes_applied.extend(self._get_order_fixes())
            
            # Re-validate after fixes
            result = self.validate_all()
            result['fixes_applied'] = fixes_applied
        
        return result
    
    def _fix_duplicate_orders(self):
        """Fix duplicate sheet orders within categories."""
        if 'sheets' not in self.config:
            return
        
        category_orders = {}
        
        for sheet in self.config['sheets']:
            category = sheet.get('category', 'unknown')
            order = sheet.get('order', 0)
            
            if category not in category_orders:
                category_orders[category] = set()
            
            # If order is already used, find the next available order
            while order in category_orders[category]:
                order += 1
            
            sheet['order'] = order
            category_orders[category].add(order)
    
    def _get_order_fixes(self) -> List[str]:
        """Get list of order fixes that were applied."""
        # This would track fixes during _fix_duplicate_orders
        # For now, return empty list
        return []


def validate_configuration(config_path: Optional[str] = None) -> Dict[str, Any]:
    """
    Convenience function to validate a configuration file.
    
    Args:
        config_path: Path to the configuration file
        
    Returns:
        dict: Validation results
    """
    validator = ConfigurationValidator(config_path)
    return validator.validate_all()


def get_dependency_order(config_path: Optional[str] = None) -> List[str]:
    """
    Get the correct order for sheet creation based on dependencies.
    
    Args:
        config_path: Path to the configuration file
        
    Returns:
        List of sheet names in dependency order
    """
    validator = ConfigurationValidator(config_path)
    validator.load_config()
    return validator.get_dependency_order()
