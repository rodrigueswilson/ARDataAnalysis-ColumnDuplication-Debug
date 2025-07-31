"""
Command Line Interface for Sheet Management

This module provides a comprehensive CLI for managing Excel sheets in the
AR Data Analysis system. It offers commands for enabling/disabling sheets,
reordering, validation, and other management operations.

Usage:
    python sheet_management_cli.py <command> [options]

Commands:
    list        - List all sheets with their status
    enable      - Enable one or more sheets
    disable     - Disable one or more sheets
    reorder     - Reorder sheets within categories
    validate    - Validate sheet configuration
    status      - Show system status
    export      - Export configuration
    optimize    - Optimize sheet order
"""

import argparse
import json
import sys
from pathlib import Path
from typing import List, Dict, Any

# Add the project root to the path
sys.path.insert(0, str(Path(__file__).parent))

from report_generator.sheet_manager import SheetManager
from report_generator.sheet_factory import SheetFactory


class SheetManagementCLI:
    """Command Line Interface for sheet management operations."""
    
    def __init__(self):
        """Initialize the CLI with a SheetManager instance."""
        self.manager = SheetManager()
        self.factory = SheetFactory()
    
    def list_sheets(self, args) -> None:
        """List all sheets with their current status."""
        print("=" * 80)
        print("SHEET CONFIGURATION STATUS")
        print("=" * 80)
        
        try:
            # Get sheets by category
            categories = ['time_series', 'summary', 'detailed', 'quality']
            
            for category in categories:
                sheets = self.factory.get_sheets_by_category(category)
                if not sheets:
                    continue
                
                print(f"\n📊 {category.upper().replace('_', ' ')} ({len(sheets)} sheets)")
                print("-" * 60)
                
                for sheet in sorted(sheets, key=lambda x: x.get('order', 999)):
                    status = "✓ ENABLED " if sheet.get('enabled', True) else "✗ DISABLED"
                    order = sheet.get('order', 'N/A')
                    module = sheet.get('module', 'N/A')
                    
                    print(f"  {status} [{order:>3}] {sheet['sheet_name']}")
                    print(f"      Pipeline: {sheet.get('pipeline', 'N/A')}")
                    print(f"      Module: {module}")
                    if args.verbose:
                        print(f"      Description: {sheet.get('description', 'N/A')}")
                    print()
            
            # Summary
            summary = self.factory.get_summary()
            print("=" * 80)
            print("SUMMARY")
            print("=" * 80)
            print(f"Total sheets: {summary['total_sheets']}")
            print(f"Enabled sheets: {summary['enabled_sheets']}")
            print(f"Disabled sheets: {summary['disabled_sheets']}")
            print(f"Categories: {summary['categories']}")
            print(f"Pipeline modules: {summary['modules']}")
            
        except Exception as e:
            print(f"❌ Error listing sheets: {e}")
    
    def enable_sheets(self, args) -> None:
        """Enable one or more sheets."""
        print(f"🔄 Enabling sheets: {', '.join(args.sheets)}")
        
        try:
            results = self.manager.enable_sheets(args.sheets, validate=not args.no_validate)
            
            print("\nResults:")
            for sheet_name, success in results.items():
                status = "✅ SUCCESS" if success else "❌ FAILED"
                print(f"  {sheet_name}: {status}")
            
            successful = sum(1 for success in results.values() if success)
            print(f"\n📊 Summary: {successful}/{len(results)} sheets enabled successfully")
            
        except Exception as e:
            print(f"❌ Error enabling sheets: {e}")
    
    def disable_sheets(self, args) -> None:
        """Disable one or more sheets."""
        print(f"🔄 Disabling sheets: {', '.join(args.sheets)}")
        
        try:
            results = self.manager.disable_sheets(args.sheets, validate=not args.no_validate)
            
            print("\nResults:")
            for sheet_name, success in results.items():
                status = "✅ SUCCESS" if success else "❌ FAILED"
                print(f"  {sheet_name}: {status}")
            
            successful = sum(1 for success in results.values() if success)
            print(f"\n📊 Summary: {successful}/{len(results)} sheets disabled successfully")
            
        except Exception as e:
            print(f"❌ Error disabling sheets: {e}")
    
    def reorder_sheets(self, args) -> None:
        """Reorder sheets within categories."""
        print("🔄 Reordering sheets...")
        
        try:
            # Parse reorder specification (format: "sheet_name:new_order,sheet_name:new_order")
            reorder_map = {}
            for spec in args.reorder_spec.split(','):
                if ':' not in spec:
                    print(f"❌ Invalid reorder specification: {spec}")
                    return
                
                sheet_name, order_str = spec.split(':', 1)
                try:
                    new_order = int(order_str)
                    reorder_map[sheet_name.strip()] = new_order
                except ValueError:
                    print(f"❌ Invalid order value: {order_str}")
                    return
            
            success = self.manager.reorder_sheets_by_category(reorder_map)
            
            if success:
                print("✅ Sheets reordered successfully")
                print("\nNew order:")
                for sheet_name, order in reorder_map.items():
                    print(f"  {sheet_name}: {order}")
            else:
                print("❌ Failed to reorder some sheets")
                
        except Exception as e:
            print(f"❌ Error reordering sheets: {e}")
    
    def validate_sheets(self, args) -> None:
        """Validate sheet configuration."""
        print("🔍 Validating sheet configuration...")
        
        try:
            validation_results = self.manager.validate_all_sheets()
            
            print("\n" + "=" * 80)
            print("VALIDATION RESULTS")
            print("=" * 80)
            
            # Errors
            if validation_results['errors']:
                print(f"\n❌ ERRORS ({len(validation_results['errors'])})")
                print("-" * 40)
                for error in validation_results['errors']:
                    print(f"  • {error}")
            
            # Warnings
            if validation_results['warnings']:
                print(f"\n⚠️  WARNINGS ({len(validation_results['warnings'])})")
                print("-" * 40)
                for warning in validation_results['warnings']:
                    print(f"  • {warning}")
            
            # Info
            if validation_results['info']:
                print(f"\nℹ️  INFORMATION")
                print("-" * 40)
                for info in validation_results['info']:
                    print(f"  • {info}")
            
            # Overall status
            if not validation_results['errors'] and not validation_results['warnings']:
                print("\n✅ Configuration is valid - no issues found!")
            elif not validation_results['errors']:
                print("\n⚠️  Configuration has warnings but no critical errors")
            else:
                print("\n❌ Configuration has critical errors that need attention")
                
        except Exception as e:
            print(f"❌ Error during validation: {e}")
    
    def show_status(self, args) -> None:
        """Show comprehensive system status."""
        print("=" * 80)
        print("SHEET MANAGEMENT SYSTEM STATUS")
        print("=" * 80)
        
        try:
            summary = self.manager.get_management_summary()
            
            # Factory summary
            factory = summary['factory_summary']
            print(f"\n📊 SHEET STATISTICS")
            print(f"  Total sheets: {factory['total_sheets']}")
            print(f"  Enabled sheets: {factory['enabled_sheets']}")
            print(f"  Disabled sheets: {factory['disabled_sheets']}")
            print(f"  Categories: {factory['categories']}")
            print(f"  Pipeline modules: {factory['modules']}")
            
            # Sheets by category
            print(f"\n📋 SHEETS BY CATEGORY")
            for category, count in factory['sheets_by_category'].items():
                print(f"  {category}: {count} sheets")
            
            # Validation status
            validation = summary['validation_summary']
            status_icon = "✅" if validation['status'] == 'healthy' else "⚠️"
            print(f"\n🔍 VALIDATION STATUS: {status_icon} {validation['status'].upper()}")
            print(f"  Errors: {validation['errors']}")
            print(f"  Warnings: {validation['warnings']}")
            
            # Management features
            features = summary['management_features']
            print(f"\n⚙️  MANAGEMENT FEATURES")
            for feature, enabled in features.items():
                icon = "✅" if enabled else "❌"
                print(f"  {icon} {feature.replace('_', ' ').title()}")
            
            # Change history
            print(f"\n📝 CHANGE HISTORY")
            print(f"  Recent changes: {summary['change_history_size']}")
            
        except Exception as e:
            print(f"❌ Error getting status: {e}")
    
    def export_config(self, args) -> None:
        """Export configuration to a file."""
        print("📤 Exporting configuration...")
        
        try:
            config_data = self.manager.export_configuration(include_disabled=args.include_disabled)
            
            output_file = args.output or "sheet_config_export.json"
            
            with open(output_file, 'w') as f:
                json.dump(config_data, f, indent=2, default=str)
            
            print(f"✅ Configuration exported to: {output_file}")
            print(f"📊 Exported {config_data['total_sheets']} sheets")
            print(f"   ({config_data['enabled_sheets']} enabled, "
                  f"{config_data['total_sheets'] - config_data['enabled_sheets']} disabled)")
            
        except Exception as e:
            print(f"❌ Error exporting configuration: {e}")
    
    def optimize_order(self, args) -> None:
        """Optimize sheet order within categories."""
        print("🔧 Optimizing sheet order...")
        
        try:
            success = self.manager.optimize_sheet_order()
            
            if success:
                print("✅ Sheet order optimized successfully")
                print("   Gaps removed, sequential ordering applied within categories")
            else:
                print("❌ Failed to optimize sheet order")
                
        except Exception as e:
            print(f"❌ Error optimizing order: {e}")
    
    def show_dependencies(self, args) -> None:
        """Show dependencies for a specific sheet."""
        print(f"🔍 Analyzing dependencies for: {args.sheet_name}")
        
        try:
            deps = self.manager.get_sheet_dependencies(args.sheet_name)
            
            if 'error' in deps:
                print(f"❌ {deps['error']}")
                return
            
            print("\n" + "=" * 60)
            print("SHEET DEPENDENCY ANALYSIS")
            print("=" * 60)
            
            print(f"\n📋 BASIC INFO")
            print(f"  Sheet Name: {deps['sheet_name']}")
            print(f"  Pipeline: {deps['pipeline']}")
            print(f"  Module: {deps['module']}")
            print(f"  Category: {deps['category']}")
            print(f"  Enabled: {'✅ Yes' if deps['enabled'] else '❌ No'}")
            print(f"  Order: {deps['order']}")
            
            print(f"\n🔗 DEPENDENCIES")
            print(f"  Module Enabled: {'✅ Yes' if deps.get('module_enabled', True) else '❌ No'}")
            print(f"  Pipeline Exists: {'✅ Yes' if deps.get('pipeline_exists', True) else '❌ No'}")
            
            if deps.get('module_description'):
                print(f"  Module Description: {deps['module_description']}")
            
        except Exception as e:
            print(f"❌ Error analyzing dependencies: {e}")


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Sheet Management CLI for AR Data Analysis",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s list                                    # List all sheets
  %(prog)s list --verbose                          # List with descriptions
  %(prog)s enable "Daily Counts" "Weekly Counts"   # Enable specific sheets
  %(prog)s disable "Detailed Analysis"             # Disable a sheet
  %(prog)s reorder "Daily Counts:5,Weekly Counts:10"  # Reorder sheets
  %(prog)s validate                                # Validate configuration
  %(prog)s status                                  # Show system status
  %(prog)s export --output config.json            # Export configuration
  %(prog)s optimize                                # Optimize sheet order
  %(prog)s deps "Daily Counts"                     # Show dependencies
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # List command
    list_parser = subparsers.add_parser('list', help='List all sheets')
    list_parser.add_argument('--verbose', '-v', action='store_true', 
                           help='Show detailed information')
    
    # Enable command
    enable_parser = subparsers.add_parser('enable', help='Enable sheets')
    enable_parser.add_argument('sheets', nargs='+', help='Sheet names to enable')
    enable_parser.add_argument('--no-validate', action='store_true',
                             help='Skip validation before enabling')
    
    # Disable command
    disable_parser = subparsers.add_parser('disable', help='Disable sheets')
    disable_parser.add_argument('sheets', nargs='+', help='Sheet names to disable')
    disable_parser.add_argument('--no-validate', action='store_true',
                               help='Skip validation before disabling')
    
    # Reorder command
    reorder_parser = subparsers.add_parser('reorder', help='Reorder sheets')
    reorder_parser.add_argument('reorder_spec', 
                               help='Reorder specification (format: "sheet:order,sheet:order")')
    
    # Validate command
    validate_parser = subparsers.add_parser('validate', help='Validate configuration')
    
    # Status command
    status_parser = subparsers.add_parser('status', help='Show system status')
    
    # Export command
    export_parser = subparsers.add_parser('export', help='Export configuration')
    export_parser.add_argument('--output', '-o', help='Output file path')
    export_parser.add_argument('--include-disabled', action='store_true',
                              help='Include disabled sheets in export')
    
    # Optimize command
    optimize_parser = subparsers.add_parser('optimize', help='Optimize sheet order')
    
    # Dependencies command
    deps_parser = subparsers.add_parser('deps', help='Show sheet dependencies')
    deps_parser.add_argument('sheet_name', help='Name of the sheet to analyze')
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return
    
    # Initialize CLI
    cli = SheetManagementCLI()
    
    # Route to appropriate command
    command_map = {
        'list': cli.list_sheets,
        'enable': cli.enable_sheets,
        'disable': cli.disable_sheets,
        'reorder': cli.reorder_sheets,
        'validate': cli.validate_sheets,
        'status': cli.show_status,
        'export': cli.export_config,
        'optimize': cli.optimize_order,
        'deps': cli.show_dependencies
    }
    
    if args.command in command_map:
        command_map[args.command](args)
    else:
        print(f"❌ Unknown command: {args.command}")
        parser.print_help()


if __name__ == "__main__":
    main()
