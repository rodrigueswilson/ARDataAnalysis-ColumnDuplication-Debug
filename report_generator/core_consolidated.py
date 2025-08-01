"""
Consolidated Core Report Generator
=================================

This is the updated core report generator that uses the consolidated sheet factory
to eliminate all duplicate sheet creation and enforce single-source-of-truth.

Key improvements:
1. Single factory for all sheet creation
2. Registry-based duplicate prevention
3. Comprehensive validation
4. Clean separation of concerns
5. Better error handling and logging
"""

import os
import datetime
import openpyxl
from pathlib import Path

from .consolidated_sheet_factory import ConsolidatedSheetFactory
from .sheet_validation import get_sheet_validator
from .formatters import ExcelFormatter


class ConsolidatedReportGenerator:
    """
    Consolidated report generator that eliminates duplicate sheet creation.
    
    This class replaces the original ReportGenerator with a cleaner architecture:
    - Uses ConsolidatedSheetFactory for all sheet creation
    - Enforces registry-based duplicate prevention
    - Provides comprehensive validation
    - Maintains backward compatibility
    """
    
    def __init__(self, db, root_dir, output_dir=None):
        """
        Initialize the consolidated report generator.
        
        Args:
            db: MongoDB database connection
            root_dir (str): Root directory for the project
            output_dir (str, optional): Output directory for reports. Defaults to root_dir.
        """
        self.db = db
        self.root_dir = root_dir
        self.output_dir = output_dir or root_dir
        self.workbook = openpyxl.Workbook()
        self.workbook.remove(self.workbook.active)  # Remove default sheet
        
        # Initialize consolidated components
        self.formatter = ExcelFormatter()
        self.sheet_factory = ConsolidatedSheetFactory(self.db, self.formatter)
        self.validator = get_sheet_validator()
        
        print("[CONSOLIDATED] Report generator initialized with consolidated architecture")
    
    def generate_report(self):
        """
        Generates the full, multi-sheet Excel report using the consolidated architecture.
        
        This method eliminates all duplicate sheet creation by using the ConsolidatedSheetFactory
        as the single source of truth for all sheet creation operations.
        """
        print("=" * 60)
        print("AR DATA ANALYSIS - CONSOLIDATED REPORT GENERATION")
        print("=" * 60)
        
        # Generate timestamp for output file
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"AR_Analysis_Report_Consolidated_{timestamp}.xlsx"
        output_path = os.path.join(self.output_dir, output_filename)
        
        print(f"[INFO] Starting consolidated report generation...")
        print(f"[INFO] Output file: {output_path}")
        print(f"[INFO] Using ConsolidatedSheetFactory for all sheet creation")
        
        try:
            # Phase 1: Pre-generation validation
            print("\n[PHASE 1] Pre-generation validation...")
            config_path = Path(self.root_dir) / "report_config.json"
            config_valid, config_issues = self.validator.validate_configuration(str(config_path))
            
            if config_issues:
                print("[VALIDATION] Configuration issues found:")
                self.validator.print_validation_report(config_issues)
            
            if not config_valid:
                print("[ERROR] Configuration validation failed. Aborting report generation.")
                return False
            
            # Phase 2: Create all sheets using consolidated factory
            print("\n[PHASE 2] Creating all sheets via ConsolidatedSheetFactory...")
            creation_results = self.sheet_factory.create_all_sheets(self.workbook)
            
            # Phase 3: Post-creation validation
            print("\n[PHASE 3] Post-creation validation...")
            workbook_valid, workbook_issues = self.validator.validate_workbook_integrity(self.workbook)
            
            if workbook_issues:
                print("[VALIDATION] Workbook issues found:")
                self.validator.print_validation_report(workbook_issues)
            
            # Phase 4: Final validation and consistency check
            print("\n[PHASE 4] Final validation and consistency check...")
            validation_results = self.sheet_factory.validate_workbook(self.workbook)
            
            if not validation_results["is_consistent"]:
                print("[WARNING] Registry-workbook inconsistency detected:")
                if validation_results["untracked_sheets"]:
                    print(f"  Untracked sheets: {validation_results['untracked_sheets']}")
                if validation_results["missing_sheets"]:
                    print(f"  Missing sheets: {validation_results['missing_sheets']}")
            
            # Phase 5: Save the workbook
            print("\n[PHASE 5] Saving consolidated workbook...")
            self.workbook.save(output_path)
            
            # Phase 6: Generate final report
            print("\n[PHASE 6] Generating final report...")
            self._print_generation_summary(creation_results, output_path)
            
            return True
            
        except Exception as e:
            print(f"[ERROR] Consolidated report generation failed: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _print_generation_summary(self, creation_results, output_path):
        """
        Print a comprehensive summary of the report generation process.
        
        Args:
            creation_results: Dictionary of sheet creation results
            output_path: Path to the generated report file
        """
        print("\n" + "="*60)
        print("CONSOLIDATED REPORT GENERATION SUMMARY")
        print("="*60)
        
        # Overall statistics
        total_sheets = len(creation_results)
        successful_sheets = sum(1 for success in creation_results.values() if success)
        failed_sheets = total_sheets - successful_sheets
        
        print(f"Report File: {output_path}")
        print(f"Total Sheets Attempted: {total_sheets}")
        print(f"Successfully Created: {successful_sheets}")
        print(f"Failed to Create: {failed_sheets}")
        
        if failed_sheets > 0:
            print(f"\nFailed Sheets:")
            for sheet_name, success in creation_results.items():
                if not success:
                    print(f"  ❌ {sheet_name}")
        
        print(f"\nSuccessful Sheets:")
        for sheet_name, success in creation_results.items():
            if success:
                print(f"  ✅ {sheet_name}")
        
        # Registry summary
        registry_summary = self.sheet_factory.registry.get_creation_summary()
        print(f"\nRegistry Statistics:")
        print(f"  Total Creation Attempts: {registry_summary['total_attempts']}")
        print(f"  Successful Creations: {registry_summary['successful_creations']}")
        print(f"  Failed Creations: {registry_summary['failed_creations']}")
        
        # File size information
        try:
            file_size = os.path.getsize(output_path)
            file_size_mb = file_size / (1024 * 1024)
            print(f"\nGenerated File Size: {file_size_mb:.2f} MB")
        except:
            pass
        
        # Success indicator
        if failed_sheets == 0:
            print("\n🎉 REPORT GENERATION COMPLETED SUCCESSFULLY!")
            print("   All sheets created without duplications or conflicts.")
        else:
            print(f"\n⚠️  REPORT GENERATION COMPLETED WITH {failed_sheets} ISSUES")
            print("   Review failed sheets and check logs for details.")
        
        print("="*60)
    
    def validate_before_generation(self) -> bool:
        """
        Perform comprehensive validation before starting report generation.
        
        Returns:
            bool: True if validation passes, False otherwise
        """
        print("[VALIDATION] Performing pre-generation validation...")
        
        # Validate configuration
        config_path = Path(self.root_dir) / "report_config.json"
        config_valid, config_issues = self.validator.validate_configuration(str(config_path))
        
        # Validate database connection
        try:
            collection = self.db['media_records']
            record_count = collection.count_documents({})
            print(f"[VALIDATION] Database connection OK - {record_count} records found")
        except Exception as e:
            print(f"[VALIDATION] Database connection failed: {e}")
            return False
        
        # Print any validation issues
        if config_issues:
            self.validator.print_validation_report(config_issues)
        
        return config_valid
    
    def get_generation_status(self) -> dict:
        """
        Get the current status of report generation components.
        
        Returns:
            dict: Status information for debugging
        """
        return {
            "database_connected": self.db is not None,
            "workbook_sheets": len(self.workbook.sheetnames),
            "registry_summary": self.sheet_factory.registry.get_creation_summary(),
            "output_directory": self.output_dir,
            "root_directory": self.root_dir
        }


# Backward compatibility wrapper
class ReportGenerator(ConsolidatedReportGenerator):
    """
    Backward compatibility wrapper for the original ReportGenerator.
    
    This allows existing code to continue working while using the new
    consolidated architecture under the hood.
    """
    
    def __init__(self, db, root_dir, output_dir=None):
        """Initialize with backward compatibility."""
        super().__init__(db, root_dir, output_dir)
        print("[COMPATIBILITY] Using ConsolidatedReportGenerator for backward compatibility")
    
    # Legacy methods that may be called by existing code
    def _get_school_year(self, date):
        """Legacy method for backward compatibility."""
        if hasattr(date, 'month'):
            month = date.month
            year = date.year
        else:
            year, month = int(date[:4]), int(date[5:7])
        
        if month >= 8:
            return f"{year}-{str(year + 1)[2:]}"
        else:
            return f"{year - 1}-{str(year)[2:]}"
    
    def _seconds_to_hms(self, seconds):
        """Legacy method for backward compatibility."""
        if seconds is None or seconds == 0:
            return "00:00:00"
        
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    
    def _run_aggregation(self, pipeline, use_base_filter=True, collection_name='media_records'):
        """Legacy method - delegates to sheet factory."""
        return self.sheet_factory.base_creator._run_aggregation(pipeline, use_base_filter, collection_name)


# Factory function for easy instantiation
def create_report_generator(db, root_dir, output_dir=None, use_consolidated=True):
    """
    Factory function to create a report generator.
    
    Args:
        db: MongoDB database connection
        root_dir: Root directory for the project
        output_dir: Output directory for reports (optional)
        use_consolidated: Whether to use the new consolidated architecture (default: True)
        
    Returns:
        ReportGenerator instance (consolidated or legacy)
    """
    if use_consolidated:
        return ConsolidatedReportGenerator(db, root_dir, output_dir)
    else:
        # Import the original for fallback (if needed)
        from .core import ReportGenerator as OriginalReportGenerator
        return OriginalReportGenerator(db, root_dir, output_dir)
