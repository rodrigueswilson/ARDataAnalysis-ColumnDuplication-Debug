#!/usr/bin/env python3
r"""
AR Data Analysis - Modular Report Generator
==========================================

This is the new lightweight orchestrator for generating comprehensive Excel reports
from MongoDB data. It uses the modular report_generator package for improved
maintainability and scalability.

The original generate_report.py (2,695 lines) has been refactored into:
- report_generator/core.py: Main ReportGenerator class and orchestration
- report_generator/sheet_creators.py: Individual sheet creation methods
- report_generator/formatters.py: Excel formatting and styling utilities
- report_generator/dashboard.py: Dashboard and summary generation
- report_generator/utils.py: Utility functions and helpers

Usage:
    python generate_report_modular.py --db_path D:\ARDataAnalysis\db
"""

import os
import argparse
import atexit

from report_generator import (
    ReportGenerator,
    backup_py_and_md_files,
    get_db_connection
)

def main():
    """
    Main function to drive the modular report generation.
    
    This lightweight orchestrator handles command-line arguments, database
    connections, and delegates the actual report generation to the modular
    ReportGenerator class.
    """
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description='Generates a formatted Excel report from the ARDataAnalysis MongoDB.',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument(
        '--db_path',
        required=False,
        default='db',
        help='The directory where the MongoDB data files are stored.\nDefaults to "db" in the current directory.'
    )
    parser.add_argument(
        '--output_dir',
        required=False,
        help='Directory to save the report. Defaults to the script\'s location.'
    )
    
    args = parser.parse_args()
    
    # Determine output directory
    output_dir = args.output_dir or os.path.dirname(os.path.abspath(__file__))
    
    print("=" * 60)
    print("AR DATA ANALYSIS - MODULAR REPORT GENERATOR")
    print("=" * 60)
    print(f"Database path: {args.db_path}")
    print(f"Output directory: {output_dir}")
    print("=" * 60)
    
    # Perform automatic backup
    print("[BACKUP] Creating automatic backup...")
    backup_py_and_md_files()
    
    # Get database connection
    print("[DATABASE] Connecting to MongoDB...")
    db = get_db_connection()
    if db is None:
        print("[ERROR] Failed to connect to database")
        return 1

    try:
        # Create and run the modular report generator
        print("[GENERATOR] Initializing modular report generator...")
        # Use current directory as root_dir, not the db_path
        root_dir = os.path.dirname(os.path.abspath(__file__))
        reporter = ReportGenerator(db, root_dir, output_dir)
        
        print("[GENERATOR] Starting report generation...")
        reporter.generate_report()
        
        print("[SUCCESS] Report generation completed successfully!")
        return 0
        
    except Exception as e:
        print(f"[ERROR] Report generation failed: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == '__main__':
    exit_code = main()
    exit(exit_code)
