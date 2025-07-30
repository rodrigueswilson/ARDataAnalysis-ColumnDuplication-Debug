#!/usr/bin/env python
"""
AR Data Analysis - Directory Decluttering Script

This script automatically organizes files in the ARDataAnalysis directory according to the
project's directory organization guidelines. It moves non-essential files from the root
directory to appropriate subdirectories under 'archive', while keeping essential files in place.
"""

import os
import shutil
import re
from datetime import datetime
import json

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

# Define essential files that should remain in the root directory
ESSENTIAL_FILES = [
    # Core Python Scripts
    'ar_utils.py',
    'populate_db.py',
    'generate_report.py',
    'tag_jpg_files.py',
    'tag_mp3_files.py',
    'pipelines.py',
    # Config Files
    'report_config.json',
    'requirements.txt',
    # Essential Documentation
    'README.md',
    'DIRECTORY_ORGANIZATION.md',
]

# Essential directories that should remain in the root
ESSENTIAL_DIRS = [
    'db',
    '21_22 Audio',
    '21_22 Photos',
    '22_23 Audio',
    '22_23 Photos',
    'TestData',
    'charts',
    'charts_output',
    'venv',
    'archive',
    'pipelines',
    'report_generator',
]

# Create necessary archive subdirectories if they don't exist
def create_archive_subdirs():
    """Create the necessary archive subdirectories if they don't exist."""
    subdirs = [
        'recovery',
        'reports',
        'temp_files',
        'exports',
        'testing',
        'documentation',
    ]
    
    for subdir in subdirs:
        full_path = os.path.join(ROOT_DIR, 'archive', subdir)
        if not os.path.exists(full_path):
            os.makedirs(full_path)
            print(f"Created directory: {full_path}")

# Categorize a file and determine where it should be moved
def categorize_file(filename):
    """
    Determine the appropriate archive subdirectory for a non-essential file.
    
    Args:
        filename: Name of the file to categorize
        
    Returns:
        The archive subdirectory where the file should be moved, or None if it should stay
    """
    # If it's an essential file or directory, keep it in root
    if filename in ESSENTIAL_FILES or filename in ESSENTIAL_DIRS:
        return None
        
    # Skip the script itself
    if filename == os.path.basename(__file__):
        return None
    
    # Categorize by file type and name patterns
    
    # Generated reports
    if re.match(r'AR_Analysis_Report_\d+.*\.xlsx$', filename):
        return 'reports'
        
    # Log and output files
    if filename.endswith(('.log', '.txt')) and filename != 'requirements.txt':
        return 'temp_files'
        
    # Test, debug, verification scripts
    if filename.startswith(('test_', 'debug_', 'verify_', 'check_', 'analyze_', 'inspect_')):
        return 'recovery'
        
    # Generated images
    if filename.endswith(('.png', '.jpg', '.jpeg')):
        return 'exports'
        
    # Support modules not in ESSENTIAL_FILES
    if filename.endswith('.py') and filename not in ESSENTIAL_FILES:
        return 'recovery'
        
    # Development config files
    if filename.startswith('.') or filename == '__pycache__':
        return 'temp_files'
        
    # Additional documentation
    if filename.endswith('.md') and filename not in ESSENTIAL_FILES:
        return 'documentation'
        
    # Database exports
    if filename.endswith('.json') and not filename == 'report_config.json':
        return 'exports'
        
    # Excel files not reports
    if filename.endswith('.xlsx') and not re.match(r'AR_Analysis_Report_\d+.*\.xlsx$', filename):
        return 'testing'
        
    # YAML files
    if filename.endswith('.yaml') or filename.endswith('.yml'):
        return 'temp_files'
        
    # Default case for any other files
    return 'recovery'

# Move a file to its destination archive subdirectory
def move_file(filepath, dest_subdir):
    """
    Move a file to the appropriate archive subdirectory.
    
    Args:
        filepath: Full path to the file to move
        dest_subdir: Destination subdirectory under archive/
    """
    filename = os.path.basename(filepath)
    dest_dir = os.path.join(ROOT_DIR, 'archive', dest_subdir)
    dest_path = os.path.join(dest_dir, filename)
    
    # Handle file name conflicts by adding a timestamp
    if os.path.exists(dest_path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(filename)
        dest_path = os.path.join(dest_dir, f"{base}_{timestamp}{ext}")
    
    try:
        if os.path.isdir(filepath):
            if os.path.exists(dest_path):
                # For directories, merge contents
                for item in os.listdir(filepath):
                    src_item = os.path.join(filepath, item)
                    dest_item = os.path.join(dest_path, item)
                    if os.path.isdir(src_item):
                        # Recursively copy directory
                        if not os.path.exists(dest_item):
                            shutil.copytree(src_item, dest_item)
                        else:
                            # Merge directories
                            for subitem in os.listdir(src_item):
                                shutil.move(os.path.join(src_item, subitem), dest_item)
                    else:
                        # Move file
                        if not os.path.exists(dest_item):
                            shutil.move(src_item, dest_item)
                        else:
                            # Handle conflicts for files in subdirectories
                            base, ext = os.path.splitext(item)
                            dest_item = os.path.join(dest_path, f"{base}_{timestamp}{ext}")
                            shutil.move(src_item, dest_item)
                shutil.rmtree(filepath)
            else:
                # Simple directory move
                shutil.move(filepath, dest_path)
        else:
            # Simple file move
            shutil.move(filepath, dest_path)
        return dest_path
    except Exception as e:
        print(f"Error moving {filepath}: {e}")
        return None

# Main declutter function
def declutter_directory():
    """
    Main function to declutter the directory by moving files to appropriate archive locations.
    """
    print("Starting AR Data Analysis directory decluttering...")
    create_archive_subdirs()
    
    # Collect all items in the root directory
    items = os.listdir(ROOT_DIR)
    
    # Track movement operations for reporting
    moved_files = {
        'reports': [],
        'recovery': [],
        'temp_files': [],
        'exports': [],
        'testing': [],
        'documentation': [],
    }
    
    kept_files = []
    errors = []
    
    # Process each item
    for item in items:
        if item in ESSENTIAL_FILES or item in ESSENTIAL_DIRS:
            kept_files.append(item)
            continue
            
        if item == os.path.basename(__file__):
            kept_files.append(item)
            continue
            
        filepath = os.path.join(ROOT_DIR, item)
        dest_subdir = categorize_file(item)
        
        if dest_subdir:
            try:
                dest_path = move_file(filepath, dest_subdir)
                if dest_path:
                    moved_files[dest_subdir].append(item)
            except Exception as e:
                errors.append(f"{item}: {str(e)}")
        else:
            kept_files.append(item)
    
    # Generate report
    print("\n=== AR Data Analysis Directory Decluttering Report ===\n")
    
    print(f"Essential files kept in root directory ({len(kept_files)}):")
    for file in sorted(kept_files):
        print(f"  - {file}")
    
    print("\nFiles moved to archive subdirectories:")
    for subdir, files in moved_files.items():
        if files:
            print(f"\n  archive/{subdir}/ ({len(files)}):")
            for file in sorted(files):
                print(f"    - {file}")
    
    if errors:
        print("\nErrors encountered:")
        for error in errors:
            print(f"  - {error}")
    
    print("\nDecluttering complete!")
    
    # Create a log file with the results
    log_data = {
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'kept_files': kept_files,
        'moved_files': moved_files,
        'errors': errors
    }
    
    log_file = os.path.join(ROOT_DIR, 'archive', 'declutter_log.json')
    with open(log_file, 'w') as f:
        json.dump(log_data, f, indent=2)
    
    print(f"\nLog saved to: {log_file}")
    
    # Also generate a markdown recovery guide
    recovery_guide = os.path.join(ROOT_DIR, 'DECLUTTER_RECOVERY_GUIDE.md')
    generate_recovery_guide(moved_files, recovery_guide)
    
    print(f"Recovery guide updated: {recovery_guide}")

def generate_recovery_guide(moved_files, guide_path):
    """
    Generate a recovery guide markdown file that explains how to restore moved files if needed.
    
    Args:
        moved_files: Dictionary of moved files by category
        guide_path: Path to write the recovery guide
    """
    with open(guide_path, 'w') as f:
        f.write("# AR Data Analysis - Declutter Recovery Guide\n\n")
        f.write("## Overview\n\n")
        f.write("This document tracks files that have been moved from the root directory ")
        f.write("to organized subdirectories under `archive/`. If you need to restore any files ")
        f.write("for debugging or development purposes, refer to this guide.\n\n")
        
        f.write("## Last Declutter Operation\n\n")
        f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        f.write("## File Locations\n\n")
        
        total_moved = sum(len(files) for files in moved_files.values())
        f.write(f"Total files moved: {total_moved}\n\n")
        
        for subdir, files in moved_files.items():
            if files:
                f.write(f"### Files in `archive/{subdir}/`\n\n")
                for file in sorted(files):
                    f.write(f"- `{file}`\n")
                f.write("\n")
        
        f.write("## How to Restore Files\n\n")
        f.write("To restore a file to the root directory, you can use the following command:\n\n")
        f.write("```bash\n")
        f.write("# Example: Restore a specific file\n")
        f.write("cp archive/recovery/debug_script.py ./\n\n")
        f.write("# Example: Restore all test files\n")
        f.write("cp archive/testing/test_*.py ./\n")
        f.write("```\n\n")
        
        f.write("## Notes\n\n")
        f.write("- All essential files remain in the root directory for normal operation\n")
        f.write("- This decluttering follows the organization defined in DIRECTORY_ORGANIZATION.md\n")
        f.write("- Each decluttering operation is logged in `archive/declutter_log.json`\n")
        
        # Add warning about production environment
        f.write("\n## Important\n\n")
        f.write("⚠️ **The root directory should be kept clean for production use.** Only restore files ")
        f.write("temporarily for debugging or development, then move them back to the archive when finished.\n")

# Run the decluttering process
if __name__ == "__main__":
    declutter_directory()
