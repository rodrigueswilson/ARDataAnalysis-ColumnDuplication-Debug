#!/usr/bin/env python
"""
AR Data Analysis - Enhanced Directory Decluttering Script

This script organizes files in the ARDataAnalysis directory by moving non-essential files
to appropriate archive subdirectories while keeping core codebase files in the root.

Enhanced version based on current directory analysis.
"""

import os
import shutil
import re
from datetime import datetime
import json
import zipfile
import ast
import sys

# Get the directory where this script is located (should be the project root)
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

# Backup configuration
BACKUP_DIR = os.path.join(ROOT_DIR, 'archive', 'backups')
MAX_BACKUPS = 5  # Keep last 5 backups

# Define essential files that should remain in the root directory
ESSENTIAL_FILES = [
    # Core Python Scripts (main codebase)
    'ar_utils.py',
    'populate_db.py', 
    'generate_report.py',
    'tag_jpg_files.py',
    'tag_mp3_files.py',
    'acf_pacf_charts.py',
    'dashboard_generator.py',
    'db_utils.py',
    'chart_config_helper.py',
    'column_cleanup_utils.py',
    
    # Configuration Files
    'report_config.json',
    'config.yaml',
    'requirements.txt',
    '.gitignore',
    
    # Essential Documentation
    'README.md',
]

# Essential directories that should remain in the root
ESSENTIAL_DIRS = [
    '.git',
    '__pycache__',
    'db',
    'pipelines',
    'report_generator', 
    'utils',
    'venv',
    'archive',
    # Data directories (if they exist)
    '21_22 Audio',
    '21_22 Photos', 
    '22_23 Audio',
    '22_23 Photos',
    'TestData',
    'charts',
    'charts_output',
]

def detect_codebase_files():
    """
    Intelligently detect files that belong to the codebase by analyzing:
    - Python files with imports from the project
    - Configuration files referenced by the code
    - Files that are imported or used by core scripts
    
    Returns:
        set: Set of files that are part of the codebase
    """
    codebase_files = set(ESSENTIAL_FILES)
    
    # Analyze Python files to find project dependencies
    for item in os.listdir(ROOT_DIR):
        if item.endswith('.py') and os.path.isfile(os.path.join(ROOT_DIR, item)):
            try:
                with open(os.path.join(ROOT_DIR, item), 'r', encoding='utf-8') as f:
                    content = f.read()
                    
                # Parse the AST to find imports
                try:
                    tree = ast.parse(content)
                    for node in ast.walk(tree):
                        if isinstance(node, ast.Import):
                            for alias in node.names:
                                # Check if importing local modules
                                if alias.name in ['ar_utils', 'db_utils', 'dashboard_generator', 
                                                'acf_pacf_charts', 'chart_config_helper', 
                                                'column_cleanup_utils']:
                                    codebase_files.add(item)
                        elif isinstance(node, ast.ImportFrom):
                            if node.module and node.module.startswith(('report_generator', 'pipelines', 'utils')):
                                codebase_files.add(item)
                except SyntaxError:
                    # If we can't parse, check for import patterns with regex
                    import_patterns = [
                        r'from\s+(report_generator|pipelines|utils)',
                        r'import\s+(ar_utils|db_utils|dashboard_generator|acf_pacf_charts)',
                        r'from\s+ar_utils\s+import',
                        r'from\s+db_utils\s+import'
                    ]
                    
                    for pattern in import_patterns:
                        if re.search(pattern, content):
                            codebase_files.add(item)
                            break
                            
                # Check for configuration file references
                config_patterns = [
                    r'report_config\.json',
                    r'config\.yaml',
                    r'requirements\.txt'
                ]
                
                for pattern in config_patterns:
                    if re.search(pattern, content):
                        # This file uses configuration, likely part of codebase
                        codebase_files.add(item)
                        
            except Exception as e:
                print(f"Warning: Could not analyze {item}: {str(e)}")
                
    # Add files that are commonly part of Python projects
    common_codebase_patterns = [
        r'.*\.py$',  # All Python files in root (will be filtered later)
        r'.*\.json$',  # Configuration files
        r'.*\.yaml$',  # Configuration files
        r'.*\.yml$',   # Configuration files
        r'requirements.*\.txt$',  # Requirements files
        r'\.gitignore$',  # Git files
        r'README\.md$',   # Documentation
        r'setup\.py$',    # Setup files
        r'__init__\.py$'  # Python package files
    ]
    
    for item in os.listdir(ROOT_DIR):
        if os.path.isfile(os.path.join(ROOT_DIR, item)):
            for pattern in common_codebase_patterns:
                if re.match(pattern, item, re.IGNORECASE):
                    # Exclude obvious non-codebase files
                    if not any([
                        item.startswith('debug_'),
                        item.startswith('test_'),
                        item.startswith('trace_'),
                        item.startswith('diagnose_'),
                        item.startswith('analyze_'),
                        item.startswith('fix_'),
                        item.startswith('verify_'),
                        item.startswith('check_'),
                        item.startswith('MARKER_'),
                        'temp' in item.lower(),
                        'backup' in item.lower(),
                        re.match(r'AR_Analysis_Report_.*\.xlsx$', item)
                    ]):
                        codebase_files.add(item)
                    break
    
    return codebase_files

def create_backup():
    """
    Create a timestamped backup of all codebase files before decluttering.
    
    Returns:
        str: Path to the created backup file, or None if backup failed
    """
    try:
        # Ensure backup directory exists
        if not os.path.exists(BACKUP_DIR):
            os.makedirs(BACKUP_DIR)
            print(f"Created backup directory: {BACKUP_DIR}")
        
        # Detect codebase files
        codebase_files = detect_codebase_files()
        
        # Create timestamped backup filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"codebase_backup_{timestamp}.zip"
        backup_path = os.path.join(BACKUP_DIR, backup_filename)
        
        print(f"\nCreating backup of {len(codebase_files)} codebase files...")
        
        # Create zip backup
        with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in sorted(codebase_files):
                file_path = os.path.join(ROOT_DIR, file)
                if os.path.isfile(file_path):
                    zipf.write(file_path, file)
                    print(f"  ✓ Backed up: {file}")
            
            # Also backup essential directories
            for dir_name in ESSENTIAL_DIRS:
                dir_path = os.path.join(ROOT_DIR, dir_name)
                if os.path.isdir(dir_path) and dir_name not in ['.git', '__pycache__', 'venv', 'archive']:
                    for root, dirs, files in os.walk(dir_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            # Get relative path for zip
                            rel_path = os.path.relpath(file_path, ROOT_DIR)
                            zipf.write(file_path, rel_path)
        
        print(f"✅ Backup created: {backup_filename}")
        
        # Clean up old backups (keep only MAX_BACKUPS)
        cleanup_old_backups()
        
        return backup_path
        
    except Exception as e:
        print(f"❌ Backup failed: {str(e)}")
        return None

def cleanup_old_backups():
    """
    Remove old backup files, keeping only the most recent MAX_BACKUPS.
    """
    try:
        if not os.path.exists(BACKUP_DIR):
            return
            
        # Get all backup files
        backup_files = []
        for file in os.listdir(BACKUP_DIR):
            if file.startswith('codebase_backup_') and file.endswith('.zip'):
                file_path = os.path.join(BACKUP_DIR, file)
                backup_files.append((file_path, os.path.getmtime(file_path)))
        
        # Sort by modification time (newest first)
        backup_files.sort(key=lambda x: x[1], reverse=True)
        
        # Remove old backups
        if len(backup_files) > MAX_BACKUPS:
            for file_path, _ in backup_files[MAX_BACKUPS:]:
                os.remove(file_path)
                print(f"  🗑️ Removed old backup: {os.path.basename(file_path)}")
                
    except Exception as e:
        print(f"Warning: Could not cleanup old backups: {str(e)}")

def create_archive_subdirs():
    """Create the necessary archive subdirectories if they don't exist."""
    subdirs = [
        'debugging',      # Debug and diagnostic scripts
        'testing',        # Test scripts and validation
        'reports',        # Generated Excel reports
        'temporary',      # Temporary files, markers, logs
        'documentation',  # Non-essential documentation
        'legacy',         # Legacy/backup files
        'outputs',        # Generated outputs and exports
    ]
    
    archive_dir = os.path.join(ROOT_DIR, 'archive')
    if not os.path.exists(archive_dir):
        os.makedirs(archive_dir)
        print(f"Created archive directory: {archive_dir}")
    
    for subdir in subdirs:
        full_path = os.path.join(archive_dir, subdir)
        if not os.path.exists(full_path):
            os.makedirs(full_path)
            print(f"Created directory: {full_path}")

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
        
    # Skip this script itself
    if filename == os.path.basename(__file__):
        return None
    
    # Categorize by file type and name patterns
    
    # Generated Excel reports
    if re.match(r'AR_Analysis_Report_\d+.*\.xlsx$', filename):
        return 'reports'
    
    # Test files and ACF/PACF test reports
    if filename.startswith('ACF_PACF_Test_Report_') or filename.endswith('_test.txt'):
        return 'testing'
        
    # Debug scripts
    if filename.startswith(('debug_', 'diagnose_', 'trace_', 'inspect_')):
        return 'debugging'
        
    # Test and validation scripts  
    if filename.startswith(('test_', 'validate_', 'verify_', 'check_')):
        return 'testing'
        
    # Analysis and audit scripts
    if filename.startswith(('analyze_', 'audit_', 'simple_', 'final_')):
        return 'debugging'
        
    # Fix and standardization scripts
    if filename.startswith(('fix_', 'standardize_')):
        return 'debugging'
        
    # Temporary files and markers
    if filename.startswith('MARKER_') or filename.endswith(('.txt', '.log')) and not filename == 'requirements.txt':
        return 'temporary'
        
    # Documentation files (except essential ones)
    if filename.endswith('.md') and filename not in ESSENTIAL_FILES:
        return 'documentation'
        
    # Legacy/backup files
    if 'LEGACY' in filename or 'ORIGINAL' in filename:
        return 'legacy'
        
    # Database exports and JSON files (except config)
    if filename.endswith('.json') and filename not in ESSENTIAL_FILES:
        return 'outputs'
        
    # Python scripts not in essential list
    if filename.endswith('.py') and filename not in ESSENTIAL_FILES:
        return 'debugging'
        
    # Any other Excel files
    if filename.endswith(('.xlsx', '.xls')):
        return 'testing'
        
    # CLI and management tools
    if 'cli' in filename.lower() or 'management' in filename.lower():
        return 'debugging'
        
    # Default case for any other files
    return 'temporary'

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
    
    # Handle filename conflicts
    counter = 1
    original_dest_path = dest_path
    while os.path.exists(dest_path):
        name, ext = os.path.splitext(filename)
        dest_path = os.path.join(dest_dir, f"{name}_{counter}{ext}")
        counter += 1
    
    try:
        shutil.move(filepath, dest_path)
        print(f"Moved: {filename} -> archive/{dest_subdir}/")
        return dest_path
    except Exception as e:
        print(f"Error moving {filename}: {str(e)}")
        return None

def declutter_directory():
    """
    Main function to declutter the directory by moving files to appropriate archive locations.
    """
    print("=== AR Data Analysis Directory Decluttering ===")
    print(f"Working directory: {ROOT_DIR}\n")
    
    # Step 1: Create backup of codebase files
    print("📦 STEP 1: Creating backup of codebase files...")
    backup_path = create_backup()
    if not backup_path:
        print("❌ Backup failed! Decluttering aborted for safety.")
        return
    
    # Step 2: Detect codebase files for enhanced protection
    print("\n🔍 STEP 2: Analyzing codebase structure...")
    detected_codebase = detect_codebase_files()
    print(f"Detected {len(detected_codebase)} codebase files:")
    for file in sorted(detected_codebase):
        print(f"  🔒 {file}")
    
    # Step 3: Create archive subdirectories
    print("\n📁 STEP 3: Setting up archive structure...")
    create_archive_subdirs()
    
    # Get all items in the root directory
    items = os.listdir(ROOT_DIR)
    
    # Track results
    moved_files = {
        'debugging': [],
        'testing': [],
        'reports': [],
        'temporary': [],
        'documentation': [],
        'legacy': [],
        'outputs': [],
    }
    kept_files = []
    errors = []
    
    print("Processing files...\n")
    
    # Process each item
    for item in items:
        item_path = os.path.join(ROOT_DIR, item)
        
        # Skip directories (they're handled separately)
        if os.path.isdir(item_path):
            kept_files.append(item + "/")
            continue
            
        # Categorize the file
        dest_subdir = categorize_file(item)
        
        # Double-check: never move detected codebase files
        if item in detected_codebase:
            kept_files.append(item + " (codebase)")
            if dest_subdir:
                print(f"  🔒 Protected codebase file: {item} (would have been moved to {dest_subdir})")
        elif dest_subdir:
            try:
                dest_path = move_file(item_path, dest_subdir)
                if dest_path:
                    moved_files[dest_subdir].append(item)
            except Exception as e:
                errors.append(f"{item}: {str(e)}")
        else:
            kept_files.append(item)
    
    # Generate report
    print("\n" + "="*60)
    print("DECLUTTERING COMPLETE")
    print("="*60)
    
    print(f"\nEssential files kept in root directory ({len(kept_files)}):")
    for file in sorted(kept_files):
        print(f"  ✓ {file}")
    
    print(f"\nFiles moved to archive subdirectories:")
    total_moved = 0
    for subdir, files in moved_files.items():
        if files:
            total_moved += len(files)
            print(f"\n  📁 archive/{subdir}/ ({len(files)} files):")
            for file in sorted(files):
                print(f"    → {file}")
    
    if errors:
        print(f"\n❌ Errors encountered ({len(errors)}):")
        for error in errors:
            print(f"    • {error}")
    
    print(f"\n📊 Summary:")
    print(f"  • Files kept in root: {len(kept_files)}")
    print(f"  • Files moved to archive: {total_moved}")
    print(f"  • Errors: {len(errors)}")
    print(f"  • Backup created: {os.path.basename(backup_path) if backup_path else 'Failed'}")
    
    # Create a detailed log file
    log_data = {
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'backup_path': backup_path,
        'detected_codebase_files': sorted(list(detected_codebase)),
        'summary': {
            'kept_files': len(kept_files),
            'moved_files': total_moved,
            'errors': len(errors),
            'codebase_files_detected': len(detected_codebase)
        },
        'kept_files': sorted(kept_files),
        'moved_files': moved_files,
        'errors': errors
    }
    
    log_file = os.path.join(ROOT_DIR, 'archive', 'declutter_log.json')
    try:
        with open(log_file, 'w') as f:
            json.dump(log_data, f, indent=2)
        print(f"\n📝 Detailed log saved to: archive/declutter_log.json")
    except Exception as e:
        print(f"\n❌ Could not save log file: {str(e)}")
    
    # Generate recovery guide
    try:
        recovery_guide = os.path.join(ROOT_DIR, 'archive', 'RECOVERY_GUIDE.md')
        generate_recovery_guide(moved_files, recovery_guide)
        print(f"📖 Recovery guide created: archive/RECOVERY_GUIDE.md")
    except Exception as e:
        print(f"❌ Could not create recovery guide: {str(e)}")
    
    print(f"\n✅ Decluttering complete! Root directory is now organized.")

def generate_recovery_guide(moved_files, guide_path):
    """
    Generate a recovery guide markdown file that explains how to restore moved files if needed.
    
    Args:
        moved_files: Dictionary of moved files by category
        guide_path: Path to write the recovery guide
    """
    with open(guide_path, 'w') as f:
        f.write("# AR Data Analysis - File Recovery Guide\n\n")
        f.write("This guide helps you restore files that were moved during directory decluttering.\n\n")
        
        f.write("## Decluttering Summary\n\n")
        f.write(f"**Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        total_moved = sum(len(files) for files in moved_files.values())
        f.write(f"**Total files moved:** {total_moved}\n\n")
        
        f.write("## Archive Organization\n\n")
        f.write("Files have been organized into the following categories:\n\n")
        
        category_descriptions = {
            'debugging': 'Debug scripts, diagnostic tools, and analysis utilities',
            'testing': 'Test scripts, validation tools, and test reports', 
            'reports': 'Generated Excel reports and analysis outputs',
            'temporary': 'Temporary files, markers, logs, and debug outputs',
            'documentation': 'Non-essential documentation and guides',
            'legacy': 'Legacy files, backups, and original versions',
            'outputs': 'Generated data exports and output files',
        }
        
        for subdir, files in moved_files.items():
            if files:
                f.write(f"### 📁 `archive/{subdir}/` ({len(files)} files)\n\n")
                f.write(f"**Purpose:** {category_descriptions.get(subdir, 'Miscellaneous files')}\n\n")
                f.write("**Files:**\n")
                for file in sorted(files):
                    f.write(f"- `{file}`\n")
                f.write("\n")
        
        f.write("## How to Restore Files\n\n")
        f.write("### Restore a Single File\n")
        f.write("```bash\n")
        f.write("# Copy a file back to the root directory\n")
        f.write("cp archive/debugging/debug_script.py ./\n")
        f.write("```\n\n")
        
        f.write("### Restore Multiple Files\n")
        f.write("```bash\n")
        f.write("# Restore all test files\n")
        f.write("cp archive/testing/test_*.py ./\n\n")
        f.write("# Restore all debugging scripts\n")
        f.write("cp archive/debugging/*.py ./\n")
        f.write("```\n\n")
        
        f.write("### Restore Entire Category\n")
        f.write("```bash\n")
        f.write("# Restore all files from a category\n")
        f.write("cp archive/debugging/* ./\n")
        f.write("```\n\n")
        
        f.write("## Important Notes\n\n")
        f.write("- ✅ **All essential codebase files remain in the root directory**\n")
        f.write("- 🔄 **Only restore files temporarily for debugging/development**\n")
        f.write("- 🧹 **Move files back to archive when finished to keep root clean**\n")
        f.write("- 📝 **Check `archive/declutter_log.json` for detailed operation log**\n\n")
        
        f.write("## Essential Files (Kept in Root)\n\n")
        f.write("The following files remain in the root directory for normal operation:\n\n")
        for file in sorted(ESSENTIAL_FILES):
            f.write(f"- `{file}`\n")
        f.write("\n")
        
        f.write("## Essential Directories (Kept in Root)\n\n")
        for dir_name in sorted(ESSENTIAL_DIRS):
            f.write(f"- `{dir_name}/`\n")

if __name__ == "__main__":
    print("AR Data Analysis - Enhanced Directory Decluttering with Backup")
    print("=" * 60)
    print("\n🔍 This script will:")
    print("  1. Automatically detect codebase files")
    print("  2. Create a backup of all codebase files")
    print("  3. Organize non-essential files into archive subdirectories")
    print("  4. Protect detected codebase files from being moved")
    print(f"  5. Keep the last {MAX_BACKUPS} backups for safety")
    
    # Show what would be detected as codebase
    print("\n🔍 Analyzing current directory...")
    detected = detect_codebase_files()
    print(f"\nDetected {len(detected)} codebase files that will be protected:")
    for file in sorted(list(detected)[:10]):  # Show first 10
        print(f"  🔒 {file}")
    if len(detected) > 10:
        print(f"  ... and {len(detected) - 10} more")
    
    response = input("\n❓ Proceed with backup and decluttering? (y/N): ")
    
    if response.lower() in ['y', 'yes']:
        declutter_directory()
    else:
        print("Decluttering cancelled.")
