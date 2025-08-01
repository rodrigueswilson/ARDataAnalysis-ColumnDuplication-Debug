#!/usr/bin/env python
"""
AR Data Analysis - Enhanced Directory Decluttering Script (v2)

This script organizes files in the ARDataAnalysis directory by moving non-essential files
to appropriate archive subdirectories while keeping core codebase files in the root.

Enhanced version with documentation system integration.
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

# Documentation configuration
DOCS_DIR = os.path.join(ROOT_DIR, 'docs')
DOCS_CONFIG_PATH = os.path.join(DOCS_DIR, 'docs_config.json')

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
    'docs',  # Added docs directory as essential
    # Data directories (if they exist)
    '21_22 Audio',
    '21_22 Photos', 
    '22_23 Audio',
    '22_23 Photos',
    'TestData',
    'charts',
    'charts_output',
]

def load_docs_config():
    """
    Load the documentation configuration file.
    
    Returns:
        dict: Documentation configuration or empty dict if not found
    """
    try:
        with open(DOCS_CONFIG_PATH, 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        print(f"Warning: Documentation configuration not found or invalid at {DOCS_CONFIG_PATH}")
        return {"categories": [], "documents": []}

def get_essential_doc_files():
    """
    Get list of essential documentation files from the docs configuration.
    
    Returns:
        list: List of essential documentation file paths
    """
    docs_config = load_docs_config()
    essential_docs = []
    
    # Add all registered documents from config
    for doc in docs_config.get("documents", []):
        doc_path = doc.get("path", "")
        if doc_path:
            # Convert to absolute path
            abs_path = os.path.join(DOCS_DIR, doc_path)
            essential_docs.append(abs_path)
    
    # Add main documentation files
    essential_docs.extend([
        DOCS_CONFIG_PATH,
        os.path.join(DOCS_DIR, "README.md"),
        os.path.join(DOCS_DIR, "manage_docs.py")
    ])
    
    return essential_docs

def detect_codebase_files():
    """
    Intelligently detect files that belong to the codebase by analyzing:
    - Python files with imports from the project
    - Configuration files referenced by the code
    - Files that are imported or used by core scripts
    - Documentation files from the documentation system
    
    Returns:
        set: Set of files that are part of the codebase
    """
    codebase_files = set(ESSENTIAL_FILES)
    
    # Add essential documentation files
    doc_files = get_essential_doc_files()
    for doc_file in doc_files:
        if os.path.isfile(doc_file):
            # Store relative path from ROOT_DIR
            rel_path = os.path.relpath(doc_file, ROOT_DIR)
            codebase_files.add(rel_path)
    
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
                    r'docs_config\.json',
                    r'requirements\.txt'
                ]
                
                for pattern in config_patterns:
                    if re.search(pattern, content):
                        codebase_files.add(item)
                        break
                        
            except Exception as e:
                print(f"Error analyzing {item}: {e}")
    
    return codebase_files

def create_backup():
    """
    Create a timestamped backup of all codebase files before decluttering.
    
    Returns:
        str: Path to the created backup file, or None if backup failed
    """
    try:
        # Create backup directory if it doesn't exist
        os.makedirs(BACKUP_DIR, exist_ok=True)
        
        # Generate timestamp for backup filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"codebase_backup_{timestamp}.zip"
        backup_path = os.path.join(BACKUP_DIR, backup_filename)
        
        # Detect codebase files to backup
        codebase_files = detect_codebase_files()
        
        # Create ZIP archive
        with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add individual files
            for file in codebase_files:
                file_path = os.path.join(ROOT_DIR, file)
                if os.path.isfile(file_path):
                    zipf.write(file_path, file)
            
            # Add essential directories
            for dir_name in ESSENTIAL_DIRS:
                dir_path = os.path.join(ROOT_DIR, dir_name)
                if os.path.isdir(dir_path):
                    for root, _, files in os.walk(dir_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            # Store with relative path
                            rel_path = os.path.relpath(file_path, ROOT_DIR)
                            zipf.write(file_path, rel_path)
        
        print(f"✅ Backup created at {backup_path}")
        return backup_path
    
    except Exception as e:
        print(f"❌ Backup failed: {e}")
        return None

def cleanup_old_backups():
    """
    Remove old backup files, keeping only the most recent MAX_BACKUPS.
    """
    try:
        if not os.path.exists(BACKUP_DIR):
            return
            
        backups = []
        for file in os.listdir(BACKUP_DIR):
            if file.startswith("codebase_backup_") and file.endswith(".zip"):
                file_path = os.path.join(BACKUP_DIR, file)
                backups.append((file_path, os.path.getmtime(file_path)))
        
        # Sort by modification time (newest first)
        backups.sort(key=lambda x: x[1], reverse=True)
        
        # Remove old backups
        if len(backups) > MAX_BACKUPS:
            for file_path, _ in backups[MAX_BACKUPS:]:
                os.remove(file_path)
                print(f"Removed old backup: {os.path.basename(file_path)}")
    
    except Exception as e:
        print(f"Warning: Error cleaning up old backups: {e}")

def create_archive_subdirs():
    """Create the necessary archive subdirectories if they don't exist."""
    archive_dir = os.path.join(ROOT_DIR, 'archive')
    os.makedirs(archive_dir, exist_ok=True)
    
    # Define archive subdirectories
    subdirs = [
        'debugging',    # Debug scripts, diagnostic tools
        'testing',      # Test scripts and outputs
        'reports',      # Generated Excel reports
        'temporary',    # Temporary files and markers
        'documentation', # Non-essential documentation
        'legacy',       # Legacy files and backups
        'outputs',      # Generated outputs and exports
    ]
    
    # Create each subdirectory
    for subdir in subdirs:
        subdir_path = os.path.join(archive_dir, subdir)
        os.makedirs(subdir_path, exist_ok=True)
        print(f"✓ Archive subdirectory ready: {subdir}")

def categorize_file(filename):
    """
    Determine the appropriate archive subdirectory for a non-essential file.
    
    Args:
        filename: Name of the file to categorize
        
    Returns:
        The archive subdirectory where the file should be moved, or None if it should stay
    """
    # Check if file is in essential list
    if filename in ESSENTIAL_FILES:
        return None
    
    # Check if file is a directory
    if os.path.isdir(os.path.join(ROOT_DIR, filename)):
        if filename in ESSENTIAL_DIRS:
            return None
        else:
            return 'legacy'  # Move non-essential directories to legacy
    
    # Categorize by filename patterns
    
    # Debug files
    if (filename.startswith(('debug_', 'diagnose_', 'trace_', 'analyze_', 'fix_')) or
        'debug' in filename.lower() or 'trace' in filename.lower()):
        return 'debugging'
    
    # Test files
    if (filename.startswith(('test_', 'verify_', 'check_')) or
        'test' in filename.lower() or 'unittest' in filename.lower()):
        return 'testing'
    
    # Reports
    if (filename.startswith('AR_Analysis_Report_') or 
        filename.endswith(('.xlsx', '.xls')) or
        'report' in filename.lower()):
        return 'reports'
    
    # Temporary files
    if (filename.startswith(('MARKER_', 'temp_', 'tmp_')) or
        filename.endswith(('.tmp', '.temp')) or
        'temporary' in filename.lower()):
        return 'temporary'
    
    # Documentation files
    if (filename.endswith(('.md', '.txt', '.pdf', '.docx')) and 
        filename != 'README.md'):
        return 'documentation'
    
    # Output files
    if (filename.endswith(('.json', '.csv', '.png', '.jpg')) and
        not filename.endswith(('config.json', 'report_config.json'))):
        return 'outputs'
    
    # Legacy Python files (not in essential list)
    if filename.endswith('.py') and filename not in ESSENTIAL_FILES:
        return 'legacy'
    
    # Default for unrecognized files
    return 'legacy'

def move_file(filepath, dest_subdir):
    """
    Move a file to the appropriate archive subdirectory.
    
    Args:
        filepath: Full path to the file to move
        dest_subdir: Destination subdirectory under archive/
    """
    try:
        filename = os.path.basename(filepath)
        dest_dir = os.path.join(ROOT_DIR, 'archive', dest_subdir)
        dest_path = os.path.join(dest_dir, filename)
        
        # Handle file already exists in destination
        if os.path.exists(dest_path):
            # Add timestamp to filename
            name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            new_filename = f"{name}_{timestamp}{ext}"
            dest_path = os.path.join(dest_dir, new_filename)
        
        # Move the file
        shutil.move(filepath, dest_path)
        return dest_path
    except Exception as e:
        print(f"Error moving {filepath}: {e}")
        return None

def update_documentation_index():
    """
    Update the documentation index after decluttering.
    """
    try:
        # Check if manage_docs.py exists
        manage_docs_path = os.path.join(DOCS_DIR, 'manage_docs.py')
        if not os.path.exists(manage_docs_path):
            print("Documentation manager not found, skipping index update")
            return False
            
        # Run the documentation manager to update index
        print("\nUpdating documentation index...")
        current_dir = os.getcwd()
        os.chdir(DOCS_DIR)
        
        try:
            # Import and run the documentation manager
            sys.path.insert(0, DOCS_DIR)
            from manage_docs import DocumentationManager
            
            manager = DocumentationManager('.')
            manager.update_index()
            print("✅ Documentation index updated successfully")
            
            # Validate documentation
            if manager.validate_documentation():
                print("✅ Documentation validation passed")
            else:
                print("⚠️ Documentation validation found issues")
                
        except Exception as e:
            print(f"Error updating documentation index: {e}")
            return False
        finally:
            os.chdir(current_dir)
            
        return True
    except Exception as e:
        print(f"Error in documentation update: {e}")
        return False

def declutter_directory():
    """
    Main function to declutter the directory by moving files to appropriate archive locations.
    """
    print("\n🔍 Starting directory decluttering process...")
    
    # Step 1: Detect codebase files
    print("\n📊 Detecting codebase files...")
    codebase_files = detect_codebase_files()
    print(f"✅ Detected {len(codebase_files)} codebase files to protect")
    
    # Step 2: Create backup
    print("\n💾 Creating backup of codebase files...")
    backup_path = create_backup()
    if not backup_path:
        print("❌ Backup failed, aborting decluttering for safety")
        return
    
    # Step 3: Clean up old backups
    cleanup_old_backups()
    
    # Step 4: Create archive subdirectories
    print("\n📁 Setting up archive structure...")
    create_archive_subdirs()
    
    # Step 5: Process files
    print("\n🔄 Processing files...")
    moved_files = {
        'debugging': [],
        'testing': [],
        'reports': [],
        'temporary': [],
        'documentation': [],
        'legacy': [],
        'outputs': []
    }
    
    # Get documentation files to protect
    doc_files = get_essential_doc_files()
    doc_rel_paths = [os.path.relpath(path, ROOT_DIR) for path in doc_files if os.path.isfile(path)]
    
    # Process each file in the root directory
    for item in os.listdir(ROOT_DIR):
        # Skip if it's a directory in ESSENTIAL_DIRS
        if os.path.isdir(os.path.join(ROOT_DIR, item)) and item in ESSENTIAL_DIRS:
            continue
            
        # Skip if it's a codebase file
        if item in codebase_files:
            print(f"  🔒 Protecting codebase file: {item}")
            continue
            
        # Skip if it's a documentation file
        if item in doc_rel_paths or os.path.join('docs', item) in doc_rel_paths:
            print(f"  📚 Protecting documentation file: {item}")
            continue
        
        # Determine appropriate archive category
        category = categorize_file(item)
        if category:
            filepath = os.path.join(ROOT_DIR, item)
            print(f"  📦 Moving: {item} → archive/{category}/")
            
            # Move the file
            moved_path = move_file(filepath, category)
            if moved_path:
                moved_files[category].append(item)
    
    # Step 6: Generate recovery guide
    guide_path = os.path.join(ROOT_DIR, 'archive', 'RECOVERY_GUIDE.md')
    generate_recovery_guide(moved_files, guide_path, backup_path)
    
    # Step 7: Save operation log
    log_path = os.path.join(ROOT_DIR, 'archive', 'declutter_log.json')
    log_data = {
        'timestamp': datetime.now().isoformat(),
        'backup_path': backup_path,
        'moved_files': moved_files,
        'protected_files': list(codebase_files)
    }
    
    with open(log_path, 'w') as f:
        json.dump(log_data, f, indent=2)
    
    # Step 8: Update documentation index
    update_documentation_index()
    
    # Summary
    total_moved = sum(len(files) for files in moved_files.values())
    print(f"\n✅ Decluttering complete! Moved {total_moved} files to archive/")
    print(f"📝 Recovery guide generated at {guide_path}")
    print(f"📊 Operation log saved to {log_path}")
    print(f"💾 Backup saved to {backup_path}")

def generate_recovery_guide(moved_files, guide_path, backup_path):
    """
    Generate a recovery guide markdown file that explains how to restore moved files if needed.
    
    Args:
        moved_files: Dictionary of moved files by category
        guide_path: Path to write the recovery guide
        backup_path: Path to the backup file
    """
    with open(guide_path, 'w') as f:
        f.write("# AR Data Analysis - Decluttering Recovery Guide\n\n")
        f.write(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        f.write("## Backup Information\n\n")
        f.write(f"A backup of all essential codebase files was created before decluttering:\n\n")
        f.write(f"- **Backup file:** `{os.path.basename(backup_path)}`\n")
        f.write(f"- **Location:** `{os.path.dirname(backup_path)}`\n")
        f.write(f"- **Size:** {os.path.getsize(backup_path) / (1024*1024):.2f} MB\n\n")
        
        f.write("To restore all files from backup:\n\n")
        f.write("```bash\n")
        f.write(f"# Extract backup to temporary directory\n")
        f.write(f"mkdir temp_restore\n")
        f.write(f"unzip \"{backup_path}\" -d temp_restore\n\n")
        f.write(f"# Copy files back as needed\n")
        f.write(f"cp -r temp_restore/* ./\n")
        f.write("```\n\n")
        
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
        
        f.write("## Documentation System\n\n")
        f.write("The documentation system is maintained in the `docs/` directory and includes:\n\n")
        f.write("- Configuration: `docs/docs_config.json`\n")
        f.write("- Management: `docs/manage_docs.py`\n")
        f.write("- Categories:\n")
        
        # Add documentation categories from config
        try:
            docs_config = load_docs_config()
            for category in docs_config.get("categories", []):
                name = category.get("name", "")
                path = category.get("path", "")
                desc = category.get("description", "")
                if name and path:
                    f.write(f"  - `{path}/`: {name} - {desc}\n")
        except Exception:
            f.write("  - Documentation categories not available\n")
        
        f.write("\n")
        
        f.write("## Essential Files (Kept in Root)\n\n")
        f.write("The following files remain in the root directory for normal operation:\n\n")
        for file in sorted(ESSENTIAL_FILES):
            f.write(f"- `{file}`\n")
        f.write("\n")
        
        f.write("## Essential Directories (Kept in Root)\n\n")
        for dir_name in sorted(ESSENTIAL_DIRS):
            f.write(f"- `{dir_name}/`\n")

if __name__ == "__main__":
    print("AR Data Analysis - Enhanced Directory Decluttering with Backup and Documentation Integration")
    print("=" * 80)
    print("\n🔍 This script will:")
    print("  1. Automatically detect codebase files")
    print("  2. Protect documentation files and structure")
    print("  3. Create a backup of all codebase files")
    print("  4. Organize non-essential files into archive subdirectories")
    print("  5. Update documentation index after decluttering")
    print(f"  6. Keep the last {MAX_BACKUPS} backups for safety")
    
    # Show what would be detected as codebase
    print("\n🔍 Analyzing current directory...")
    detected = detect_codebase_files()
    print(f"\nDetected {len(detected)} codebase files that will be protected:")
    for file in sorted(list(detected)[:10]):  # Show first 10
        print(f"  🔒 {file}")
    if len(detected) > 10:
        print(f"  ... and {len(detected) - 10} more")
    
    # Show documentation integration
    print("\n📚 Documentation integration:")
    try:
        docs_config = load_docs_config()
        categories = docs_config.get("categories", [])
        documents = docs_config.get("documents", [])
        print(f"  - Found {len(categories)} documentation categories")
        print(f"  - Found {len(documents)} registered documents")
        print("  - Documentation will be preserved and index updated")
    except Exception as e:
        print(f"  - Warning: Could not analyze documentation: {e}")
    
    response = input("\n❓ Proceed with backup and decluttering? (y/N): ")
    
    if response.lower() in ['y', 'yes']:
        declutter_directory()
    else:
        print("Decluttering cancelled.")
