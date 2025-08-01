#!/usr/bin/env python
"""
Enhanced Declutter Script with Documentation Integration (Version 4)
Adapted for AR Data Analysis Codebase

This script helps declutter a project directory by:
1. Detecting essential codebase files (Python imports, config references, docs)
2. Creating a backup of essential files
3. Moving non-essential files to archive subdirectories
4. Generating a recovery guide
5. Updating the documentation index after decluttering

Improvements in v4:
- Updated ESSENTIAL_FILES list to match current AR Data Analysis codebase
- Updated ESSENTIAL_DIRS to include pipelines directory
- Enhanced protection for current project structure
- Maintained all safety features from v3
"""

import os
import sys
import re
import ast
import json
import shutil
import zipfile
import datetime
import importlib.util
from pathlib import Path
from typing import Set, Dict, List, Tuple, Optional, Any

# Constants - Updated for AR Data Analysis codebase
ESSENTIAL_DIRS = ['docs', 'utils', 'report_generator', 'pipelines', 'db', 'venv', '__pycache__']
ESSENTIAL_FILES = [
    # Core analysis modules
    'ar_utils.py',
    'acf_pacf_charts.py', 
    'dashboard_generator.py',
    'db_utils.py',
    'chart_config_helper.py',
    'column_cleanup_utils.py',
    
    # Main entry points
    'generate_report.py',
    'populate_db.py',
    'tag_jpg_files.py',
    'tag_mp3_files.py',
    
    # Configuration files
    'report_config.json',
    'config.yaml',
    'requirements.txt',
    
    # Documentation
    'README.md',
    
    # Declutter script
    'declutter.py',
    
    # Git configuration
    '.gitignore'
]

# Archive directories
ARCHIVE_DIR = 'archive'
ARCHIVE_SUBDIRS = {
    'debugging': ['debug', 'log', 'trace'],
    'testing': ['test', 'mock', 'fixture'],
    'reports': ['report', 'output', 'result', 'export'],
    'temporary': ['temp', 'tmp', 'draft', 'scratch'],
    'documentation': ['doc', 'docs', 'guide', 'tutorial'],
    'legacy': ['old', 'deprecated', 'archive', 'bak', 'backup'],
    'outputs': ['output', 'generated', 'result'],
    'migration': ['migrate', 'migration', 'convert']
}

# File patterns for categorization
FILE_PATTERNS = {
    'debugging': [r'debug', r'log', r'\.log$', r'trace', r'diagnostic'],
    'testing': [r'test', r'mock', r'fixture', r'sample'],
    'reports': [r'report', r'output', r'result', r'export', r'\.xlsx$', r'\.csv$'],
    'temporary': [r'temp', r'tmp', r'draft', r'scratch', r'\.tmp$', r'MARKER_'],
    'documentation': [r'doc', r'guide', r'tutorial', r'\.md$', r'\.pdf$'],
    'legacy': [r'old', r'deprecated', r'archive', r'bak', r'backup', r'v\d+'],
    'outputs': [r'output', r'generated', r'result', r'\.png$', r'\.jpg$', r'\.jpeg$', r'\.svg$', r'\.html$'],
    'migration': [r'migrate', r'migration', r'convert']
}

# Extensions to categorize
EXTENSION_CATEGORIES = {
    '.log': 'debugging',
    '.tmp': 'temporary',
    '.bak': 'legacy',
    '.xlsx': 'reports',
    '.csv': 'reports',
    '.md': 'documentation',
    '.pdf': 'documentation',
    '.png': 'outputs',
    '.jpg': 'outputs',
    '.jpeg': 'outputs',
    '.svg': 'outputs',
    '.html': 'outputs'
}

# Maximum number of backups to keep
MAX_BACKUPS = 5

# Directories to exclude from backup
BACKUP_EXCLUDE_DIRS = ['.git', 'venv', 'db', '__pycache__', 'archive']

def get_timestamp() -> str:
    """Get current timestamp in YYYYMMDD_HHMMSS format."""
    now = datetime.datetime.now()
    return now.strftime("%Y%m%d_%H%M%S")

def load_docs_config() -> Dict[str, Any]:
    """Load documentation configuration from docs_config.json."""
    config_path = os.path.join('docs', 'docs_config.json')
    if not os.path.exists(config_path):
        print("⚠️ Documentation configuration not found. Proceeding without documentation integration.")
        return {}
    
    try:
        with open(config_path, 'r') as f:
            return json.load(f)
    except json.JSONDecodeError:
        print("⚠️ Error parsing documentation configuration. Proceeding without documentation integration.")
        return {}
    except Exception as e:
        print(f"⚠️ Error loading documentation configuration: {e}. Proceeding without documentation integration.")
        return {}

def get_essential_doc_files(docs_config: Dict[str, Any]) -> Set[str]:
    """Get essential documentation files from docs_config."""
    essential_docs = set()
    
    if not docs_config:
        return essential_docs
    
    # Add all registered documents
    if 'documents' in docs_config:
        for doc in docs_config['documents']:
            if 'path' in doc:
                essential_docs.add(doc['path'])
    
    # Add category files
    if 'categories' in docs_config:
        for category in docs_config['categories']:
            if 'files' in category:
                essential_docs.update(category['files'])
    
    return essential_docs

def detect_codebase_files() -> Set[str]:
    """
    Detect essential codebase files by analyzing:
    1. Python imports
    2. Configuration references
    3. Documentation files
    """
    codebase_files = set()
    
    # Add explicitly defined essential files
    for file in ESSENTIAL_FILES:
        if os.path.exists(file):
            codebase_files.add(os.path.normpath(file))
    
    # Add essential directories
    for dir_name in ESSENTIAL_DIRS:
        if os.path.exists(dir_name) and os.path.isdir(dir_name):
            codebase_files.add(os.path.normpath(dir_name))
    
    # Load documentation configuration
    docs_config = load_docs_config()
    essential_docs = get_essential_doc_files(docs_config)
    codebase_files.update(essential_docs)
    
    # Analyze Python files for imports and references
    for root, _, files in os.walk('.'):
        # Skip archive and excluded directories
        if any(excluded in root for excluded in BACKUP_EXCLUDE_DIRS + [ARCHIVE_DIR]):
            continue
            
        for file in files:
            if file.endswith('.py'):
                file_path = os.path.join(root, file)
                norm_path = os.path.normpath(file_path)
                
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                    
                    # Parse AST to find imports
                    try:
                        tree = ast.parse(content)
                        for node in ast.walk(tree):
                            if isinstance(node, (ast.Import, ast.ImportFrom)):
                                if isinstance(node, ast.ImportFrom) and node.module:
                                    # Check for project-specific imports
                                    if any(module in node.module for module in 
                                          ['ar_utils', 'db_utils', 'report_generator', 'pipelines', 'utils']):
                                        codebase_files.add(norm_path)
                                        break
                                elif isinstance(node, ast.Import):
                                    for alias in node.names:
                                        if any(module in alias.name for module in 
                                              ['ar_utils', 'db_utils', 'report_generator', 'pipelines', 'utils']):
                                            codebase_files.add(norm_path)
                                            break
                    except SyntaxError:
                        # If we can't parse, check for string patterns
                        pass
                    
                    # Check for configuration file references
                    config_patterns = [
                        'report_config.json',
                        'config.yaml',
                        'requirements.txt'
                    ]
                    
                    if any(pattern in content for pattern in config_patterns):
                        codebase_files.add(norm_path)
                
                except Exception:
                    # If we can't read the file, skip it
                    continue
    
    return codebase_files

def create_backup() -> Optional[str]:
    """Create a backup of codebase files."""
    timestamp = get_timestamp()
    backup_dir = os.path.join(ARCHIVE_DIR, 'backups')
    os.makedirs(backup_dir, exist_ok=True)
    
    backup_filename = f'codebase_backup_{timestamp}.zip'
    backup_path = os.path.join(backup_dir, backup_filename)
    
    print(f"💾 Creating backup: {backup_path}")
    
    try:
        with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
            # Add all codebase files
            codebase_files = detect_codebase_files()
            
            for file_path in codebase_files:
                if os.path.isfile(file_path):
                    zipf.write(file_path)
                elif os.path.isdir(file_path):
                    # Add directory contents
                    for root, _, files in os.walk(file_path):
                        for file in files:
                            full_path = os.path.join(root, file)
                            zipf.write(full_path)
        
        print(f"✅ Backup created successfully: {backup_path}")
        
        # Clean up old backups
        cleanup_old_backups(backup_dir)
        
        return backup_path
    
    except Exception as e:
        print(f"❌ Error creating backup: {e}")
        return None

def cleanup_old_backups(backup_dir: str) -> None:
    """Keep only the most recent MAX_BACKUPS backups."""
    try:
        backup_files = [f for f in os.listdir(backup_dir) if f.startswith('codebase_backup_') and f.endswith('.zip')]
        backup_files.sort(reverse=True)  # Most recent first
        
        for old_backup in backup_files[MAX_BACKUPS:]:
            old_path = os.path.join(backup_dir, old_backup)
            os.remove(old_path)
            print(f"🗑️ Removed old backup: {old_backup}")
    
    except Exception as e:
        print(f"⚠️ Error cleaning up old backups: {e}")

def categorize_file(file_path: str) -> str:
    """Categorize a file based on its name and extension."""
    file_name = os.path.basename(file_path).lower()
    file_ext = os.path.splitext(file_name)[1].lower()
    
    # Check extension first
    if file_ext in EXTENSION_CATEGORIES:
        return EXTENSION_CATEGORIES[file_ext]
    
    # Check patterns
    for category, patterns in FILE_PATTERNS.items():
        if any(re.search(pattern, file_name, re.IGNORECASE) for pattern in patterns):
            return category
    
    return 'outputs'  # Default category

def move_file_to_archive(file_path: str, category: str) -> bool:
    """Move a file to the appropriate archive subdirectory."""
    try:
        archive_subdir = os.path.join(ARCHIVE_DIR, category)
        os.makedirs(archive_subdir, exist_ok=True)
        
        file_name = os.path.basename(file_path)
        dest_path = os.path.join(archive_subdir, file_name)
        
        # Handle name conflicts
        counter = 1
        while os.path.exists(dest_path):
            name, ext = os.path.splitext(file_name)
            dest_path = os.path.join(archive_subdir, f"{name}_{counter}{ext}")
            counter += 1
        
        shutil.move(file_path, dest_path)
        return True
    
    except Exception as e:
        print(f"❌ Error moving {file_path}: {e}")
        return False

def generate_recovery_guide(moved_files: Dict[str, List[str]], docs_config: Dict[str, Any]) -> None:
    """Generate a guide for recovering moved files."""
    timestamp = get_timestamp()
    guide_path = os.path.join(ARCHIVE_DIR, f'RECOVERY_GUIDE_{timestamp}.md')
    
    try:
        with open(guide_path, 'w') as f:
            f.write(f"# Recovery Guide - {timestamp}\n\n")
            f.write("This guide helps you recover files that were moved during decluttering.\n\n")
            
            f.write("## Backup Information\n")
            f.write(f"- Backup created: {timestamp}\n")
            f.write(f"- Backup location: `{os.path.join(ARCHIVE_DIR, 'backups', f'codebase_backup_{timestamp}.zip')}`\n\n")
            
            f.write("## Files Moved by Category\n\n")
            for category, files in moved_files.items():
                if files:
                    f.write(f"### {category.title()}\n")
                    f.write(f"Location: `{os.path.join(ARCHIVE_DIR, category)}`\n\n")
                    for file in files[:10]:  # Show first 10 files
                        f.write(f"- {file}\n")
                    if len(files) > 10:
                        f.write(f"- ... and {len(files) - 10} more files\n")
                    f.write("\n")
            
            f.write("## Recovery Instructions\n\n")
            f.write("### To recover specific files:\n")
            f.write("1. Navigate to the appropriate archive subdirectory\n")
            f.write("2. Copy or move the file back to the root directory\n\n")
            
            f.write("### To restore from backup:\n")
            f.write("```bash\n")
            f.write(f"unzip {os.path.join(ARCHIVE_DIR, 'backups', f'codebase_backup_{timestamp}.zip')} -d restore_dir\n")
            f.write("```\n\n")
            
            f.write("## Documentation Integration\n")
            if docs_config:
                f.write("- Documentation system is active\n")
                f.write("- Essential documentation files were protected\n")
            else:
                f.write("- No documentation configuration found\n")
            f.write("\n")
        
        print(f"✅ Recovery guide created: {guide_path}")
    except Exception as e:
        print(f"❌ Error creating recovery guide: {e}")

def update_documentation_index() -> bool:
    """Update the documentation index after decluttering."""
    print("📚 Updating documentation index...")
    
    docs_manager_path = os.path.join('docs', 'manage_docs.py')
    if not os.path.exists(docs_manager_path):
        print("⚠️ Documentation manager not found. Skipping index update.")
        return False
    
    try:
        # Import the DocumentationManager class
        spec = importlib.util.spec_from_file_location("manage_docs", docs_manager_path)
        manage_docs = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(manage_docs)
        
        # Create an instance of DocumentationManager
        doc_manager = manage_docs.DocumentationManager()
        
        # Update the index
        doc_manager.update_index()
        
        # Validate documentation
        doc_manager.validate_documentation()
        
        print("✅ Documentation index updated successfully.")
        return True
    except Exception as e:
        print(f"❌ Error updating documentation index: {e}")
        return False

def declutter_directory() -> None:
    """Main function to declutter the directory."""
    print("🧹 Starting directory decluttering (AR Data Analysis v4)...")
    print("📂 Detecting codebase files...")
    
    # Load documentation configuration
    docs_config = load_docs_config()
    if docs_config:
        print(f"📚 Found documentation configuration with {len(get_essential_doc_files(docs_config))} registered documents.")
    
    # Detect codebase files
    codebase_files = detect_codebase_files()
    print(f"✅ Detected {len(codebase_files)} essential codebase files.")
    
    # Show some detected files for verification
    print("🔍 Essential files detected (sample):")
    for i, file in enumerate(sorted(codebase_files)):
        if i < 10:  # Show first 10
            print(f"   - {file}")
        elif i == 10:
            print(f"   - ... and {len(codebase_files) - 10} more files")
            break
    
    # Create backup
    backup_path = create_backup()
    if not backup_path:
        print("❌ Backup failed, aborting decluttering for safety.")
        return
    
    # Create archive directory if it doesn't exist
    for subdir in ARCHIVE_SUBDIRS.keys():
        os.makedirs(os.path.join(ARCHIVE_DIR, subdir), exist_ok=True)
    
    # Find files to move
    files_to_move = []
    for root, _, files in os.walk('.'):
        # Skip archive directory and essential directories
        if (ARCHIVE_DIR in root or 
            any(excluded in root for excluded in BACKUP_EXCLUDE_DIRS) or
            any(root.startswith(f".{os.sep}{dir}") for dir in ESSENTIAL_DIRS)):
            continue
        
        for file in files:
            file_path = os.path.join(root, file)
            norm_path = os.path.normpath(file_path)
            
            # Skip essential files and files in the archive directory
            if norm_path in codebase_files or ARCHIVE_DIR in norm_path:
                continue
            
            files_to_move.append(norm_path)
    
    print(f"🔍 Found {len(files_to_move)} non-essential files to move.")
    
    # Show some files that will be moved for verification
    if files_to_move:
        print("📋 Files to be moved (sample):")
        for i, file in enumerate(files_to_move):
            if i < 5:  # Show first 5
                category = categorize_file(file)
                print(f"   - {file} → {category}")
            elif i == 5:
                print(f"   - ... and {len(files_to_move) - 5} more files")
                break
    
    # Ask for confirmation
    if files_to_move:
        response = input(f"\n🤔 Proceed with moving {len(files_to_move)} files to archive? (y/N): ")
        if response.lower() != 'y':
            print("❌ Decluttering cancelled by user.")
            return
    
    # Move files to archive
    moved_files = {category: [] for category in ARCHIVE_SUBDIRS.keys()}
    moved_count = 0
    
    for file_path in files_to_move:
        category = categorize_file(file_path)
        if move_file_to_archive(file_path, category):
            moved_files[category].append(file_path)
            moved_count += 1
    
    print(f"✅ Moved {moved_count} files to archive subdirectories.")
    
    # Generate recovery guide
    generate_recovery_guide(moved_files, docs_config)
    
    # Update documentation index
    update_documentation_index()
    
    print("✨ Directory decluttering completed successfully!")
    print(f"📝 Recovery guide created in the archive directory.")
    print(f"💾 Backup created: {backup_path}")
    print(f"📁 Files organized in: {ARCHIVE_DIR}/")

if __name__ == "__main__":
    # Create archive directory if it doesn't exist
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    
    # Run declutter
    declutter_directory()
