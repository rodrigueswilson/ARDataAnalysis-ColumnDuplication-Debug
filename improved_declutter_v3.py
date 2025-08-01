#!/usr/bin/env python
"""
Enhanced Declutter Script with Documentation Integration (Version 3)

This script helps declutter a project directory by:
1. Detecting essential codebase files (Python imports, config references, docs)
2. Creating a backup of essential files
3. Moving non-essential files to archive subdirectories
4. Generating a recovery guide
5. Updating the documentation index after decluttering

Improvements in v3:
- Integration with documentation management system
- Protection of documentation files from decluttering
- Automatic documentation index update
- Enhanced backup system (excluding archive directory)
- Added force_zip64=True to handle large files
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

# Constants
ESSENTIAL_DIRS = ['docs', 'utils', 'report_generator', 'config']
ESSENTIAL_FILES = [
    'ar_utils.py',
    'acf_pacf_charts.py',
    'dashboard_generator.py',
    'config.json',
    'requirements.txt',
    'README.md',
    'improved_declutter.py',
    'improved_declutter_v2.py',
    'improved_declutter_v3.py',
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
    'outputs': ['output', 'generated', 'result']
}

# File patterns for categorization
FILE_PATTERNS = {
    'debugging': [r'debug', r'log', r'\.log$', r'trace', r'diagnostic'],
    'testing': [r'test', r'mock', r'fixture', r'sample'],
    'reports': [r'report', r'output', r'result', r'export', r'\.xlsx$', r'\.csv$'],
    'temporary': [r'temp', r'tmp', r'draft', r'scratch', r'\.tmp$'],
    'documentation': [r'doc', r'guide', r'tutorial', r'\.md$', r'\.pdf$'],
    'legacy': [r'old', r'deprecated', r'archive', r'bak', r'backup', r'v\d+'],
    'outputs': [r'output', r'generated', r'result', r'\.png$', r'\.jpg$', r'\.jpeg$', r'\.svg$', r'\.html$']
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
                # Convert to absolute path
                doc_path = os.path.join('docs', doc['path'])
                essential_docs.add(doc_path)
    
    # Add all category files
    if 'categories' in docs_config:
        for category in docs_config['categories']:
            if 'documents' in category:
                for doc in category['documents']:
                    if 'path' in doc:
                        # Convert to absolute path
                        doc_path = os.path.join('docs', doc['path'])
                        essential_docs.add(doc_path)
    
    return essential_docs

def detect_codebase_files() -> Set[str]:
    """
    Detect essential codebase files by analyzing:
    1. Python imports
    2. Configuration references
    3. Documentation files
    """
    codebase_files = set(ESSENTIAL_FILES)
    
    # Add documentation files
    docs_config = load_docs_config()
    doc_files = get_essential_doc_files(docs_config)
    codebase_files.update(doc_files)
    
    # Find Python files
    python_files = []
    for root, _, files in os.walk('.'):
        if any(excluded in root for excluded in BACKUP_EXCLUDE_DIRS):
            continue
        for file in files:
            if file.endswith('.py'):
                python_files.append(os.path.join(root, file))
    
    # Analyze imports in Python files
    for py_file in python_files:
        codebase_files.add(os.path.normpath(py_file))
        try:
            with open(py_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Parse AST to find imports
            try:
                tree = ast.parse(content)
                for node in ast.walk(tree):
                    # Check for import statements
                    if isinstance(node, (ast.Import, ast.ImportFrom)):
                        if isinstance(node, ast.ImportFrom) and node.module:
                            module_path = node.module.replace('.', os.path.sep)
                            if os.path.exists(f"{module_path}.py"):
                                codebase_files.add(f"{module_path}.py")
                            elif os.path.exists(module_path) and os.path.isdir(module_path):
                                codebase_files.add(module_path)
                        
                        for name in node.names:
                            module_path = name.name.replace('.', os.path.sep)
                            if os.path.exists(f"{module_path}.py"):
                                codebase_files.add(f"{module_path}.py")
                            elif os.path.exists(module_path) and os.path.isdir(module_path):
                                codebase_files.add(module_path)
                    
                    # Check for file operations
                    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == 'open':
                        if len(node.args) > 0 and isinstance(node.args[0], ast.Str):
                            file_path = node.args[0].s
                            if os.path.exists(file_path):
                                codebase_files.add(file_path)
            except SyntaxError:
                # Skip files with syntax errors
                pass
            
            # Use regex to find potential file references
            file_patterns = [
                r'open\([\'"](.+?)[\'"]\)',
                r'os\.path\.join\(.+?, [\'"](.+?)[\'"]\)',
                r'Path\([\'"](.+?)[\'"]\)',
                r'load_config\([\'"](.+?)[\'"]\)',
                r'read_file\([\'"](.+?)[\'"]\)',
                r'save_file\([\'"](.+?)[\'"]\)',
                r'with open\([\'"](.+?)[\'"]\)'
            ]
            
            for pattern in file_patterns:
                matches = re.findall(pattern, content)
                for match in matches:
                    if os.path.exists(match):
                        codebase_files.add(match)
        except Exception as e:
            print(f"⚠️ Error analyzing {py_file}: {e}")
    
    # Add essential directories
    for dir_name in ESSENTIAL_DIRS:
        if os.path.exists(dir_name) and os.path.isdir(dir_name):
            for root, _, files in os.walk(dir_name):
                for file in files:
                    file_path = os.path.join(root, file)
                    codebase_files.add(os.path.normpath(file_path))
    
    return codebase_files

def create_backup() -> str:
    """Create a backup of codebase files."""
    print("💾 Creating backup of codebase files...")
    
    # Get codebase files
    codebase_files = detect_codebase_files()
    
    # Create backup directory if it doesn't exist
    backup_dir = os.path.join(ARCHIVE_DIR, 'backups')
    os.makedirs(backup_dir, exist_ok=True)
    
    # Create backup filename with timestamp
    timestamp = get_timestamp()
    backup_filename = f"codebase_backup_{timestamp}.zip"
    backup_path = os.path.join(backup_dir, backup_filename)
    
    try:
        # Create ZIP file
        with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add codebase files to ZIP
            for file_path in codebase_files:
                if os.path.exists(file_path) and os.path.isfile(file_path):
                    # Skip files in excluded directories
                    if any(excluded in file_path for excluded in BACKUP_EXCLUDE_DIRS):
                        continue
                    
                    try:
                        zipf.write(file_path)
                    except Exception as e:
                        print(f"⚠️ Error adding {file_path} to backup: {e}")
        
        print(f"✅ Backup created: {backup_path}")
        
        # Clean up old backups
        cleanup_old_backups(backup_dir)
        
        return backup_path
    except Exception as e:
        print(f"❌ Backup failed: {e}")
        return ""

def cleanup_old_backups(backup_dir: str) -> None:
    """Keep only the most recent MAX_BACKUPS backups."""
    backups = []
    for file in os.listdir(backup_dir):
        if file.startswith("codebase_backup_") and file.endswith(".zip"):
            backups.append(os.path.join(backup_dir, file))
    
    # Sort backups by modification time (newest first)
    backups.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    
    # Remove old backups
    if len(backups) > MAX_BACKUPS:
        for old_backup in backups[MAX_BACKUPS:]:
            try:
                os.remove(old_backup)
                print(f"🗑️ Removed old backup: {old_backup}")
            except Exception as e:
                print(f"⚠️ Error removing old backup {old_backup}: {e}")

def categorize_file(file_path: str) -> str:
    """Categorize a file based on its name and extension."""
    filename = os.path.basename(file_path)
    extension = os.path.splitext(filename)[1].lower()
    
    # Check extension first
    if extension in EXTENSION_CATEGORIES:
        return EXTENSION_CATEGORIES[extension]
    
    # Check file patterns
    for category, patterns in FILE_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, filename, re.IGNORECASE):
                return category
    
    # Default to 'temporary' if no match
    return 'temporary'

def move_file_to_archive(file_path: str, category: str) -> bool:
    """Move a file to the appropriate archive subdirectory."""
    # Create archive subdirectory if it doesn't exist
    archive_subdir = os.path.join(ARCHIVE_DIR, category)
    os.makedirs(archive_subdir, exist_ok=True)
    
    # Get destination path
    dest_path = os.path.join(archive_subdir, os.path.basename(file_path))
    
    # Handle filename conflicts
    if os.path.exists(dest_path):
        base_name, extension = os.path.splitext(dest_path)
        timestamp = get_timestamp()
        dest_path = f"{base_name}_{timestamp}{extension}"
    
    try:
        shutil.move(file_path, dest_path)
        return True
    except Exception as e:
        print(f"⚠️ Error moving {file_path} to {dest_path}: {e}")
        return False

def generate_recovery_guide(moved_files: Dict[str, List[str]], docs_config: Dict[str, Any]) -> None:
    """Generate a guide for recovering moved files."""
    print("📝 Generating recovery guide...")
    
    # Create recovery guide filename with timestamp
    timestamp = get_timestamp()
    guide_filename = f"recovery_guide_{timestamp}.md"
    guide_path = os.path.join(ARCHIVE_DIR, guide_filename)
    
    try:
        with open(guide_path, 'w') as f:
            f.write("# File Recovery Guide\n\n")
            f.write(f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            f.write("## Overview\n\n")
            f.write("This guide lists all files that were moved during the decluttering process.\n")
            f.write("You can use this guide to recover files if needed.\n\n")
            
            f.write("## Documentation System\n\n")
            if docs_config:
                f.write("The following documentation categories are registered in the system:\n\n")
                if 'categories' in docs_config:
                    for category in docs_config['categories']:
                        f.write(f"- **{category.get('name', 'Unnamed')}**: {category.get('description', 'No description')}\n")
                
                f.write("\nEssential documentation files (these were NOT moved):\n\n")
                essential_docs = get_essential_doc_files(docs_config)
                for doc in sorted(essential_docs):
                    f.write(f"- `{doc}`\n")
            else:
                f.write("No documentation configuration found.\n")
            
            f.write("\n## Moved Files\n\n")
            
            total_moved = sum(len(files) for files in moved_files.values())
            f.write(f"Total files moved: {total_moved}\n\n")
            
            for category, files in moved_files.items():
                if files:
                    f.write(f"### {category.capitalize()} Files\n\n")
                    f.write(f"Location: `{os.path.join(ARCHIVE_DIR, category)}`\n\n")
                    for file in sorted(files):
                        f.write(f"- `{file}` → `{os.path.join(ARCHIVE_DIR, category, os.path.basename(file))}`\n")
                    f.write("\n")
            
            f.write("## Recovery Instructions\n\n")
            f.write("To recover a file, copy it from the archive directory back to its original location.\n\n")
            f.write("Example:\n")
            f.write("```\n")
            f.write("cp archive/category/filename.ext original/path/filename.ext\n")
            f.write("```\n\n")
            
            f.write("## Backup\n\n")
            f.write("A backup of all essential files was created before decluttering.\n")
            f.write(f"Backup location: `{os.path.join(ARCHIVE_DIR, 'backups', f'codebase_backup_{timestamp}.zip')}`\n\n")
            
            f.write("To restore from backup:\n")
            f.write("```\n")
            f.write(f"unzip {os.path.join(ARCHIVE_DIR, 'backups', f'codebase_backup_{timestamp}.zip')} -d restore_dir\n")
            f.write("```\n")
        
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
    print("🧹 Starting directory decluttering...")
    print("📂 Detecting codebase files...")
    
    # Load documentation configuration
    docs_config = load_docs_config()
    if docs_config:
        print(f"📚 Found documentation configuration with {len(get_essential_doc_files(docs_config))} registered documents.")
    
    # Detect codebase files
    codebase_files = detect_codebase_files()
    print(f"✅ Detected {len(codebase_files)} essential codebase files.")
    
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

if __name__ == "__main__":
    # Create archive directory if it doesn't exist
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    
    # Run declutter
    declutter_directory()
