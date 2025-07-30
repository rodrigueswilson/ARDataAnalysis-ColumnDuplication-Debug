"""
Utility Functions for Report Generator
=====================================

This module contains utility functions for MongoDB connection management,
backup operations, and other helper functions used throughout the report
generation process.
"""

import os
import datetime
# Import db_utils conditionally to avoid import errors
try:
    from db_utils import get_db_connection
except ImportError:
    get_db_connection = None

def backup_py_and_md_files():
    """
    Creates a timestamped backup of all Python, Markdown, and configuration files
    in the ARDataAnalysis directory.
    """
    try:
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_dir = os.path.join('D:\\ARDataAnalysis\\backups', f'backup_{timestamp}')
        os.makedirs(backup_dir, exist_ok=True)
        
        # Backup .py files and other important files
        for file in os.listdir('D:\\ARDataAnalysis'):
            if any(file.endswith(ext) for ext in ['.py', '.md', '.json', '.txt.', '.txt', '.xlsx', '.yaml']):
                src = os.path.join('D:\\ARDataAnalysis', file)
                dst = os.path.join(backup_dir, file)
                try:
                    with open(src, 'rb') as fsrc, open(dst, 'wb') as fdst:
                        fdst.write(fsrc.read())
                except Exception as file_err:
                    print(f"[Backup Warning] Could not back up {file}: {file_err}")
        
        print(f"[Backup] Backup completed: {backup_dir}")
    except Exception as e:
        print(f"[Backup Error] Backup routine failed: {e}")


