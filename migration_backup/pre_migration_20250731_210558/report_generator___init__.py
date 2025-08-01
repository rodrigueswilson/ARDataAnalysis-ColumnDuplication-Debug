"""
Modular Report Generator Package for AR Data Analysis
====================================================

This package provides a modular, maintainable architecture for generating
comprehensive Excel reports from MongoDB data. The package is split into
logical components:

- core.py: Main ReportGenerator class and orchestration logic
- sheet_creators.py: Individual sheet creation methods
- formatters.py: Excel formatting and styling utilities
- dashboard.py: Dashboard and summary generation
- utils.py: Utility functions and helpers

The modular design improves maintainability, testability, and allows for
easier extension of report functionality.
"""

from .core import ReportGenerator
from .utils import (
    backup_py_and_md_files,
    get_db_connection
)

__all__ = [
    'ReportGenerator',
    'backup_py_and_md_files',
    'get_db_connection'
]
