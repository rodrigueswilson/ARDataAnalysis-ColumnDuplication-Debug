"""
Daily Count Pipelines for AR Data Analysis

This module contains MongoDB aggregation pipelines for daily-level analysis.
These pipelines aggregate file counts, sizes, and metadata by date (ISO_Date).

Key Pipelines:
- DAILY_COUNTS_ALL: Basic daily aggregation
- DAILY_COUNTS_ALL_WITH_ZEROES: Daily aggregation designed for time series analysis
- DAILY_COUNTS_COLLECTION_ONLY: Daily counts filtered to collection days only
"""

# =============================================================================
# 1. DAILY_COUNTS_ALL
# =============================================================================
# Aggregates file counts and total size per day.
# - Groups records by ISO_Date.
# - Calculates total files, MP3 files, JPG files, and total size in MB.
# - Sorts the results by date.
# This pipeline provides a simple daily summary of data volume.
DAILY_COUNTS_ALL = [
    {
        "$group": {
            "_id": "$ISO_Date",
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"}
        }
    },
    {
        "$addFields": {
            "has_files": {"$gt": ["$Total_Files", 0]}
        }
    },
    {"$sort": {"_id": 1}}
]

# =============================================================================
# 2. DAILY_COUNTS_ALL_WITH_ZEROES (for ACF/PACF with complete time series)
# =============================================================================
# This pipeline is designed to include ALL collection days, even those with zero files.
# It's essential for proper time series analysis where missing days need to be represented
# with zero counts rather than being omitted entirely. The 'has_files' column will be
# FALSE for collection days with no files, which is critical for ACF/PACF analysis.
DAILY_COUNTS_ALL_WITH_ZEROES = [
    {
        "$group": {
            "_id": "$ISO_Date",
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"}
        }
    },
    {
        "$addFields": {
            "has_files": {"$gt": ["$Total_Files", 0]}
        }
    },
    {"$sort": {"_id": 1}}
]

# =============================================================================
# 3. DAILY_COUNTS_COLLECTION_ONLY
# =============================================================================
# Daily counts filtered to only include collection days (excludes holidays, breaks, etc.)
# This pipeline is used for analysis that should focus only on actual school collection days.
DAILY_COUNTS_COLLECTION_ONLY = [
    {
        "$match": {
            "is_collection_day": True,
            "Outlier_Status": False
        }
    },
    {
        "$group": {
            "_id": "$ISO_Date",
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"}
        }
    },
    {
        "$addFields": {
            "has_files": {"$gt": ["$Total_Files", 0]}
        }
    },
    {"$sort": {"_id": 1}}
]

# Export all daily pipelines
DAILY_PIPELINES = {
    "DAILY_COUNTS_ALL": DAILY_COUNTS_ALL,
    "DAILY_COUNTS_ALL_WITH_ZEROES": DAILY_COUNTS_ALL_WITH_ZEROES,
    "DAILY_COUNTS_COLLECTION_ONLY": DAILY_COUNTS_COLLECTION_ONLY
}
