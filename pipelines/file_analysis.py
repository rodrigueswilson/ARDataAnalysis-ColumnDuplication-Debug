"""
File Analysis Pipelines for AR Data Analysis

This module contains MongoDB aggregation pipelines for analyzing file characteristics,
including file sizes, capture volumes, and distribution patterns.

Key Pipelines:
- FILE_SIZE_SUMMARY_BY_DAY: Daily file size summaries
- CAPTURE_VOLUME_PER_MONTH: Monthly capture volume analysis
- FILE_SIZE_STATS: File size statistics by type
- AUDIO_EFFICIENCY_ANALYSIS: Audio efficiency metrics
"""

# =============================================================================
# FILE_SIZE_SUMMARY_BY_DAY
# =============================================================================
# Summarizes file size statistics on a daily basis.
# - Groups records by ISO_Date.
# - Calculates the total number of files, the total size in MB, and the average
#   file size in MB for each day.
# - This pipeline is useful for monitoring daily data volume and identifying
#   anomalies in file sizes.
FILE_SIZE_SUMMARY_BY_DAY = [
    {
        "$group": {
            "_id": "$ISO_Date",
            "Total_Files": {"$sum": 1},
            "Total_Size_MB": {"$sum": "$File_Size_MB"},
            "Avg_Size_MB": {"$avg": "$File_Size_MB"}
        }
    },
    {"$sort": {"_id": 1}}
]

# =============================================================================
# CAPTURE_VOLUME_PER_MONTH
# =============================================================================
# Aggregates the total number of files captured per month.
# - Groups records by a composite key of year and month from the ISO_Date.
# - Provides a high-level overview of data collection volume over time.
# - Sorts chronologically.
# - Fixed to properly convert ISO_Date string to date object before extraction.
CAPTURE_VOLUME_PER_MONTH = [
    {
        "$group": {
            "_id": {
                "Year": {
                    "$year": {
                        "$dateFromString": {
                            "dateString": "$ISO_Date"
                        }
                    }
                },
                "Month": {
                    "$month": {
                        "$dateFromString": {
                            "dateString": "$ISO_Date"
                        }
                    }
                }
            },
            "Count": {"$sum": 1}
        }
    },
    {
        "$sort": {
            "_id.Year": 1,
            "_id.Month": 1
        }
    }
]

# =============================================================================
# FILE_SIZE_STATS
# =============================================================================
# Calculates basic file size statistics for each file type (JPG, MP3, etc.).
# - Groups records by 'file_type'.
# - For each type, it calculates the average, minimum, and maximum size in MB,
#   as well as the total count of files.
# - This provides a quick overview of the size characteristics of different data types.
FILE_SIZE_STATS = [
    {
        "$group": {
            "_id": "$file_type",
            "AvgSizeMB": {"$avg": "$File_Size_MB"},
            "MinSizeMB": {"$min": "$File_Size_MB"},
            "MaxSizeMB": {"$max": "$File_Size_MB"},
            "Count": {"$sum": 1}
        }
    }
]

# =============================================================================
# AUDIO_EFFICIENCY_ANALYSIS
# =============================================================================
# Analyzes audio recording efficiency by day
AUDIO_EFFICIENCY_ANALYSIS = [
    {
        "$match": {
            "file_type": "MP3",
            "Duration_Seconds": {"$exists": True, "$gt": 0}
        }
    },
    {
        "$group": {
            "_id": "$ISO_Date",
            "total_mp3_files": {"$sum": 1},
            "total_duration_minutes": {"$sum": {"$divide": ["$Duration_Seconds", 60]}},
            "avg_duration_seconds": {"$avg": "$Duration_Seconds"},
            "min_duration_seconds": {"$min": "$Duration_Seconds"},
            "max_duration_seconds": {"$max": "$Duration_Seconds"}
        }
    },
    {
        "$addFields": {
            "efficiency_files_per_audio_minute": {
                "$cond": {
                    "if": {"$gt": ["$total_duration_minutes", 0]},
                    "then": {"$divide": ["$total_mp3_files", "$total_duration_minutes"]},
                    "else": 0
                }
            },
            "total_duration_hours": {"$divide": ["$total_duration_minutes", 60]}
        }
    },
    {
        "$project": {
            "_id": 0,
            "date": "$_id",
            "total_mp3_files": 1,
            "total_duration_minutes": {"$round": ["$total_duration_minutes", 2]},
            "total_duration_hours": {"$round": ["$total_duration_hours", 2]},
            "avg_duration_seconds": {"$round": ["$avg_duration_seconds", 1]},
            "min_duration_seconds": 1,
            "max_duration_seconds": 1,
            "efficiency_files_per_audio_minute": {"$round": ["$efficiency_files_per_audio_minute", 3]}
        }
    },
    {"$sort": {"date": 1}}
]

# Export all file analysis pipelines
FILE_PIPELINES = {
    "FILE_SIZE_SUMMARY_BY_DAY": FILE_SIZE_SUMMARY_BY_DAY,
    "CAPTURE_VOLUME_PER_MONTH": CAPTURE_VOLUME_PER_MONTH,
    "FILE_SIZE_STATS": FILE_SIZE_STATS,
    "AUDIO_EFFICIENCY_ANALYSIS": AUDIO_EFFICIENCY_ANALYSIS
}
