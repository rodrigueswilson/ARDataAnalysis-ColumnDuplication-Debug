"""
Dashboard Data Pipelines for AR Data Analysis

This module contains MongoDB aggregation pipelines specifically designed to provide
data for dashboard components and summary statistics.

Key Pipelines:
- DATA_QUALITY_CHECK: Data quality metrics
- DASHBOARD_YEAR_SUMMARY: Year-over-year summary data
- DASHBOARD_PERIOD_SUMMARY: Period breakdown data
- DASHBOARD_DATA_QUALITY: Data quality indicators
- DASHBOARD_COVERAGE: Coverage analysis
- DATA_CLEANING_SUMMARY: Data cleaning metrics
"""

# =============================================================================
# DATA_QUALITY_CHECK
# =============================================================================
# Comprehensive data quality analysis
DATA_QUALITY_CHECK = [
    {
        "$group": {
            "_id": None,
            "total_files": {"$sum": 1},
            "jpg_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "mp3_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "outliers": {"$sum": {"$cond": [{"$eq": ["$Outlier_Status", True]}, 1, 0]}},
            "collection_day_files": {"$sum": {"$cond": [{"$eq": ["$is_collection_day", True]}, 1, 0]}},
            "total_size_mb": {"$sum": "$File_Size_MB"},
            "avg_file_size_mb": {"$avg": "$File_Size_MB"},
            "unique_dates": {"$addToSet": "$ISO_Date"},
            "unique_school_years": {"$addToSet": "$School_Year"}
        }
    },
    {
        "$addFields": {
            "outlier_percentage": {
                "$multiply": [
                    {"$divide": ["$outliers", "$total_files"]}, 100
                ]
            },
            "collection_day_percentage": {
                "$multiply": [
                    {"$divide": ["$collection_day_files", "$total_files"]}, 100
                ]
            },
            "unique_date_count": {"$size": "$unique_dates"},
            "school_year_count": {"$size": "$unique_school_years"}
        }
    }
]

# =============================================================================
# DASHBOARD_YEAR_SUMMARY
# =============================================================================
# Year-over-year summary for dashboard
DASHBOARD_YEAR_SUMMARY = [
    {
        "$match": {
            "School_Year": {"$ne": "N/A"}
        }
    },
    {
        "$group": {
            "_id": "$School_Year",
            "total_files": {"$sum": 1},
            "jpg_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "mp3_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "total_size_mb": {"$sum": "$File_Size_MB"},
            "outliers": {"$sum": {"$cond": [{"$eq": ["$Outlier_Status", True]}, 1, 0]}},
            "unique_dates": {"$addToSet": "$ISO_Date"},
            "first_date": {"$min": "$ISO_Date"},
            "last_date": {"$max": "$ISO_Date"}
        }
    },
    {
        "$addFields": {
            "days_with_data": {"$size": "$unique_dates"},
            "avg_files_per_day": {
                "$round": [
                    {"$divide": ["$total_files", {"$size": "$unique_dates"}]}, 2
                ]
            },
            "outlier_percentage": {
                "$round": [
                    {"$multiply": [
                        {"$divide": ["$outliers", "$total_files"]}, 100
                    ]}, 2
                ]
            }
        }
    },
    {
        "$project": {
            "_id": 0,
            "school_year": "$_id",
            "total_files": 1,
            "jpg_files": 1,
            "mp3_files": 1,
            "total_size_mb": {"$round": ["$total_size_mb", 2]},
            "outliers": 1,
            "outlier_percentage": 1,
            "days_with_data": 1,
            "avg_files_per_day": 1,
            "first_date": 1,
            "last_date": 1
        }
    },
    {"$sort": {"school_year": 1}}
]

# =============================================================================
# DASHBOARD_PERIOD_SUMMARY
# =============================================================================
# Period breakdown for dashboard
DASHBOARD_PERIOD_SUMMARY = [
    {
        "$match": {
            "Collection_Period": {"$ne": "N/A"}
        }
    },
    {
        "$group": {
            "_id": {
                "school_year": "$School_Year",
                "period": "$Collection_Period"
            },
            "total_files": {"$sum": 1},
            "jpg_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "mp3_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "total_size_mb": {"$sum": "$File_Size_MB"},
            "unique_dates": {"$addToSet": "$ISO_Date"},
            "first_date": {"$min": "$ISO_Date"},
            "last_date": {"$max": "$ISO_Date"}
        }
    },
    {
        "$addFields": {
            "days_with_data": {"$size": "$unique_dates"},
            "avg_files_per_day": {
                "$round": [
                    {"$divide": ["$total_files", {"$size": "$unique_dates"}]}, 2
                ]
            }
        }
    },
    {
        "$project": {
            "_id": 0,
            "school_year": "$_id.school_year",
            "period": "$_id.period",
            "total_files": 1,
            "jpg_files": 1,
            "mp3_files": 1,
            "total_size_mb": {"$round": ["$total_size_mb", 2]},
            "days_with_data": 1,
            "avg_files_per_day": 1,
            "first_date": 1,
            "last_date": 1
        }
    },
    {
        "$sort": {
            "school_year": 1,
            "period": 1
        }
    }
]

# =============================================================================
# DASHBOARD_DATA_QUALITY
# =============================================================================
# Data quality indicators for dashboard
DASHBOARD_DATA_QUALITY = [
    {
        "$facet": {
            "overall_quality": [
                {
                    "$group": {
                        "_id": None,
                        "total_files": {"$sum": 1},
                        "outliers": {"$sum": {"$cond": [{"$eq": ["$Outlier_Status", True]}, 1, 0]}},
                        "collection_day_files": {"$sum": {"$cond": [{"$eq": ["$is_collection_day", True]}, 1, 0]}},
                        "files_with_duration": {"$sum": {"$cond": [{"$and": [{"$eq": ["$file_type", "MP3"]}, {"$gt": ["$Duration_Seconds", 0]}]}, 1, 0]}},
                        "mp3_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}}
                    }
                }
            ],
            "by_file_type": [
                {
                    "$group": {
                        "_id": "$file_type",
                        "total_files": {"$sum": 1},
                        "outliers": {"$sum": {"$cond": [{"$eq": ["$Outlier_Status", True]}, 1, 0]}},
                        "avg_size_mb": {"$avg": "$File_Size_MB"},
                        "total_size_mb": {"$sum": "$File_Size_MB"}
                    }
                }
            ]
        }
    }
]

# =============================================================================
# DASHBOARD_COVERAGE
# =============================================================================
# Coverage analysis for dashboard
DASHBOARD_COVERAGE = [
    {
        "$group": {
            "_id": {
                "school_year": "$School_Year",
                "date": "$ISO_Date"
            },
            "total_files": {"$sum": 1},
            "jpg_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "mp3_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "is_collection_day": {"$first": "$is_collection_day"}
        }
    },
    {
        "$group": {
            "_id": "$_id.school_year",
            "total_days": {"$sum": 1},
            "days_with_files": {"$sum": {"$cond": [{"$gt": ["$total_files", 0]}, 1, 0]}},
            "collection_days": {"$sum": {"$cond": [{"$eq": ["$is_collection_day", True]}, 1, 0]}},
            "collection_days_with_files": {"$sum": {"$cond": [{"$and": [{"$eq": ["$is_collection_day", True]}, {"$gt": ["$total_files", 0]}]}, 1, 0]}}
        }
    },
    {
        "$addFields": {
            "coverage_percentage": {
                "$round": [
                    {"$multiply": [
                        {"$divide": ["$days_with_files", "$total_days"]}, 100
                    ]}, 2
                ]
            },
            "collection_day_coverage": {
                "$round": [
                    {"$multiply": [
                        {"$divide": ["$collection_days_with_files", "$collection_days"]}, 100
                    ]}, 2
                ]
            }
        }
    },
    {
        "$project": {
            "_id": 0,
            "school_year": "$_id",
            "total_days": 1,
            "days_with_files": 1,
            "collection_days": 1,
            "collection_days_with_files": 1,
            "coverage_percentage": 1,
            "collection_day_coverage": 1
        }
    },
    {"$sort": {"school_year": 1}}
]

# =============================================================================
# DATA_CLEANING_SUMMARY
# =============================================================================
# Data cleaning summary for transparency
DATA_CLEANING_SUMMARY = [
    {
        "$facet": {
            "overall_summary": [
                {
                    "$group": {
                        "_id": None,
                        "total_files": {"$sum": 1},
                        "outliers": {"$sum": {"$cond": [{"$eq": ["$Outlier_Status", True]}, 1, 0]}},
                        "after_cleaning": {"$sum": {"$cond": [{"$eq": ["$Outlier_Status", False]}, 1, 0]}}
                    }
                }
            ],
            "by_school_year_and_type": [
                {
                    "$group": {
                        "_id": {
                            "school_year": "$School_Year",
                            "file_type": "$file_type"
                        },
                        "original_count": {"$sum": 1},
                        "outliers": {"$sum": {"$cond": [{"$eq": ["$Outlier_Status", True]}, 1, 0]}},
                        "after_cleaning": {"$sum": {"$cond": [{"$eq": ["$Outlier_Status", False]}, 1, 0]}}
                    }
                },
                {
                    "$addFields": {
                        "retention_percentage": {
                            "$round": [
                                {"$multiply": [
                                    {"$divide": ["$after_cleaning", "$original_count"]}, 100
                                ]}, 2
                            ]
                        }
                    }
                }
            ]
        }
    }
]

# Export all dashboard data pipelines
DASHBOARD_PIPELINES = {
    "DATA_QUALITY_CHECK": DATA_QUALITY_CHECK,
    "DASHBOARD_YEAR_SUMMARY": DASHBOARD_YEAR_SUMMARY,
    "DASHBOARD_PERIOD_SUMMARY": DASHBOARD_PERIOD_SUMMARY,
    "DASHBOARD_DATA_QUALITY": DASHBOARD_DATA_QUALITY,
    "DASHBOARD_COVERAGE": DASHBOARD_COVERAGE,
    "DATA_CLEANING_SUMMARY": DATA_CLEANING_SUMMARY
}
