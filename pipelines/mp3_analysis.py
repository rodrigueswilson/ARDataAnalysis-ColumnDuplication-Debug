"""
MP3 Analysis Pipelines for AR Data Analysis

This module contains MongoDB aggregation pipelines specifically for analyzing
MP3 audio files, including duration analysis, efficiency metrics, and temporal patterns.

Key Pipelines:
- MP3_DURATION_BY_SCHOOL_YEAR: Duration analysis by school year
- MP3_DURATION_BY_PERIOD: Duration analysis by collection period
- MP3_DURATION_BY_MONTH: Monthly duration analysis with year-over-year comparison
"""

# =============================================================================
# MP3_DURATION_BY_SCHOOL_YEAR
# =============================================================================
# Aggregates MP3 duration metrics by school year with comprehensive statistics
MP3_DURATION_BY_SCHOOL_YEAR = [
    {
        "$match": {
            "file_type": "MP3",
            "School_Year": {"$ne": "N/A"},
            "Duration_Seconds": {"$exists": True, "$gt": 0},
            "is_collection_day": True,
            "Outlier_Status": False
        }
    },
    {
        "$group": {
            "_id": "$School_Year",
            "Total_MP3_Files": {"$sum": 1},
            "Total_Duration_Seconds": {"$sum": "$Duration_Seconds"},
            "Avg_Duration_Seconds": {"$avg": "$Duration_Seconds"},
            "Min_Duration_Seconds": {"$min": "$Duration_Seconds"},
            "Max_Duration_Seconds": {"$max": "$Duration_Seconds"},
            "Unique_Days": {"$addToSet": "$ISO_Date"}
        }
    },
    {
        "$addFields": {
            "Total_Duration_Hours": {
                "$round": [
                    {"$divide": ["$Total_Duration_Seconds", 3600]}, 2
                ]
            },
            "Days_With_MP3": {"$size": "$Unique_Days"},
            "Avg_Files_Per_Day": {
                "$round": [
                    {"$divide": ["$Total_MP3_Files", {"$size": "$Unique_Days"}]}, 2
                ]
            }
        }
    },
    {
        "$project": {
            "_id": 0,
            "School_Year": "$_id",
            "Total_MP3_Files": 1,
            "Total_Duration_Hours": 1,
            "Avg_Duration_Seconds": {"$round": ["$Avg_Duration_Seconds", 1]},
            "Min_Duration_Seconds": 1,
            "Max_Duration_Seconds": 1,
            "Days_With_MP3": 1,
            "Avg_Files_Per_Day": 1
        }
    },
    {
        "$sort": {
            "School_Year": 1
        }
    }
]

# =============================================================================
# MP3_DURATION_BY_PERIOD
# =============================================================================
# Aggregates MP3 duration metrics by collection period with efficiency calculations
MP3_DURATION_BY_PERIOD = [
    {
        "$match": {
            "file_type": "MP3",
            "Collection_Period": {"$ne": "N/A"},
            "Duration_Seconds": {"$exists": True, "$gt": 0},
            "is_collection_day": True,
            "Outlier_Status": False
        }
    },
    {
        "$group": {
            "_id": {
                "School_Year": "$School_Year",
                "Period": "$Collection_Period"
            },
            "Total_MP3_Files": {"$sum": 1},
            "Total_Duration_Seconds": {"$sum": "$Duration_Seconds"},
            "Avg_Duration_Seconds": {"$avg": "$Duration_Seconds"},
            "Unique_Days": {"$addToSet": "$ISO_Date"}
        }
    },
    {
        "$addFields": {
            "Total_Duration_Hours": {
                "$round": [
                    {"$divide": ["$Total_Duration_Seconds", 3600]}, 2
                ]
            },
            "Days_With_MP3": {"$size": "$Unique_Days"},
            "Avg_Files_Per_Day": {
                "$round": [
                    {"$divide": ["$Total_MP3_Files", {"$size": "$Unique_Days"}]}, 2
                ]
            },
            "Period_Efficiency": {
                "$round": [
                    {"$divide": ["$Total_MP3_Files", "$Total_Duration_Hours"]}, 2
                ]
            }
        }
    },
    {
        "$project": {
            "_id": 0,
            "School_Year": "$_id.School_Year",
            "Period": "$_id.Period",
            "Total_MP3_Files": 1,
            "Total_Duration_Hours": 1,
            "Avg_Duration_Seconds": {"$round": ["$Avg_Duration_Seconds", 1]},
            "Days_With_MP3": 1,
            "Avg_Files_Per_Day": 1,
            "Period_Efficiency": 1
        }
    },
    {
        "$sort": {
            "School_Year": 1,
            "Period": 1
        }
    }
]

# =============================================================================
# MP3_DURATION_BY_MONTH
# =============================================================================
# Aggregates MP3 duration metrics by month with school year comparison
MP3_DURATION_BY_MONTH = [
    {
        "$match": {
            "file_type": "MP3",
            "School_Year": {"$ne": "N/A"},
            "Duration_Seconds": {"$exists": True, "$gt": 0},
            "is_collection_day": True,
            "Outlier_Status": False
        }
    },
    {
        "$addFields": {
            "Month": {
                "$month": {
                    "$dateFromString": {
                        "dateString": "$ISO_Date"
                    }
                }
            }
        }
    },
    {
        "$group": {
            "_id": {
                "Month": "$Month",
                "School_Year": "$School_Year"
            },
            "Total_MP3_Files": {"$sum": 1},
            "Total_Duration_Seconds": {"$sum": "$Duration_Seconds"},
            "Avg_Duration_Seconds": {"$avg": "$Duration_Seconds"}
        }
    },
    {
        "$addFields": {
            "Total_Duration_Hours": {
                "$round": [
                    {"$divide": ["$Total_Duration_Seconds", 3600]}, 2
                ]
            }
        }
    },
    {
        "$sort": {
            "_id.Month": 1,
            "_id.School_Year": 1
        }
    }
]

# Export all MP3 analysis pipelines
MP3_PIPELINES = {
    "MP3_DURATION_BY_SCHOOL_YEAR": MP3_DURATION_BY_SCHOOL_YEAR,
    "MP3_DURATION_BY_PERIOD": MP3_DURATION_BY_PERIOD,
    "MP3_DURATION_BY_MONTH": MP3_DURATION_BY_MONTH
}
