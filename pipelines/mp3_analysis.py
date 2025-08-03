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
# MP3 DURATION ANALYSIS PIPELINES
# =============================================================================
# Core match criteria to ensure consistent filtering
# Using direct match structure rather than filters to ensure compatibility
CORE_MATCH = {
    "file_type": "MP3",
    "School_Year": {"$ne": "N/A"},
    "Duration_Seconds": {"$exists": True, "$gt": 0},
    "is_collection_day": True,
    "Outlier_Status": False
}

# =============================================================================
# MP3_DURATION_BY_SCHOOL_YEAR
# =============================================================================
# Aggregates MP3 duration metrics by school year with comprehensive statistics
MP3_DURATION_BY_SCHOOL_YEAR = [
    {
        "$match": CORE_MATCH
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
            "Total_Duration_Hours": {"$divide": ["$Total_Duration_Seconds", 3600]},
            "Days_With_MP3": {"$size": "$Unique_Days"},
            "Avg_Files_Per_Day": {"$cond": [
                {"$eq": [{"$size": "$Unique_Days"}, 0]},
                0,
                {"$divide": ["$Total_MP3_Files", {"$size": "$Unique_Days"}]}
            ]}
        }
    },
    {
        "$project": {
            "School_Year": "$_id",
            "Total_MP3_Files": 1,
            "Total_Duration_Hours": 1,
            "Total_Duration_Seconds": 1,
            "Avg_Duration_Seconds": 1,
            "Min_Duration_Seconds": 1,
            "Max_Duration_Seconds": 1,
            "Days_With_MP3": 1,
            "Avg_Files_Per_Day": 1,
            "_id": 0
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
        "$match": CORE_MATCH
    },
    {
        "$match": {
            "Collection_Period": {
                "$exists": True,
                "$ne": None,
                "$ne": "N/A",
                "$ne": "",
                "$nin": ["Unknown", "unknown", "UNKNOWN"]
            }
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
            "Total_Duration_Hours": {"$divide": ["$Total_Duration_Seconds", 3600]},
            "Days_With_MP3": {"$size": "$Unique_Days"},
            "Avg_Files_Per_Day": {"$cond": [
                {"$eq": [{"$size": "$Unique_Days"}, 0]},
                0,
                {"$divide": ["$Total_MP3_Files", {"$size": "$Unique_Days"}]}
            ]},
            "Period_Efficiency": {"$cond": [
                {"$eq": [{"$divide": ["$Total_Duration_Seconds", 3600]}, 0]},
                0,
                {"$divide": ["$Total_MP3_Files", {"$divide": ["$Total_Duration_Seconds", 3600]}]}
            ]}
        }
    },
    {
        "$project": {
            "School_Year": "$_id.School_Year",
            "Period": "$_id.Period",
            "Total_MP3_Files": 1,
            "Total_Duration_Hours": 1,
            "Total_Duration_Seconds": 1,
            "Avg_Duration_Seconds": 1,
            "Days_With_MP3": 1,
            "Avg_Files_Per_Day": 1,
            "Period_Efficiency": 1,
            "_id": 0
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
        "$match": CORE_MATCH
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
            "Total_Duration_Hours": {"$divide": ["$Total_Duration_Seconds", 3600]}
        }
    },
    {
        "$project": {
            "Month": "$_id.Month",
            "School_Year": "$_id.School_Year", 
            "Total_MP3_Files": 1,
            "Total_Duration_Hours": 1,
            "Total_Duration_Seconds": 1,
            "Avg_Duration_Seconds": 1,
            "_id": 0
        }
    },
    {
        "$sort": {
            "Month": 1,
            "School_Year": 1
        }
    }
]

# Export all MP3 analysis pipelines
MP3_PIPELINES = {
    "MP3_DURATION_BY_SCHOOL_YEAR": MP3_DURATION_BY_SCHOOL_YEAR,
    "MP3_DURATION_BY_PERIOD": MP3_DURATION_BY_PERIOD,
    "MP3_DURATION_BY_MONTH": MP3_DURATION_BY_MONTH
}

# Document the purpose of the pipelines
"""
MP3 Duration Analysis Pipeline Usage:

1. MP3_DURATION_BY_SCHOOL_YEAR:
   - Provides a comprehensive summary of MP3 files and durations by school year
   - Used in the MP3 Duration Analysis sheet's School Year MP3 Duration Summary table
   - Includes fields for total files, duration metrics, and activity patterns

2. MP3_DURATION_BY_PERIOD:
   - Breaks down MP3 data by collection period and school year
   - Used in the MP3 Duration Analysis sheet's Collection Period Duration table
   - Includes period efficiency metrics (files per hour of recording)

3. MP3_DURATION_BY_MONTH:
   - Provides month-by-month MP3 metrics across school years
   - Used in the MP3 Duration Analysis sheet's Monthly MP3 Duration Distribution table
   - Enables year-over-year monthly comparisons

All pipelines consistently apply data cleaning filters:
- Collection days only (is_collection_day: True)
- Non-outliers only (Outlier_Status: False)
- Valid school years (School_Year ≠ "N/A")
- Valid MP3 durations (Duration_Seconds > 0)
"""

