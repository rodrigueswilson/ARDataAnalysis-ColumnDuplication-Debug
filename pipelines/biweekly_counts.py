"""
Biweekly Count Pipelines for AR Data Analysis

This module contains the BIWEEKLY_COUNTS pipeline for intermediate time series analysis.
This provides a middle ground between weekly and monthly analysis.
"""

# =============================================================================
# BIWEEKLY_COUNTS (for ACF/PACF)
# =============================================================================
# Aggregates file counts every two weeks.
# - Calculates the bi-week number within the year.
# - Groups by year and bi-week number.
# - Creates a descriptive label (e.g., '2022-B1', '2022-B2').
# - Useful for finding a middle ground between weekly and monthly analysis,
#   potentially revealing different cyclical patterns.
BIWEEKLY_COUNTS = [
    {
        "$addFields": {
            "ISO_Year": {
                "$year": {
                    "$dateFromString": {
                        "dateString": "$ISO_Date"
                    }
                }
            },
            "Biweek_Number": {
                "$floor": {
                    "$divide": [
                        {"$subtract": ["$ISO_Week", 1]}, 2
                    ]
                }
            }
        }
    },
    {
        "$group": {
            "_id": {
                "Year": "$ISO_Year",
                "Biweek": "$Biweek_Number"
            },
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"},
            "Unique_Days": {"$addToSet": "$ISO_Date"},
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"},
            "School_Year": {"$first": "$School_Year"}
        }
    },
    {
        "$addFields": {
            "Period_Label": {
                "$concat": [
                    {"$toString": "$_id.Year"},
                    "-B",
                    {"$toString": {"$add": ["$_id.Biweek", 1]}}
                ]
            },
            "Days_With_Data": {"$size": "$Unique_Days"}
        }
    },
    {
        "$project": {
            "_id": "$Period_Label",
            "Year": "$_id.Year",
            "Biweek_Number": "$_id.Biweek",
            "School_Year": 1,
            "First_Date": 1,
            "Last_Date": 1,
            "Total_Files": 1,
            "JPG_Files": 1,
            "MP3_Files": 1,
            "Days_With_Data": 1,
            "Total_Size_MB": 1
        }
    },
    {"$sort": {"Year": 1, "Biweek_Number": 1}}
]

# Export biweekly pipelines
BIWEEKLY_PIPELINES = {
    "BIWEEKLY_COUNTS": BIWEEKLY_COUNTS
}
