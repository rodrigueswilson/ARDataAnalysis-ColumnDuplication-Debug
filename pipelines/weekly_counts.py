"""
Weekly Count Pipelines for AR Data Analysis

This module contains MongoDB aggregation pipelines for weekly-level analysis.
These pipelines aggregate file counts, sizes, and metadata by week (ISO_Week).

Key Pipelines:
- WEEKLY_COUNTS: Basic weekly aggregation
- WEEKLY_COUNTS_WITH_ZEROES: Weekly aggregation for time series analysis
- WEEKLY_COUNTS_FUTURE_PROOF: Simplified version using pre-calculated fields
- MONTHLY_COUNTS_WITH_ZEROES: Monthly aggregation for time series
- PERIOD_COUNTS_WITH_ZEROES: Academic period aggregation
- BIWEEKLY_COUNTS: Bi-weekly aggregation for intermediate analysis
"""

# =============================================================================
# 1. WEEKLY_COUNTS
# =============================================================================
# Aggregates file counts and total size per week.
# - Groups records by ISO_Week.
# - Calculates total files, MP3 files, JPG files, and total size in MB.
# - Sorts the results by week number.
WEEKLY_COUNTS = [
    {
        "$group": {
            "_id": "$ISO_Week",
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"}
        }
    },
    {"$sort": {"_id": 1}}
]

# =============================================================================
# 2. WEEKLY_COUNTS_WITH_ZEROES (for ACF/PACF)
# =============================================================================
# Aggregates file counts per week, structured for time-series analysis.
# - Extracts year and week from ISO_Date for accurate grouping.
# - Creates a composite key (_id) of Year and Week.
# - Generates a user-friendly 'YYYY-Www' label for charting.
# - This pipeline is essential for ACF/PACF analysis, which requires a
#   continuous time index. The 'WITH_ZEROES' implies that downstream
#   processing will fill in any missing weeks with zero counts.
WEEKLY_COUNTS_WITH_ZEROES = [
    {
        "$addFields": {
            "ISO_Year": {
                "$year": {
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
                "Year": "$ISO_Year",
                "Week": "$ISO_Week"
            },
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"},
            "School_Year": {"$first": "$School_Year"},
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"}
        }
    },
    {
        "$addFields": {
            "Year_Week_Label": {
                "$concat": [
                    {"$toString": "$_id.Year"},
                    "-W",
                    {"$toString": "$_id.Week"}
                ]
            }
        }
    },
    {
        "$project": {
            "_id": "$Year_Week_Label",
            "Year": "$_id.Year",
            "Week": "$_id.Week",
            "School_Year": 1,
            "First_Date": 1,
            "Last_Date": 1,
            "Total_Files": 1,
            "JPG_Files": 1,
            "MP3_Files": 1,
            "Total_Size_MB": 1
        }
    },
    {"$sort": {"Year": 1, "Week": 1}}
]

# =============================================================================
# 3. WEEKLY_COUNTS_FUTURE_PROOF (Simplified using ISO_YearWeek)
# =============================================================================
# A simplified, more efficient version of the weekly counts pipeline.
# - Relies on a pre-calculated 'ISO_YearWeek' field in the source data.
# - This avoids complex date manipulation within the aggregation pipeline.
# - NOTE: This pipeline is aspirational and requires a data migration to be useful.
WEEKLY_COUNTS_FUTURE_PROOF = [
    {
        "$group": {
            "_id": "$ISO_YearWeek",
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"},
            "Year": {"$first": "$ISO_Year"},
            "Week": {"$first": "$ISO_Week"},
            "School_Year": {"$first": "$School_Year"},
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"}
        }
    },
    {"$sort": {"Year": 1, "Week": 1}}
]

# =============================================================================
# 4. MONTHLY_COUNTS_WITH_ZEROES (for ACF/PACF)
# =============================================================================
# Aggregates file counts per month, structured for time-series analysis.
# - Extracts a 'YYYY-MM' string from ISO_Date for grouping.
# - Calculates total files, file types, size, and days with data.
# - Similar to the weekly version, this is designed for time-series analysis
#   where missing months will be filled with zeroes downstream.
MONTHLY_COUNTS_WITH_ZEROES = [
    {
        "$addFields": {
            "Year_Month": {
                "$dateToString": {
                    "format": "%Y-%m",
                    "date": {
                        "$dateFromString": {
                            "dateString": "$ISO_Date"
                        }
                    }
                }
            }
        }
    },
    {
        "$group": {
            "_id": "$Year_Month",
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"},
            "Unique_Days": {"$addToSet": "$ISO_Date"},
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"}
        }
    },
    {
        "$addFields": {
            "Days_With_Data": {"$size": "$Unique_Days"},
            "Avg_Files_Per_Day": {
                "$round": [
                    {"$divide": ["$Total_Files", {"$size": "$Unique_Days"}]}, 2
                ]
            }
        }
    },
    {
        "$project": {
            "_id": 1,
            "First_Date": 1,
            "Last_Date": 1,
            "Total_Files": 1,
            "JPG_Files": 1,
            "MP3_Files": 1,
            "Days_With_Data": 1,
            "Avg_Files_Per_Day": 1,
            "Total_Size_MB": 1
        }
    },
    {"$sort": {"_id": 1}}
]

# =============================================================================
# 5. PERIOD_COUNTS_WITH_ZEROES (for ACF/PACF)
# =============================================================================
# Aggregates file counts by academic collection period (e.g., 'SY 21-22 P1').
# - Filters out records where 'Collection_Period' is 'N/A'.
# - Groups by School_Year and Period.
# - Calculates a wide range of statistics, including file counts, size,
#   unique days, and average files per day.
# - Uses a $switch statement to map periods to fixed calendar start/end dates,
#   ensuring consistent period lengths for comparison.
# - Essential for analyzing data volume across discrete, non-standard time intervals.
PERIOD_COUNTS_WITH_ZEROES = [
    {
        "$match": {
            "Collection_Period": {"$ne": "N/A"}
        }
    },
    {
        "$group": {
            "_id": {
                "School_Year": "$School_Year",
                "Period": "$Collection_Period"
            },
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"},
            "Unique_Days": {"$addToSet": "$ISO_Date"},
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"}
        }
    },
    {
        "$addFields": {
            "Days_With_Data": {"$size": "$Unique_Days"},
            "Avg_Files_Per_Day": {
                "$round": [
                    {"$divide": ["$Total_Files", {"$size": "$Unique_Days"}]}, 2
                ]
            },
            # Replace actual first date with calendar start date based on period
            "Calendar_First_Date": {
                "$switch": {
                    "branches": [
                        {"case": {"$eq": ["$_id.Period", "SY 21-22 P1"]}, "then": "2021-09-13"},
                        {"case": {"$eq": ["$_id.Period", "SY 21-22 P2"]}, "then": "2021-11-08"},
                        {"case": {"$eq": ["$_id.Period", "SY 21-22 P3"]}, "then": "2022-02-22"},
                        {"case": {"$eq": ["$_id.Period", "SY 22-23 P1"]}, "then": "2022-09-12"},
                        {"case": {"$eq": ["$_id.Period", "SY 22-23 P2"]}, "then": "2022-12-13"},
                        {"case": {"$eq": ["$_id.Period", "SY 22-23 P3"]}, "then": "2023-03-16"}
                    ],
                    "default": "$First_Date"  # Fallback to actual date if period not recognized
                }
            },
            # Replace actual last date with calendar end date based on period
            "Calendar_Last_Date": {
                "$switch": {
                    "branches": [
                        {"case": {"$eq": ["$_id.Period", "SY 21-22 P1"]}, "then": "2021-11-05"},
                        {"case": {"$eq": ["$_id.Period", "SY 21-22 P2"]}, "then": "2022-02-18"},
                        {"case": {"$eq": ["$_id.Period", "SY 21-22 P3"]}, "then": "2022-05-27"},
                        {"case": {"$eq": ["$_id.Period", "SY 22-23 P1"]}, "then": "2022-12-12"},
                        {"case": {"$eq": ["$_id.Period", "SY 22-23 P2"]}, "then": "2023-03-15"},
                        {"case": {"$eq": ["$_id.Period", "SY 22-23 P3"]}, "then": "2023-06-01"}
                    ],
                    "default": "$Last_Date"  # Fallback to actual date if period not recognized
                }
            }
        }
    },
    {
        "$project": {
            "_id": 0,
            "School_Year": "$_id.School_Year",
            "Period": "$_id.Period",
            "First_Date": "$Calendar_First_Date",  # Use calendar start date instead of actual first file date
            "Last_Date": "$Calendar_Last_Date",  # Use calendar end date instead of actual last file date
            "Total_Files": 1,
            "JPG_Files": 1,
            "MP3_Files": 1,
            "Days_With_Data": 1,
            "Avg_Files_Per_Day": 1,
            "Total_Size_MB": 1
        }
    },
    {"$sort": {"School_Year": 1, "Period": 1}}
]

# Export all weekly-related pipelines
WEEKLY_PIPELINES = {
    "WEEKLY_COUNTS": WEEKLY_COUNTS,
    "WEEKLY_COUNTS_WITH_ZEROES": WEEKLY_COUNTS_WITH_ZEROES,
    "WEEKLY_COUNTS_FUTURE_PROOF": WEEKLY_COUNTS_FUTURE_PROOF,
    "MONTHLY_COUNTS_WITH_ZEROES": MONTHLY_COUNTS_WITH_ZEROES,
    "PERIOD_COUNTS_WITH_ZEROES": PERIOD_COUNTS_WITH_ZEROES
}
