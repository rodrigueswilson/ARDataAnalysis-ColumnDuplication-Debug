"""
This module contains all MongoDB aggregation pipelines for the AR Data Analysis project.

The pipelines are used to query and transform data from the media_records collection
to generate various data views for reports and analysis, including time-series counts,
summary statistics, and data quality metrics.

A central dictionary, `PIPELINES`, is defined at the end of this module, which maps
human-readable names to their corresponding pipeline definitions. This dictionary serves
as the primary interface for accessing pipelines from other parts of the application.
"""

# Import time series pipelines from modular structure
from pipelines.time_series import TIME_SERIES_PIPELINES

# Initialize the main dictionary to hold all pipelines
PIPELINES = {}



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
# 1b. DAILY_COUNTS_ALL_WITH_ZEROES (for ACF/PACF with complete time series)
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
# 2. WEEKLY_COUNTS
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
# 3. WEEKLY_COUNTS_WITH_ZEROES (for ACF/PACF)
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
# 3b. WEEKLY_COUNTS_FUTURE_PROOF (Simplified using ISO_YearWeek)
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

# =============================================================================
# 6. BIWEEKLY_COUNTS (for ACF/PACF)
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

# ================================
# DASHBOARD & OTHER PIPELINES
# ================================



# =============================================================================
# AUDIO_NOTE_CHARACTERISTICS
# =============================================================================
# Provides a detailed statistical analysis of audio note durations.
# - Filters for MP3 files with a valid duration.
# - Uses a $facet stage to calculate statistics both overall and per school year.
# - For each group, it calculates the mean, min, and max duration.
# - It also collects all durations into an array, which can be used for more
#   advanced statistical analysis (like calculating quartiles or standard deviation)
#   in the application layer.
# - This pipeline is key for understanding the nature of the audio data.

AUDIO_NOTE_CHARACTERISTICS = [
    {
        '$match': {
            'file_type': 'MP3',
            'Duration_Seconds': { '$exists': True, '$gt': 0 }
        }
    },
    {
        '$addFields': {
            'date_obj': { '$toDate': '$ISO_Date' }
        }
    },
    {
        '$addFields': {
            'school_year': {
                '$cond': {
                    'if': { '$gte': [{ '$month': '$date_obj' }, 8] },
                    'then': { '$concat': [ { '$toString': { '$year': '$date_obj' } }, '-', { '$toString': { '$add': [{ '$year': '$date_obj' }, 1] } } ] },
                    'else': { '$concat': [ { '$toString': { '$subtract': [{ '$year': '$date_obj' }, 1] } }, '-', { '$toString': { '$year': '$date_obj' } } ] }
                }
            }
        }
    },
    {
        '$facet': {
            'overall': [
                { 
                    '$group': {
                        '_id': 'overall',
                        'mean_duration': { '$avg': '$Duration_Seconds' },
                        'min_duration': { '$min': '$Duration_Seconds' },
                        'max_duration': { '$max': '$Duration_Seconds' },
                        'durations': { '$push': '$Duration_Seconds' }
                    }
                }
            ],
            'per_year': [
                {
                    '$group': {
                        '_id': '$school_year',
                        'mean_duration': { '$avg': '$Duration_Seconds' },
                        'min_duration': { '$min': '$Duration_Seconds' },
                        'max_duration': { '$max': '$Duration_Seconds' },
                        'durations': { '$push': '$Duration_Seconds' }
                    }
                }
            ]
        }
    }
]





CAMERA_USAGE_DATE_RANGE = [
    {
        "$match": {
            "file_type": "JPG"
        }
    },
    {
        "$group": {
            "_id": "$Camera_Model",
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"},
            "Total_Files": {"$sum": 1}
        }
    },
    {"$sort": {"First_Date": 1}}
]

# =============================================================================
# CAMERA_USAGE_BY_YEAR_RANGE
# =============================================================================
# Analyzes camera usage statistics, broken down by model and school year.
# - Filters for JPG files to focus on camera-generated data.
# - Groups records by Camera_Model and School_Year.
# - Calculates the first and last date of use and the total file count for each group.
# - This helps track the activity and lifespan of each camera across academic years.
CAMERA_USAGE_BY_YEAR_RANGE = [
    {
        "$match": {
            "file_type": "JPG"
        }
    },
    {
        "$group": {
            "_id": {
                "Camera_Model": "$Camera_Model",
                "School_Year": "$School_Year"
            },
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"},
            "Total_Files": {"$sum": 1}
        }
    },
    {
        "$sort": {
            "_id.Camera_Model": 1,
            "_id.School_Year": 1
        }
    },
    {
        "$project": {
            "_id": 0,
            "Camera_Model": "$_id.Camera_Model",
            "School_Year": "$_id.School_Year",
            "First_Date": 1,
            "Last_Date": 1,
            "Total_Files": 1
        }
    }
]

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
# 8. CAPTURE_VOLUME_PER_MONTH
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
# 9. TIME_OF_DAY_DISTRIBUTION
# =============================================================================
# Calculates the distribution of file captures across different times of the day.
# - Groups records by the 'Time_of_Day' field (e.g., 'Morning', 'Afternoon').
# - Sorts by count in descending order to show the most active periods.
# - Useful for understanding daily patterns of activity.
TIME_OF_DAY_DISTRIBUTION = [
    {
        "$group": {
            "_id": "$Time_of_Day",
            "Count": {"$sum": 1}
        }
    },
    {
        "$sort": {
            "Count": -1
        }
    }
]

# =============================================================================
# SCHEDULED_ACTIVITY_BREAKDOWN
# =============================================================================
# Counts the number of files associated with each type of scheduled activity.
# - Groups records by the 'Scheduled_Activity' field.
# - Sorts the results to show the most frequent activities first.
# - This pipeline is useful for understanding which scheduled events generate
#   the most data.
SCHEDULED_ACTIVITY_BREAKDOWN = [
    {
        "$group": {
            "_id": "$Scheduled_Activity",
            "Count": {"$sum": 1}
        }
    },
    {
        "$sort": {
            "Count": -1
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
# COLLECTION_PERIODS_DISTRIBUTION
# =============================================================================
# Calculates the frequency and total data volume for each collection period.
# - Filters out records not assigned to a collection period.
# - Groups by 'Collection_Period' to get the count of files and sum of file sizes.
# - Sorts by count to show the most active periods first.
# - This is useful for understanding the distribution of data collection efforts.
COLLECTION_PERIODS_DISTRIBUTION = [
    {
        "$match": {
            "Collection_Period": {"$ne": "N/A"}
        }
    },
    {
        "$group": {
            "_id": "$Collection_Period",
            "Count": {"$sum": 1},
            "SizeSum": {"$sum": "$File_Size_MB"}
        }
    },
    {
        "$project": {
            "_id": 0,
            "Period": "$_id",
            "Count": 1,
            "Total Size (MB)": {"$round": ["$SizeSum", 2]}
        }
    },
    {
        "$sort": {"Count": -1}
    }
]

# =============================================================================
# WEEKDAY_BY_PERIOD
# =============================================================================
# Analyzes the distribution of file captures across days of the week for each
# collection period.
# - Filters out records not assigned to a collection period.
# - Groups data by a composite key of Day_of_Week and Collection_Period.
# - This helps identify weekly patterns or biases in data collection within
#   specific periods.
WEEKDAY_BY_PERIOD = [
    {
        "$match": {
            "Collection_Period": {"$ne": "N/A"}
        }
    },
    {
        "$group": {
            "_id": {
                "day": "$Day_of_Week",
                "period": "$Collection_Period"
            },
            "total_files": {"$sum": 1}
        }
    },
    {
        "$sort": {
            "_id.day": 1,
            "_id.period": 1
        }
    },
    {
        "$project": {
            "_id": 0,
            "Day": "$_id.day",
            "Period": "$_id.period",
            "Total Files": "$total_files"
        }
    }
]

# =============================================================================
# SY_DAY_COUNT
# =============================================================================
# Verifies the total number of unique collection days within each school year.
# - Filters out records where 'Day_Number_in_SYCollection' is not available.
# - Groups by 'School_Year' and finds the maximum 'Day_Number_in_SYCollection'.
# - This provides a definitive count of active data collection days for each
#   school year, which is useful for validating other metrics.
SY_DAY_COUNT = [
    {
        '$match': {
            'Day_Number_in_SYCollection': {
                '$ne': 'N/A'
            }
        }
    }, 
    {
        '$group': {
            '_id': '$School_Year', 
            'Total_Collection_Days': {
                '$max': '$Day_Number_in_SYCollection'
            }
        }
    }, 
    {
        '$sort': {
            '_id': 1
        }
    }, 
    {
        '$project': {
            '_id': 0, 
            'School Year': '$_id', 
            'Calculated Collection Days': '$Total_Collection_Days'
        }
    }
]

# =============================================================================
# DATA_QUALITY_CHECK
# =============================================================================
# Identifies files that have not been assigned to a collection period.
# - Filters for records where 'Collection_Period' is 'N/A'.
# - Projects key fields like file path, date, and day of the week.
# - This pipeline is essential for data cleaning and ensuring all records are
#   properly categorized.
DATA_QUALITY_CHECK = [
    {
        "$match": {
            "Collection_Period": "N/A"
        }
    },
    {
        "$project": {
            "_id": 0,
            "File Path": "$file_path",
            "Date": "$ISO_Date",
            "Day of Week": "$Day_of_Week"
        }
    }
]

# ================================
# DASHBOARD PIPELINES
# ================================

# =============================================================================
# DASHBOARD_YEAR_SUMMARY
# =============================================================================
# Provides a year-over-year summary of key metrics for the main dashboard.
# - Groups all records by School_Year.
# - Calculates high-level statistics like total files, total size, file type counts,
#   outlier counts, and the number of unique days with data.
# - Computes derived metrics like outlier rate and average files per day.
# - This is a cornerstone pipeline for observing long-term trends.
DASHBOARD_YEAR_SUMMARY = [
    {
        "$group": {
            "_id": "$School_Year",
            "total_files": {"$sum": 1},
            "total_size_mb": {"$sum": "$File_Size_MB"},
            "jpg_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "mp3_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "outlier_files": {"$sum": {"$cond": ["$Outlier_Status", 1, 0]}},
            "unique_days": {"$addToSet": "$ISO_Date"},
            "avg_file_size_mb": {"$avg": "$File_Size_MB"}
        }
    },
    {
        "$project": {
            "_id": 0,
            "School_Year": "$_id",
            "Total_Files": "$total_files",
            "Total_Size_MB": {"$round": ["$total_size_mb", 2]},
            "JPG_Files": "$jpg_files",
            "MP3_Files": "$mp3_files",
            "Outlier_Files": "$outlier_files",
            "Outlier_Rate_Percent": {"$round": [{"$multiply": [{"$divide": ["$outlier_files", "$total_files"]}, 100]}, 2]},
            "Days_With_Data": {"$size": "$unique_days"},
            "Avg_File_Size_MB": {"$round": ["$avg_file_size_mb", 3]},
            "Avg_Files_Per_Day": {"$round": [{"$divide": ["$total_files", {"$size": "$unique_days"}]}, 1]}
        }
    },
    {"$sort": {"School_Year": 1}}
]

# =============================================================================
# DASHBOARD_PERIOD_SUMMARY
# =============================================================================
# Aggregates key statistics at the collection period level for dashboard display.
# - Filters out any records not assigned to a specific collection period.
# - Groups data by School_Year and Collection_Period.
# - Calculates total files, size, file type counts, and outlier metrics.
# - Computes derived metrics like outlier rate and average files per day.
# - This pipeline is critical for comparing data volume and quality across
#   different academic collection periods.
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
            "total_size_mb": {"$sum": "$File_Size_MB"},
            "jpg_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "mp3_files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "outlier_files": {"$sum": {"$cond": ["$Outlier_Status", 1, 0]}},
            "unique_days": {"$addToSet": "$ISO_Date"}
        }
    },
    {
        "$project": {
            "_id": 0,
            "School_Year": "$_id.school_year",
            "Period": "$_id.period",
            "Total_Files": "$total_files",
            "Total_Size_MB": {"$round": ["$total_size_mb", 2]},
            "JPG_Files": "$jpg_files",
            "MP3_Files": "$mp3_files",
            "Outlier_Files": "$outlier_files",
            "Outlier_Rate_Percent": {"$round": [{"$multiply": [{"$divide": ["$outlier_files", "$total_files"]}, 100]}, 2]},
            "Days_With_Files": {"$size": "$unique_days"},
            "Avg_Files_Per_Day": {"$round": [{"$divide": ["$total_files", {"$size": "$unique_days"}]}, 1]}
        }
    },
    {"$sort": {"School_Year": 1, "Period": 1}}
]

# =============================================================================
# DASHBOARD_DATA_QUALITY
# =============================================================================
# Provides a high-level overview of data quality issues, grouped by School_Year.
# - Groups records by School_Year.
# - Counts files with missing or invalid metadata (e.g., no ISO_Date, Activity, or Period).
# - Calculates the number of outlier files and zero-size files.
# - Computes the average file size for JPGs and MP3s separately.
# - This pipeline is crucial for identifying systemic data integrity problems.
DASHBOARD_DATA_QUALITY = [
    {
        "$group": {
            "_id": "$School_Year",
            "total_files": {"$sum": 1},
            "files_missing_metadata": {
                "$sum": {
                    "$cond": [
                        {
                            "$or": [
                                {"$eq": ["$ISO_Date", None]},
                                {"$eq": ["$Activity", "N/A"]},
                                {"$eq": ["$Collection_Period", "N/A"]}
                            ]
                        },
                        1, 0
                    ]
                }
            },
            "outlier_files": {"$sum": {"$cond": ["$Outlier_Status", 1, 0]}},
            "files_with_zero_size": {"$sum": {"$cond": [{"$eq": ["$File_Size_MB", 0]}, 1, 0]}},
            "jpg_avg_size": {
                "$avg": {
                    "$cond": [{"$eq": ["$file_type", "JPG"]}, "$File_Size_MB", None]
                }
            },
            "mp3_avg_size": {
                "$avg": {
                    "$cond": [{"$eq": ["$file_type", "MP3"]}, "$File_Size_MB", None]
                }
            }
        }
    },
    {
        "$project": {
            "_id": 0,
            "School_Year": "$_id",
            "Total_Files": "$total_files",
            "Missing_Metadata_Files": "$files_missing_metadata",
            "Missing_Metadata_Percent": {"$round": [{"$multiply": [{"$divide": ["$files_missing_metadata", "$total_files"]}, 100]}, 2]},
            "Outlier_Files": "$outlier_files",
            "Outlier_Rate_Percent": {"$round": [{"$multiply": [{"$divide": ["$outlier_files", "$total_files"]}, 100]}, 2]},
            "Zero_Size_Files": "$files_with_zero_size",
            "JPG_Avg_Size_MB": {"$round": ["$jpg_avg_size", 3]},
            "MP3_Avg_Size_MB": {"$round": ["$mp3_avg_size", 3]}
        }
    },
    {"$sort": {"School_Year": 1}}
]

# =============================================================================
# DASHBOARD_COVERAGE
# =============================================================================
# Assesses the completeness and consistency of data collection over time.
# - Groups records by School_Year.
# - Determines the total number of unique days with any data.
# - Calculates the average number of files collected per active day.
# - This pipeline helps answer questions about data collection intensity and identifies
#   potential gaps in the dataset.
DASHBOARD_COVERAGE = [
    {
        "$group": {
            "_id": "$School_Year",
            "unique_collection_days": {"$addToSet": "$ISO_Date"},
            "total_files": {"$sum": 1},
            "days_with_only_jpg": {
                "$addToSet": {
                    "$cond": [{"$eq": ["$file_type", "JPG"]}, "$ISO_Date", None]
                }
            },
            "days_with_only_mp3": {
                "$addToSet": {
                    "$cond": [{"$eq": ["$file_type", "MP3"]}, "$ISO_Date", None]
                }
            }
        }
    },
    {
        "$project": {
            "_id": 0,
            "School_Year": "$_id",
            "Days_With_Data": {"$size": "$unique_collection_days"},
            "Total_Files": "$total_files",
            "Avg_Files_Per_Day": {"$round": [{"$divide": ["$total_files", {"$size": "$unique_collection_days"}]}, 1]}
        }
    },
    {"$sort": {"School_Year": 1}}
]

# ================================
# MANUSCRIPT-FOCUSED ANALYSIS PIPELINES
# ================================

# =============================================================================
# DAY_OF_WEEK_COUNTS
# =============================================================================
# Aggregates file counts by day of the week to identify weekly patterns.
# - Groups records by Day_of_Week field.
# - Calculates total files, MP3 files, JPG files, and total size for each day.
# - Useful for understanding weekly activity patterns.
DAY_OF_WEEK_COUNTS = [
    {
        "$group": {
            "_id": "$Day_of_Week",
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"}
        }
    },
    {
        "$sort": {
            "_id": 1
        }
    }
]

# =============================================================================
# ACTIVITY_COUNTS
# =============================================================================
# Aggregates file counts by scheduled activity to understand activity patterns.
# - Groups records by Activity field.
# - Calculates total files, file types, and size for each activity.
# - Filters out N/A activities and sorts by count.
ACTIVITY_COUNTS = [
    {
        "$match": {
            "Activity": {"$ne": "N/A"}
        }
    },
    {
        "$group": {
            "_id": "$Activity",
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"}
        }
    },
    {
        "$sort": {
            "Total_Files": -1
        }
    }
]

# =============================================================================
# CAMERA_USAGE_BY_YEAR
# =============================================================================
# Aggregates camera usage statistics by school year.
# - Groups JPG files by Camera_Model and School_Year.
# - Calculates file counts and date ranges for each camera per year.
# - Useful for tracking camera utilization over time.
CAMERA_USAGE_BY_YEAR = [
    {
        "$match": {
            "file_type": "JPG"
        }
    },
    {
        "$group": {
            "_id": {
                "Camera_Model": "$Camera_Model",
                "School_Year": "$School_Year"
            },
            "Total_Files": {"$sum": 1},
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"}
        }
    },
    {
        "$project": {
            "_id": 0,
            "Camera_Model": "$_id.Camera_Model",
            "School_Year": "$_id.School_Year",
            "Total_Files": 1,
            "First_Date": 1,
            "Last_Date": 1
        }
    },
    {
        "$sort": {
            "School_Year": 1,
            "Camera_Model": 1
        }
    }
]

# =============================================================================
# CAMERA_USAGE_BY_YEAR_RANGE_DETAILED
# =============================================================================
# Detailed camera usage analysis with additional metrics.
# - Groups JPG files by Camera_Model and School_Year.
# - Includes detailed statistics like average files per day, unique days used.
# - Provides comprehensive camera utilization analysis.
CAMERA_USAGE_BY_YEAR_RANGE_DETAILED = [
    {
        "$match": {
            "file_type": "JPG"
        }
    },
    {
        "$group": {
            "_id": {
                "Camera_Model": "$Camera_Model",
                "School_Year": "$School_Year"
            },
            "Total_Files": {"$sum": 1},
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"},
            "Unique_Days": {"$addToSet": "$ISO_Date"},
            "Total_Size_MB": {"$sum": "$File_Size_MB"},
            "Avg_File_Size_MB": {"$avg": "$File_Size_MB"}
        }
    },
    {
        "$addFields": {
            "Days_Used": {"$size": "$Unique_Days"},
            "Avg_Files_Per_Day": {
                "$round": [
                    {"$divide": ["$Total_Files", {"$size": "$Unique_Days"}]}, 2
                ]
            }
        }
    },
    {
        "$project": {
            "_id": 0,
            "Camera_Model": "$_id.Camera_Model",
            "School_Year": "$_id.School_Year",
            "Total_Files": 1,
            "First_Date": 1,
            "Last_Date": 1,
            "Days_Used": 1,
            "Avg_Files_Per_Day": 1,
            "Total_Size_MB": {"$round": ["$Total_Size_MB", 2]},
            "Avg_File_Size_MB": {"$round": ["$Avg_File_Size_MB", 3]}
        }
    },
    {
        "$sort": {
            "School_Year": 1,
            "Total_Files": -1
        }
    }
]

# =============================================================================
# AUDIO_EFFICIENCY_ANALYSIS
# =============================================================================
# Analyzes the efficiency of audio data collection on a daily basis.
# - Filters for non-outlier MP3 files.
# - Groups by day to calculate total audio duration and count.
# - Uses a $lookup to get the total file count for the same day.
# - Calculates an efficiency metric: total files per minute of audio.
# - This helps assess how much field time is required to generate data.
AUDIO_EFFICIENCY_ANALYSIS = [
    # Match all non-outlier records to process them in a single pass.
    {"$match": {"Outlier_Status": False}},
    {
        # Group by day and calculate all metrics conditionally in one go.
        "$group": {
            "_id": "$ISO_Date",
            # Total files for the day (replaces the expensive lookup)
            "total_files_day": {"$sum": 1},
            # Conditionally sum MP3 counts and durations
            "mp3_count": {
                "$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}
            },
            "total_duration_seconds": {
                "$sum": {
                    "$cond": [{"$eq": ["$file_type", "MP3"]}, "$Duration_Seconds", 0]
                }
            },
            # Use None in the conditional to correctly calculate avg/min/max on MP3s only
            "avg_duration_seconds": {
                "$avg": {
                    "$cond": [{"$eq": ["$file_type", "MP3"]}, "$Duration_Seconds", None]
                }
            },
            "min_duration_seconds": {
                "$min": {
                    "$cond": [{"$eq": ["$file_type", "MP3"]}, "$Duration_Seconds", None]
                }
            },
            "max_duration_seconds": {
                "$max": {
                    "$cond": [{"$eq": ["$file_type", "MP3"]}, "$Duration_Seconds", None]
                }
            }
        }
    },
    # Filter out days that had no MP3 files after grouping.
    {"$match": {"mp3_count": {"$gt": 0}}},
    {
        # Add final calculated fields for efficiency metrics.
        "$addFields": {
            "total_duration_minutes": {"$divide": ["$total_duration_seconds", 60]},
            "efficiency_files_per_audio_minute": {
                "$cond": [
                    {"$gt": ["$total_duration_minutes", 0]},
                    {"$divide": ["$total_files_day", "$total_duration_minutes"]},
                    0  # Avoid division by zero
                ]
            }
        }
    },
    {"$sort": {"_id": 1}}
]

# =============================================================================
# DATA_CLEANING_SUMMARY
# =============================================================================
# Provides a summary of original file counts, outliers, and cleaned counts.
# - Groups by School_Year and file_type.
# - Filters out records where School_Year is "N/A".
# - Calculates the total number of original files.
# - Calculates the number of files flagged as outliers.
# This is used for the 'Data Cleaning' sheet to show sample composition.
DATA_CLEANING_SUMMARY = [
    {
        "$match": {
            "School_Year": {"$ne": "N/A"}
        }
    },
    {
        "$group": {
            "_id": {
                "school_year": "$School_Year",
                "file_type": "$file_type"
            },
            "Original": {"$sum": 1},
            "Outliers": {
                "$sum": {
                    "$cond": ["$Outlier_Status", 1, 0]
                }
            }
        }
    },
    {
        "$sort": {
            "_id.school_year": 1,
            "_id.file_type": 1
        }
    }
]

# =============================================================================
# MP3 DURATION ANALYSIS PIPELINES
# =============================================================================
# These pipelines analyze MP3 file durations across different time scales
# for comprehensive audio content analysis.

# MP3_DURATION_BY_SCHOOL_YEAR
# Aggregates MP3 duration metrics by school year
MP3_DURATION_BY_SCHOOL_YEAR = [
    {
        "$match": {
            "file_type": "MP3",
            "School_Year": {"$ne": "N/A"},
            "Duration_Seconds": {"$exists": True, "$gt": 0}
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
            "Days_With_MP3": {"$size": "$Unique_Days"},
            "Avg_Files_Per_Day": {
                "$round": [
                    {"$divide": ["$Total_MP3_Files", {"$size": "$Unique_Days"}]}, 1
                ]
            },
            "Total_Duration_Hours": {
                "$round": [
                    {"$divide": ["$Total_Duration_Seconds", 3600]}, 2
                ]
            }
        }
    },
    {"$sort": {"_id": 1}}
]

# MP3_DURATION_BY_PERIOD
# Aggregates MP3 duration metrics by collection period
MP3_DURATION_BY_PERIOD = [
    {
        "$match": {
            "file_type": "MP3",
            "Collection_Period": {"$ne": "N/A"},
            "Duration_Seconds": {"$exists": True, "$gt": 0}
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
            "Min_Duration_Seconds": {"$min": "$Duration_Seconds"},
            "Max_Duration_Seconds": {"$max": "$Duration_Seconds"},
            "Unique_Days": {"$addToSet": "$ISO_Date"}
        }
    },
    {
        "$addFields": {
            "Days_With_MP3": {"$size": "$Unique_Days"},
            "Avg_Files_Per_Day": {
                "$round": [
                    {"$divide": ["$Total_MP3_Files", {"$size": "$Unique_Days"}]}, 1
                ]
            },
            "Total_Duration_Hours": {
                "$round": [
                    {"$divide": ["$Total_Duration_Seconds", 3600]}, 2
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
        "$sort": {
            "_id.School_Year": 1,
            "_id.Period": 1
        }
    }
]

# MP3_DURATION_BY_MONTH
# Aggregates MP3 duration metrics by month with school year comparison
MP3_DURATION_BY_MONTH = [
    {
        "$match": {
            "file_type": "MP3",
            "School_Year": {"$ne": "N/A"},
            "Duration_Seconds": {"$exists": True, "$gt": 0}
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



# ==============================================================================
# FINAL PIPELINE REGISTRATION
# ==============================================================================
# This final update call ensures all pipelines defined in this module are
# registered in the central PIPELINES dictionary. The dictionary is initialized
# at the top of the file and populated here to ensure all definitions are captured.

# First, add all time series pipelines from the modular structure
PIPELINES.update(TIME_SERIES_PIPELINES)

# Then add the remaining pipelines
PIPELINES.update({
    # Legacy time-series pipelines (keeping for compatibility)
    "DAILY_COUNTS_ALL": DAILY_COUNTS_ALL,
    "DAILY_COUNTS_ALL_WITH_ZEROES": DAILY_COUNTS_ALL_WITH_ZEROES,
    "WEEKLY_COUNTS": WEEKLY_COUNTS,

    # Activity and pattern analysis pipelines
    "DAY_OF_WEEK_COUNTS": DAY_OF_WEEK_COUNTS,
    "ACTIVITY_COUNTS": ACTIVITY_COUNTS,
    "CAMERA_USAGE_BY_YEAR": CAMERA_USAGE_BY_YEAR,
    "CAMERA_USAGE_BY_YEAR_RANGE_DETAILED": CAMERA_USAGE_BY_YEAR_RANGE_DETAILED,

    # Detailed analysis pipelines
    "AUDIO_NOTE_CHARACTERISTICS": AUDIO_NOTE_CHARACTERISTICS,
    "CAMERA_USAGE_DATE_RANGE": CAMERA_USAGE_DATE_RANGE,
    "CAMERA_USAGE_BY_YEAR_RANGE": CAMERA_USAGE_BY_YEAR_RANGE,
    "FILE_SIZE_SUMMARY_BY_DAY": FILE_SIZE_SUMMARY_BY_DAY,
    "CAPTURE_VOLUME_PER_MONTH": CAPTURE_VOLUME_PER_MONTH,
    "TIME_OF_DAY_DISTRIBUTION": TIME_OF_DAY_DISTRIBUTION,
    "SCHEDULED_ACTIVITY_BREAKDOWN": SCHEDULED_ACTIVITY_BREAKDOWN,
    "FILE_SIZE_STATS": FILE_SIZE_STATS,
    "COLLECTION_PERIODS_DISTRIBUTION": COLLECTION_PERIODS_DISTRIBUTION,
    "WEEKDAY_BY_PERIOD": WEEKDAY_BY_PERIOD,
    "SY_DAY_COUNT": SY_DAY_COUNT,
    "AUDIO_EFFICIENCY_ANALYSIS": AUDIO_EFFICIENCY_ANALYSIS,

    # Data quality and dashboard pipelines
    "DATA_QUALITY_CHECK": DATA_QUALITY_CHECK,
    "DASHBOARD_YEAR_SUMMARY": DASHBOARD_YEAR_SUMMARY,
    "DASHBOARD_PERIOD_SUMMARY": DASHBOARD_PERIOD_SUMMARY,
    "DASHBOARD_DATA_QUALITY": DASHBOARD_DATA_QUALITY,
    "DASHBOARD_COVERAGE": DASHBOARD_COVERAGE,
    "DATA_CLEANING_SUMMARY": DATA_CLEANING_SUMMARY,

    # MP3 Duration Analysis pipelines
    "MP3_DURATION_BY_SCHOOL_YEAR": MP3_DURATION_BY_SCHOOL_YEAR,
    "MP3_DURATION_BY_PERIOD": MP3_DURATION_BY_PERIOD,
    "MP3_DURATION_BY_MONTH": MP3_DURATION_BY_MONTH
})
