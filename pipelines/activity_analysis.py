"""
Activity Analysis Pipelines for AR Data Analysis

This module contains MongoDB aggregation pipelines for analyzing activity patterns,
time distributions, and scheduled activities in the data collection.

Key Pipelines:
- DAY_OF_WEEK_COUNTS: File counts by day of week
- ACTIVITY_COUNTS: Counts by scheduled activity type
- TIME_OF_DAY_DISTRIBUTION: Distribution across time periods
- SCHEDULED_ACTIVITY_BREAKDOWN: Breakdown by scheduled activities
- COLLECTION_PERIODS_DISTRIBUTION: Distribution across collection periods
- WEEKDAY_BY_PERIOD: Weekday analysis by collection period
- SY_DAY_COUNT: School year day counts
"""

# =============================================================================
# DAY_OF_WEEK_COUNTS
# =============================================================================
# Aggregates file counts by day of the week
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
    {"$sort": {"_id": 1}}
]

# =============================================================================
# ACTIVITY_COUNTS
# =============================================================================
# Aggregates file counts by scheduled activity
ACTIVITY_COUNTS = [
    {
        "$group": {
            "_id": "$Scheduled_Activity",
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"}
        }
    },
    {"$sort": {"Total_Files": -1}}
]

# =============================================================================
# TIME_OF_DAY_DISTRIBUTION
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
WEEKDAY_BY_PERIOD = [
    {
        "$match": {
            "Collection_Period": {"$ne": "N/A"}
        }
    },
    {
        "$group": {
            "_id": {
                "Period": "$Collection_Period",
                "Day_of_Week": "$Day_of_Week"
            },
            "Count": {"$sum": 1}
        }
    },
    {
        "$project": {
            "_id": 0,
            "Period": "$_id.Period",
            "Day_of_Week": "$_id.Day_of_Week",
            "Count": 1
        }
    },
    {
        "$sort": {
            "Period": 1,
            "Day_of_Week": 1
        }
    }
]

# =============================================================================
# SY_DAY_COUNT
# =============================================================================
# Counts files by school year and day
SY_DAY_COUNT = [
    {
        "$group": {
            "_id": {
                "School_Year": "$School_Year",
                "ISO_Date": "$ISO_Date"
            },
            "Total_Files": {"$sum": 1},
            "MP3_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "MP3"]}, 1, 0]}},
            "JPG_Files": {"$sum": {"$cond": [{"$eq": ["$file_type", "JPG"]}, 1, 0]}},
            "Total_Size_MB": {"$sum": "$File_Size_MB"}
        }
    },
    {
        "$project": {
            "_id": 0,
            "School_Year": "$_id.School_Year",
            "ISO_Date": "$_id.ISO_Date",
            "Total_Files": 1,
            "MP3_Files": 1,
            "JPG_Files": 1,
            "Total_Size_MB": 1
        }
    },
    {
        "$sort": {
            "School_Year": 1,
            "ISO_Date": 1
        }
    }
]

# Export all activity analysis pipelines
ACTIVITY_PIPELINES = {
    "DAY_OF_WEEK_COUNTS": DAY_OF_WEEK_COUNTS,
    "ACTIVITY_COUNTS": ACTIVITY_COUNTS,
    "TIME_OF_DAY_DISTRIBUTION": TIME_OF_DAY_DISTRIBUTION,
    "SCHEDULED_ACTIVITY_BREAKDOWN": SCHEDULED_ACTIVITY_BREAKDOWN,
    "COLLECTION_PERIODS_DISTRIBUTION": COLLECTION_PERIODS_DISTRIBUTION,
    "WEEKDAY_BY_PERIOD": WEEKDAY_BY_PERIOD,
    "SY_DAY_COUNT": SY_DAY_COUNT
}
