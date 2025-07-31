"""
Camera Usage Analysis Pipelines for AR Data Analysis

This module contains MongoDB aggregation pipelines for analyzing camera usage patterns,
including camera models, usage by year, and date ranges.

Key Pipelines:
- CAMERA_USAGE_BY_YEAR: Camera usage statistics by year
- CAMERA_USAGE_BY_YEAR_RANGE_DETAILED: Detailed camera usage with date ranges
- CAMERA_USAGE_DATE_RANGE: Camera usage date ranges
- CAMERA_USAGE_BY_YEAR_RANGE: Camera usage by year with ranges
- AUDIO_NOTE_CHARACTERISTICS: Audio file characteristics analysis
"""

# =============================================================================
# CAMERA_USAGE_BY_YEAR
# =============================================================================
# Analyzes camera usage by school year
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
            "Total_Size_MB": {"$sum": "$File_Size_MB"},
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
            "Total_Size_MB": 1,
            "First_Date": 1,
            "Last_Date": 1
        }
    },
    {
        "$sort": {
            "Camera_Model": 1,
            "School_Year": 1
        }
    }
]

# =============================================================================
# CAMERA_USAGE_BY_YEAR_RANGE_DETAILED
# =============================================================================
# Detailed camera usage analysis with comprehensive statistics
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
            "Total_Size_MB": {"$sum": "$File_Size_MB"},
            "Avg_Size_MB": {"$avg": "$File_Size_MB"},
            "First_Date": {"$min": "$ISO_Date"},
            "Last_Date": {"$max": "$ISO_Date"},
            "Unique_Days": {"$addToSet": "$ISO_Date"}
        }
    },
    {
        "$addFields": {
            "Days_Active": {"$size": "$Unique_Days"},
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
            "Total_Size_MB": {"$round": ["$Total_Size_MB", 2]},
            "Avg_Size_MB": {"$round": ["$Avg_Size_MB", 2]},
            "First_Date": 1,
            "Last_Date": 1,
            "Days_Active": 1,
            "Avg_Files_Per_Day": 1
        }
    },
    {
        "$sort": {
            "Camera_Model": 1,
            "School_Year": 1
        }
    }
]

# =============================================================================
# CAMERA_USAGE_DATE_RANGE
# =============================================================================
# Analyzes camera usage date ranges
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

# Export all camera usage pipelines
CAMERA_PIPELINES = {
    "CAMERA_USAGE_BY_YEAR": CAMERA_USAGE_BY_YEAR,
    "CAMERA_USAGE_BY_YEAR_RANGE_DETAILED": CAMERA_USAGE_BY_YEAR_RANGE_DETAILED,
    "CAMERA_USAGE_DATE_RANGE": CAMERA_USAGE_DATE_RANGE,
    "CAMERA_USAGE_BY_YEAR_RANGE": CAMERA_USAGE_BY_YEAR_RANGE,
    "AUDIO_NOTE_CHARACTERISTICS": AUDIO_NOTE_CHARACTERISTICS
}
