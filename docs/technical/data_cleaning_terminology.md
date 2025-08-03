# Data Cleaning Terminology and Filter Logic

## Overview

This document describes the terminology and filtering logic used in the Data Cleaning sheet of the AR Data Analysis report system. The sheet presents a comprehensive analysis of the data filtering process, showing how raw data files are classified and filtered to create the final research dataset.

## Key Concepts

### 2×2 Matrix Filtering Logic

The data filtering system uses a 2×2 matrix approach based on two independent boolean criteria:

1. **Collection Period** (`is_collection_day`): Whether the file was recorded during an official school collection period
   - Valid Period (TRUE): File was recorded during a designated school collection day
   - Invalid Period (FALSE): File was recorded outside designated collection periods

2. **Recording Quality** (`Outlier_Status`): Whether the file meets technical quality standards
   - Valid Recording (FALSE): File meets technical quality requirements
   - Manual Exclusion (TRUE): File has been manually classified as a recording error

These two independent criteria create four mutually exclusive categories, ensuring no file is counted or excluded more than once:

| Category | Collection Period | Recording Quality | Status |
|----------|------------------|-------------------|--------|
| School Normal | Valid Period | Valid Recording | Included in research dataset |
| Recording Errors (School Days) | Valid Period | Manual Exclusion | Excluded: recording error |
| Non-Instructional Recordings | Invalid Period | Valid Recording | Excluded: wrong timing |
| Combined Exclusions | Invalid Period | Manual Exclusion | Excluded: both issues |

### Academic Terminology

The Data Cleaning sheet uses academically rigorous terminology to clearly distinguish between data states before and after the cleaning process:

| Term | Description |
|------|-------------|
| Media Type | File format (JPG or MP3) |
| Initial Collection Size | Complete raw dataset before quality assessment |
| Recording Errors Filtered | Files manually classified as recording mistakes (too short, too long, recorder not stopped properly, or otherwise irrelevant for AR collection analysis) |
| Non-Instructional Days Filtered | Files captured outside designated school collection periods |
| Combined Filters Applied | Files with both recording errors and non-instructional timing |
| Total Records Filtered | Sum of all excluded files |
| Research Dataset Size | Final dataset after all filters are applied |
| Filter Application Rate (%) | Percentage of files excluded from the initial collection |
| Dataset Validity Rate (%) | Percentage of files retained in the final research dataset |

## Data Cleaning Sheet Structure

The Data Cleaning sheet consists of three main sections:

1. **Intersection Analysis Table**: Shows the filtering statistics by media type and across all files
2. **Year-by-Year Breakdown Table**: Shows filtering statistics broken down by school year
3. **Logic Explanation Table**: Documents the 2×2 matrix filtering approach with counts for each category

## Mathematical Verification

The filtering logic has been mathematically verified to ensure:

- No files are excluded multiple times
- All files fall into exactly one category
- The sum of all categories equals the total number of files
- Filter Application Rate + Dataset Validity Rate = 100%

## Example Statistics

Based on real data analysis:
- Total files: 10,089
- School Normal (kept): 9,731 files (96.5%)
- School Outliers (excluded): 11 files
- Non-School Normal (excluded): 164 files
- Non-School Outliers (excluded): 183 files
- Total Excluded: 358 files (3.5%)

## Implementation Notes

- "Recording Errors" refers specifically to files that have been manually classified as outliers through human review
- This distinction is important for educational research contexts where data quality assessment protocols require explicit documentation
- The terminology has been specifically chosen to enhance clarity for academic audiences

## Related Code

The implementation of this filtering logic and terminology can be found in:
- `report_generator/sheet_creators/__init__.py` in the `create_data_cleaning_sheet` method
