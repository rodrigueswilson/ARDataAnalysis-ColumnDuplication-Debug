"""
Data Cleaning Utilities
======================

This module provides utilities for data cleaning in the AR Data Analysis system.
It implements the 2×2 matrix filtering approach based on:
- is_collection_day (TRUE/FALSE): Whether the day is a designated collection day
- Outlier_Status (TRUE/FALSE): Whether the file is flagged as an outlier

This creates four mutually exclusive categories:
1. School Normal (is_collection_day=TRUE, Outlier_Status=FALSE): Files kept in final dataset
2. School Outliers (is_collection_day=TRUE, Outlier_Status=TRUE): Outlier files from school days
3. Non-School Normal (is_collection_day=FALSE, Outlier_Status=FALSE): Regular files from non-school days
4. Non-School Outliers (is_collection_day=FALSE, Outlier_Status=TRUE): Outlier files from non-school days
"""

import pandas as pd


class DataCleaningUtils:
    """
    Utility class for AR Data Analysis data cleaning logic.
    
    Implements the 2×2 matrix filtering approach based on:
    - is_collection_day (TRUE/FALSE): Whether the day is a designated collection day
    - Outlier_Status (TRUE/FALSE): Whether the file is flagged as an outlier
    
    Provides methods for:
    1. Generating MongoDB aggregation pipelines for each category
    2. Calculating intersection values for the 2×2 matrix
    3. Computing retention and exclusion percentages
    4. Creating summary statistics for reporting
    """
    
    def __init__(self, db):
        """
        Initialize with MongoDB database connection.
        
        Args:
            db: MongoDB database connection
        """
        self.db = db
    
    def get_raw_pipeline(self, file_types=None, school_year=None):
        """
        Generate pipeline for total raw data (no filters except School_Year != "N/A").
        
        Args:
            file_types: Optional list of file types to filter (default: ["JPG", "MP3"])
            school_year: Optional school year filter (default: any non-N/A)
            
        Returns:
            List of MongoDB aggregation stages
        """
        if file_types is None:
            file_types = ["JPG", "MP3"]
            
        match_criteria = {
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": file_types}
        }
        
        if school_year:
            match_criteria["School_Year"] = school_year
            
        return [
            {"$match": match_criteria},
            {"$group": {
                "_id": "$file_type",
                "count": {"$sum": 1}
            }}
        ]
    
    def get_collection_pipeline(self, file_types=None, school_year=None):
        """
        Generate pipeline for collection days only (is_collection_day: TRUE).
        
        Args:
            file_types: Optional list of file types to filter (default: ["JPG", "MP3"])
            school_year: Optional school year filter (default: any non-N/A)
            
        Returns:
            List of MongoDB aggregation stages
        """
        if file_types is None:
            file_types = ["JPG", "MP3"]
            
        match_criteria = {
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": file_types},
            "is_collection_day": True
        }
        
        if school_year:
            match_criteria["School_Year"] = school_year
            
        return [
            {"$match": match_criteria},
            {"$group": {
                "_id": "$file_type",
                "count": {"$sum": 1}
            }}
        ]
    
    def get_non_outlier_pipeline(self, file_types=None, school_year=None):
        """
        Generate pipeline for non-outliers only (Outlier_Status: FALSE).
        
        Args:
            file_types: Optional list of file types to filter (default: ["JPG", "MP3"])
            school_year: Optional school year filter (default: any non-N/A)
            
        Returns:
            List of MongoDB aggregation stages
        """
        if file_types is None:
            file_types = ["JPG", "MP3"]
            
        match_criteria = {
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": file_types},
            "Outlier_Status": False
        }
        
        if school_year:
            match_criteria["School_Year"] = school_year
            
        return [
            {"$match": match_criteria},
            {"$group": {
                "_id": "$file_type",
                "count": {"$sum": 1}
            }}
        ]
    
    def get_both_pipeline(self, file_types=None, school_year=None):
        """
        Generate pipeline for both criteria (is_collection_day: TRUE AND Outlier_Status: FALSE).
        
        Args:
            file_types: Optional list of file types to filter (default: ["JPG", "MP3"])
            school_year: Optional school year filter (default: any non-N/A)
            
        Returns:
            List of MongoDB aggregation stages
        """
        if file_types is None:
            file_types = ["JPG", "MP3"]
            
        match_criteria = {
            "School_Year": {"$ne": "N/A"},
            "file_type": {"$in": file_types},
            "is_collection_day": True,
            "Outlier_Status": False
        }
        
        if school_year:
            match_criteria["School_Year"] = school_year
            
        return [
            {"$match": match_criteria},
            {"$group": {
                "_id": "$file_type",
                "count": {"$sum": 1}
            }}
        ]
    
    def run_aggregation(self, pipeline):
        """
        Execute MongoDB aggregation pipeline and return DataFrame.
        
        Args:
            pipeline: MongoDB aggregation pipeline
            
        Returns:
            pandas DataFrame with results
        """
        collection = self.db['media_records']
        results = list(collection.aggregate(pipeline))
        
        if not results:
            return pd.DataFrame()
        
        # The "_id" field might be a simple value or a dictionary depending on the group stage
        processed_results = []
        for result in results:
            processed = {'count': result.get('count', 0)}
            
            # Handle different types of _id field
            if isinstance(result.get('_id'), dict):
                processed.update(result['_id'])
            elif isinstance(result.get('_id'), str):
                processed['file_type'] = result['_id']
                
            processed_results.append(processed)
            
        return pd.DataFrame(processed_results)
    
    def df_to_dict(self, df):
        """
        Convert aggregation DataFrame to dictionary for easier lookup.
        
        Args:
            df: pandas DataFrame from aggregation
            
        Returns:
            Dictionary mapping file_type to count
        """
        if df.empty:
            return {}
        
        result = {}
        for _, row in df.iterrows():
            file_type = row.get('file_type', row.get('_id', 'Unknown'))
            count = row.get('count', 0)
            result[file_type] = count
        return result
    
    def calculate_intersection_data(self, raw_df, collection_df, non_outlier_df, both_df, file_types=None):
        """
        Calculate intersection values for the 2×2 matrix.
        
        Args:
            raw_df: DataFrame with raw data counts
            collection_df: DataFrame with collection days counts
            non_outlier_df: DataFrame with non-outlier counts
            both_df: DataFrame with counts meeting both criteria
            file_types: Optional list of file types (default: ["JPG", "MP3"])
            
        Returns:
            List of dictionaries with intersection data
        """
        if file_types is None:
            file_types = ["JPG", "MP3"]
        
        # Process results into dictionaries for easier lookup
        raw_counts = self.df_to_dict(raw_df)
        collection_counts = self.df_to_dict(collection_df)
        non_outlier_counts = self.df_to_dict(non_outlier_df)
        both_counts = self.df_to_dict(both_df)
        
        # Calculate intersection values
        intersection_data = []
        
        for file_type in file_types:
            total_raw = raw_counts.get(file_type, 0)
            collection_only = collection_counts.get(file_type, 0)
            non_outlier_only = non_outlier_counts.get(file_type, 0)
            both_criteria = both_counts.get(file_type, 0)
            
            # Calculate files that meet only one criterion (exclusive)
            # These are school outliers: collection days but are outliers
            school_outliers = collection_only - both_criteria
            
            # These are non-school normal: not collection days but not outliers
            non_school_normal = non_outlier_only - both_criteria
            
            # Calculate files that meet neither criterion
            # Neither = Total - (Collection Days Total) - (Non-Outliers Total) + (Both)
            # This gives us files that are: NOT collection days AND ARE outliers
            non_school_outliers = total_raw - collection_only - non_outlier_only + both_criteria
            
            # Final dataset = both criteria (school normal)
            school_normal = both_criteria
            
            # Calculate total excluded files (all categories except School Normal)
            total_excluded = school_outliers + non_school_normal + non_school_outliers
            
            # Calculate percentages
            retention_pct = (school_normal / total_raw * 100) if total_raw > 0 else 0
            exclusion_pct = (total_excluded / total_raw * 100) if total_raw > 0 else 0
            
            intersection_data.append({
                'file_type': file_type,
                'total_files': total_raw,
                'school_outliers': school_outliers,
                'non_school_normal': non_school_normal,
                'non_school_outliers': non_school_outliers,
                'total_excluded': total_excluded,
                'school_normal': school_normal,
                'exclusion_pct': exclusion_pct,
                'retention_pct': retention_pct
            })
        
        return intersection_data
    
    def get_complete_cleaning_data(self, file_types=None, school_year=None):
        """
        Get complete data cleaning analysis in a single call.
        
        Args:
            file_types: Optional list of file types (default: ["JPG", "MP3"])
            school_year: Optional school year filter
            
        Returns:
            Dictionary with intersection_data and totals
        """
        if file_types is None:
            file_types = ["JPG", "MP3"]
            
        # Run all four aggregations
        raw_pipeline = self.get_raw_pipeline(file_types, school_year)
        collection_pipeline = self.get_collection_pipeline(file_types, school_year)
        non_outlier_pipeline = self.get_non_outlier_pipeline(file_types, school_year)
        both_pipeline = self.get_both_pipeline(file_types, school_year)
        
        raw_df = self.run_aggregation(raw_pipeline)
        collection_df = self.run_aggregation(collection_pipeline)
        non_outlier_df = self.run_aggregation(non_outlier_pipeline)
        both_df = self.run_aggregation(both_pipeline)
        
        # Calculate intersection data
        intersection_data = self.calculate_intersection_data(raw_df, collection_df, non_outlier_df, both_df, file_types)
        
        # Calculate totals
        totals = {
            'file_type': 'TOTAL',
            'total_files': sum(item['total_files'] for item in intersection_data),
            'school_outliers': sum(item['school_outliers'] for item in intersection_data),
            'non_school_normal': sum(item['non_school_normal'] for item in intersection_data),
            'non_school_outliers': sum(item['non_school_outliers'] for item in intersection_data),
            'total_excluded': sum(item['total_excluded'] for item in intersection_data),
            'school_normal': sum(item['school_normal'] for item in intersection_data)
        }
        
        # Calculate percentages for totals
        if totals['total_files'] > 0:
            totals['exclusion_pct'] = (totals['total_excluded'] / totals['total_files'] * 100)
            totals['retention_pct'] = (totals['school_normal'] / totals['total_files'] * 100)
        else:
            totals['exclusion_pct'] = 0
            totals['retention_pct'] = 0
            
        # Add totals to the intersection data for convenience
        intersection_data.append(totals)
        
        return {
            'intersection_data': intersection_data,
            'totals': totals
        }
    
    def get_year_breakdown_data(self, years=None, file_types=None):
        """
        Get year-by-year breakdown of data cleaning metrics.
        
        Args:
            years: List of school years to analyze (default: ["2021-2022", "2022-2023"])
            file_types: List of file types to analyze (default: ["JPG", "MP3"])
            
        Returns:
            List of dictionaries with year breakdown data and totals
        """
        if years is None:
            years = ["2021-2022", "2022-2023"]
            
        if file_types is None:
            file_types = ["JPG", "MP3"]
            
        collection = self.db['media_records']
        year_breakdown_data = []
        
        for year in years:
            for file_type in file_types:
                category = f"{year} {file_type} Files"
                
                # Total files for this year/type
                total_files = collection.count_documents({
                    "School_Year": year,
                    "file_type": file_type
                })
                
                # Outliers: is_collection_day: TRUE AND Outlier_Status: TRUE
                outliers = collection.count_documents({
                    "School_Year": year,
                    "file_type": file_type,
                    "is_collection_day": True,
                    "Outlier_Status": True
                })
                
                # Non-School Days: is_collection_day: FALSE AND Outlier_Status: FALSE
                non_school_days = collection.count_documents({
                    "School_Year": year,
                    "file_type": file_type,
                    "is_collection_day": False,
                    "Outlier_Status": False
                })
                
                # Non-School Days and Outliers: is_collection_day: FALSE AND Outlier_Status: TRUE
                non_school_outliers = collection.count_documents({
                    "School_Year": year,
                    "file_type": file_type,
                    "is_collection_day": False,
                    "Outlier_Status": True
                })
                
                # School Days (Final Dataset): is_collection_day: TRUE AND Outlier_Status: FALSE
                school_days = collection.count_documents({
                    "School_Year": year,
                    "file_type": file_type,
                    "is_collection_day": True,
                    "Outlier_Status": False
                })
                
                # Calculate totals and percentages
                total_excluded = outliers + non_school_days + non_school_outliers
                exclusion_pct = (total_excluded / total_files * 100) if total_files > 0 else 0
                retention_pct = (school_days / total_files * 100) if total_files > 0 else 0
                
                year_breakdown_data.append({
                    'category': category,
                    'total_files': total_files,
                    'outliers': outliers,
                    'non_school_days': non_school_days,
                    'non_school_outliers': non_school_outliers,
                    'total_excluded': total_excluded,
                    'school_days': school_days,
                    'exclusion_pct': exclusion_pct,
                    'retention_pct': retention_pct
                })
        
        # Calculate totals
        year_totals = {
            'category': 'TOTAL',
            'total_files': sum(item['total_files'] for item in year_breakdown_data),
            'outliers': sum(item['outliers'] for item in year_breakdown_data),
            'non_school_days': sum(item['non_school_days'] for item in year_breakdown_data),
            'non_school_outliers': sum(item['non_school_outliers'] for item in year_breakdown_data),
            'total_excluded': sum(item['total_excluded'] for item in year_breakdown_data),
            'school_days': sum(item['school_days'] for item in year_breakdown_data)
        }
        
        # Calculate percentages for totals
        if year_totals['total_files'] > 0:
            year_totals['exclusion_pct'] = (year_totals['total_excluded'] / year_totals['total_files'] * 100)
            year_totals['retention_pct'] = (year_totals['school_days'] / year_totals['total_files'] * 100)
        else:
            year_totals['exclusion_pct'] = 0
            year_totals['retention_pct'] = 0
            
        # Add totals to the year breakdown data
        year_breakdown_data.append(year_totals)
        
        return year_breakdown_data
    
    def get_logic_explanation_data(self, totals):
        """
        Generate logic explanation data for the filtering categories.
        
        Args:
            totals: Dictionary with data cleaning totals
            
        Returns:
            List of tuples with logic explanation data
        """
        return [
            ('Recording Errors (School Days)', 'Valid Period', 'Manual Exclusion', totals['school_outliers']),
            ('Non-Instructional Recordings', 'Invalid Period', 'Valid Recording', totals['non_school_normal']),
            ('Combined Exclusions', 'Invalid Period', 'Manual Exclusion', totals['non_school_outliers']),
            ('Research Dataset (Final)', 'Valid Period', 'Valid Recording', totals['school_normal'])
        ]
