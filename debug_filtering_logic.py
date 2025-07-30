#!/usr/bin/env python3
"""
Debug the filtering logic to understand why statistical metrics aren't changing.
"""

import pandas as pd
from ar_utils import _load_config

def debug_filtering_logic():
    """Debug the filtering logic step by step."""
    print("=== DEBUGGING FILTERING LOGIC ===")
    
    try:
        # Step 1: Load config and check non-collection days
        print("Step 1: Loading config...")
        config = _load_config()
        
        print(f"Config keys: {list(config.keys())}")
        
        if 'non_collection_days' in config:
            non_collection_days_raw = config['non_collection_days']
            print(f"Raw non-collection days type: {type(non_collection_days_raw)}")
            print(f"Sample raw entries: {list(non_collection_days_raw.items())[:3]}")
            
            # Convert to set of strings (as the filtering logic does)
            non_collection_days = set(non_collection_days_raw.keys())
            print(f"Converted to string set: {len(non_collection_days)} days")
            print(f"Sample converted: {sorted(list(non_collection_days))[:5]}")
            
            # Check data types
            sample_key = list(non_collection_days)[0]
            print(f"Sample key type: {type(sample_key)}")
            print(f"Sample key value: {sample_key}")
        else:
            print("❌ No 'non_collection_days' found in config!")
            return
        
        # Step 2: Create sample data similar to what the function processes
        print(f"\nStep 2: Testing with sample data...")
        
        # Create sample data that includes some non-collection days
        sample_dates = [
            '2021-09-13',  # Collection day
            '2021-09-14',  # Collection day  
            '2021-09-06',  # Non-collection day (Labor Day)
            '2021-11-25',  # Non-collection day (Thanksgiving)
            '2021-12-24'   # Non-collection day (Christmas Eve)
        ]
        
        sample_data = {
            '_id': sample_dates,
            'Total_Files': [25, 30, 0, 0, 5]  # Non-collection days might have 0 or few files
        }
        
        df = pd.DataFrame(sample_data)
        print(f"Sample DataFrame:")
        print(df)
        
        # Step 3: Apply the same filtering logic as in the Summary Statistics
        print(f"\nStep 3: Applying filtering logic...")
        
        # Convert dates to string format (as done in the actual code)
        df['date_str'] = df['_id'].astype(str)
        print(f"Date strings: {df['date_str'].tolist()}")
        
        # Check which dates are in non_collection_days
        print(f"\nChecking which dates are non-collection days:")
        for date_str in df['date_str']:
            is_non_collection = date_str in non_collection_days
            print(f"  {date_str}: {'NON-COLLECTION' if is_non_collection else 'Collection'}")
        
        # Apply filtering
        df_filtered = df[~df['date_str'].isin(non_collection_days)].copy()
        
        print(f"\nFiltering results:")
        print(f"  Original records: {len(df)}")
        print(f"  Filtered records: {len(df_filtered)}")
        print(f"  Removed records: {len(df) - len(df_filtered)}")
        
        if len(df_filtered) < len(df):
            print(f"✅ Filtering is working!")
            print(f"Filtered DataFrame:")
            print(df_filtered)
            
            # Calculate statistics
            original_stats = {
                'mean': df['Total_Files'].mean(),
                'median': df['Total_Files'].median(),
                'min': df['Total_Files'].min(),
                'max': df['Total_Files'].max()
            }
            
            filtered_stats = {
                'mean': df_filtered['Total_Files'].mean(),
                'median': df_filtered['Total_Files'].median(), 
                'min': df_filtered['Total_Files'].min(),
                'max': df_filtered['Total_Files'].max()
            }
            
            print(f"\nStatistics comparison:")
            for stat in ['mean', 'median', 'min', 'max']:
                orig = original_stats[stat]
                filt = filtered_stats[stat]
                changed = abs(orig - filt) > 0.01
                print(f"  {stat:6}: Original={orig:6.1f}, Filtered={filt:6.1f}, Changed={'YES' if changed else 'No'}")
        else:
            print(f"❌ No filtering occurred!")
        
        # Step 4: Check the actual issue - date format mismatch
        print(f"\nStep 4: Investigating potential date format issues...")
        
        # The issue might be that config dates are datetime.date objects, not strings
        print(f"Non-collection days sample types:")
        for i, key in enumerate(list(non_collection_days)[:3]):
            print(f"  {i+1}. {key} (type: {type(key)})")
        
        # Convert non-collection days to strings if they're not already
        non_collection_days_str = set()
        for day in non_collection_days:
            if hasattr(day, 'strftime'):  # It's a date object
                non_collection_days_str.add(day.strftime('%Y-%m-%d'))
            else:  # It's already a string
                non_collection_days_str.add(str(day))
        
        print(f"Converted non-collection days to strings: {len(non_collection_days_str)}")
        print(f"Sample converted: {sorted(list(non_collection_days_str))[:5]}")
        
        # Test filtering with string conversion
        df_filtered_fixed = df[~df['date_str'].isin(non_collection_days_str)].copy()
        
        print(f"\nFixed filtering results:")
        print(f"  Original records: {len(df)}")
        print(f"  Fixed filtered records: {len(df_filtered_fixed)}")
        print(f"  Removed records: {len(df) - len(df_filtered_fixed)}")
        
        if len(df_filtered_fixed) < len(df):
            print(f"✅ Fixed filtering is working!")
            print(f"🔧 SOLUTION: Need to convert date objects to strings in the filtering logic")
        else:
            print(f"❌ Fixed filtering still not working")
    
    except Exception as e:
        print(f"❌ Error in debugging: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_filtering_logic()
