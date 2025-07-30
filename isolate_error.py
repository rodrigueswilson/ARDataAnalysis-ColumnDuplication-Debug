#!/usr/bin/env python3
"""
Isolate the specific error in the statistical metrics calculation.
"""

import pandas as pd
from ar_utils import _load_config

def test_config_loading():
    """Test if config loading works correctly."""
    print("Testing config loading...")
    
    try:
        config = _load_config()
        print("✅ Config loaded successfully")
        
        # Test non-collection days extraction
        non_collection_days = set()
        for year_data in config['school_years'].values():
            for period_data in year_data['periods'].values():
                non_collection_days.update(period_data.get('non_collection_days', []))
        
        print(f"✅ Found {len(non_collection_days)} non-collection days")
        print(f"   Sample: {list(sorted(non_collection_days))[:5]}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error in config loading: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_dataframe_operations():
    """Test DataFrame operations that might be causing issues."""
    print("\nTesting DataFrame operations...")
    
    try:
        # Create a sample DataFrame similar to what the function receives
        sample_data = {
            '_id': ['2021-09-13', '2021-09-14', '2021-09-15'],
            'Total_Files': [10, 0, 25]
        }
        df = pd.DataFrame(sample_data)
        
        print("✅ Sample DataFrame created")
        print(f"   Shape: {df.shape}")
        print(f"   Columns: {list(df.columns)}")
        
        # Test date string conversion
        df['date_str'] = df['_id'].astype(str)
        print("✅ Date string conversion successful")
        
        # Test filtering with sample non-collection days
        sample_non_collection = {'2021-09-14'}  # Assume this is a non-collection day
        df_filtered = df[~df['date_str'].isin(sample_non_collection)].copy()
        print(f"✅ Filtering successful: {len(df)} -> {len(df_filtered)} rows")
        
        # Test statistical calculations
        if len(df_filtered) > 0:
            filtered_files = df_filtered['Total_Files']
            mean_val = filtered_files.mean()
            median_val = filtered_files.median()
            std_val = filtered_files.std()
            min_val = filtered_files.min()
            max_val = filtered_files.max()
            
            print(f"✅ Statistics calculated: mean={mean_val}, median={median_val}, std={std_val}")
            print(f"   min={min_val}, max={max_val}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error in DataFrame operations: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    config_ok = test_config_loading()
    df_ok = test_dataframe_operations()
    
    if config_ok and df_ok:
        print("\n✅ All tests passed - the issue might be elsewhere")
    else:
        print("\n❌ Found issues that need to be fixed")
