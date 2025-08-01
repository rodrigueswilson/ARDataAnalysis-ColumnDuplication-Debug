#!/usr/bin/env python3
"""
Quick Totals System Verification
===============================

Simple test to verify that the key totals system components are working.
"""

import sys
import pandas as pd
from pathlib import Path

# Add the project root to the path
sys.path.append(str(Path(__file__).parent))

def test_basic_functionality():
    """Test basic totals system functionality."""
    print("🔍 Testing Basic Totals System Functionality")
    print("=" * 50)
    
    try:
        # Test 1: TotalsManager import and initialization
        print("\n1. Testing TotalsManager import and initialization...")
        from report_generator.totals_manager import TotalsManager
        totals_manager = TotalsManager()
        print("✅ TotalsManager imported and initialized successfully")
        
        # Test 2: Check if add_totals_to_worksheet method exists
        print("\n2. Testing add_totals_to_worksheet method...")
        if hasattr(totals_manager, 'add_totals_to_worksheet'):
            print("✅ add_totals_to_worksheet method exists")
        else:
            print("❌ add_totals_to_worksheet method missing")
            return False
        
        # Test 3: TotalsIntegrationHelper import and method
        print("\n3. Testing TotalsIntegrationHelper...")
        from report_generator.totals_integration_guide import TotalsIntegrationHelper
        helper = TotalsIntegrationHelper()
        
        if hasattr(helper, 'generate_recommended_config'):
            print("✅ TotalsIntegrationHelper has generate_recommended_config method")
            
            # Test the method with sample data
            test_df = pd.DataFrame({
                'Date': ['2023-01-01', '2023-01-02'],
                'Total_Files': [10, 15],
                'ACF_Lag_1': [0.5, 0.3]  # Should be excluded
            })
            
            config = helper.generate_recommended_config(test_df, 'daily')
            if config and 'add_totals' in config:
                print("✅ generate_recommended_config works correctly")
                print(f"   Generated config: {config}")
            else:
                print("❌ generate_recommended_config failed")
                return False
        else:
            print("❌ generate_recommended_config method missing")
            return False
        
        # Test 4: Validation rules structure
        print("\n4. Testing validation rules configuration...")
        import json
        rules_path = Path("totals_validation_rules.json")
        if rules_path.exists():
            with open(rules_path, 'r') as f:
                rules = json.load(f)
            
            if 'validation_rules' in rules:
                validation_rules = rules['validation_rules']
                if 'global_settings' in validation_rules and 'validation_groups' in validation_rules:
                    print("✅ Validation rules structure is correct")
                else:
                    print("❌ Validation rules missing required nested keys")
                    return False
            else:
                print("❌ Validation rules missing top-level key")
                return False
        else:
            print("❌ Validation rules file not found")
            return False
        
        # Test 5: Basic worksheet totals functionality
        print("\n5. Testing basic totals calculation...")
        import openpyxl
        
        # Create test workbook and worksheet
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # Create test DataFrame
        test_data = pd.DataFrame({
            'Category': ['A', 'B', 'C'],
            'Value1': [10, 20, 30],
            'Value2': [5, 15, 25]
        })
        
        # Add test data to worksheet
        for r_idx, row in enumerate(test_data.itertuples(index=False), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
        
        # Test configuration
        config = {
            'add_row_totals': False,  # Keep it simple for this test
            'add_column_totals': True,
            'add_grand_total': False,
            'exclude_columns': ['Category'],
            'totals_label': 'TOTALS'
        }
        
        # Try to add totals
        totals_manager.add_totals_to_worksheet(
            worksheet=ws,
            dataframe=test_data,
            start_row=1,
            start_col=1,
            config=config
        )
        
        print("✅ Basic totals calculation completed without errors")
        
        print("\n🎉 All basic functionality tests passed!")
        return True
        
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_basic_functionality()
    if success:
        print("\n✅ Totals system basic functionality is working!")
        print("You can now run the full integration tests.")
    else:
        print("\n❌ Basic functionality test failed. Please review the errors above.")
    
    sys.exit(0 if success else 1)
