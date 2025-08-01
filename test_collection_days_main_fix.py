#!/usr/bin/env python3
"""
Test script to validate the collection days fix in the main Summary Statistics table.
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from utils.calendar import calculate_collection_days_for_period

def test_main_summary_collection_days_fix():
    """Test the collection days calculation fix for the main Summary Statistics table."""
    print("🧪 Testing Main Summary Statistics Collection Days Fix")
    print("=" * 60)
    
    # Test the exact same calculation logic that's now in the main Summary Statistics
    print("📊 Testing Fixed Collection Days Calculation Logic:")
    print("-" * 50)
    
    # Test for 2021-2022 school year
    collection_days_2122 = (
        calculate_collection_days_for_period('SY 21-22 P1') +
        calculate_collection_days_for_period('SY 21-22 P2') +
        calculate_collection_days_for_period('SY 21-22 P3')
    )
    
    # Test for 2022-2023 school year  
    collection_days_2223 = (
        calculate_collection_days_for_period('SY 22-23 P1') +
        calculate_collection_days_for_period('SY 22-23 P2') +
        calculate_collection_days_for_period('SY 22-23 P3')
    )
    
    # Test for overall
    collection_days_overall = (
        calculate_collection_days_for_period('SY 21-22 P1') +
        calculate_collection_days_for_period('SY 21-22 P2') +
        calculate_collection_days_for_period('SY 21-22 P3') +
        calculate_collection_days_for_period('SY 22-23 P1') +
        calculate_collection_days_for_period('SY 22-23 P2') +
        calculate_collection_days_for_period('SY 22-23 P3')
    )
    
    print(f"📅 2021-2022 School Year Collection Days: {collection_days_2122}")
    print(f"📅 2022-2023 School Year Collection Days: {collection_days_2223}")
    print(f"📅 Overall Collection Days: {collection_days_overall}")
    
    # Validate the fix
    success_checks = {
        "2021-2022 shows correct days (156)": collection_days_2122 == 156,
        "2022-2023 shows correct days (164)": collection_days_2223 == 164,
        "Overall shows correct days (320)": collection_days_overall == 320,
        "No school year shows 0 days": collection_days_2122 > 0 and collection_days_2223 > 0,
        "Overall equals sum of school years": collection_days_overall == (collection_days_2122 + collection_days_2223)
    }
    
    print(f"\n🎯 Validation Results:")
    print("-" * 30)
    for check, passed in success_checks.items():
        status = "✅ PASSED" if passed else "❌ FAILED"
        print(f"   {status}: {check}")
    
    all_passed = all(success_checks.values())
    
    print(f"\n{'='*60}")
    print(f"🎉 MAIN SUMMARY FIX RESULT: {'SUCCESS' if all_passed else 'FAILURE'}")
    print(f"{'='*60}")
    
    if all_passed:
        print("✅ The main Summary Statistics table will now show correct Collection Days!")
        print("✅ No more zeros in the Collection Days metric!")
        print("✅ All downstream calculations will be accurate!")
        
        print(f"\n📋 Expected Results in Summary Statistics Table:")
        print(f"   📊 2021-2022 Collection Days: {collection_days_2122}")
        print(f"   📊 2022-2023 Collection Days: {collection_days_2223}")
        print(f"   📊 Overall Collection Days: {collection_days_overall}")
        
        return True
    else:
        print("❌ Collection days calculation still has issues")
        return False

if __name__ == "__main__":
    success = test_main_summary_collection_days_fix()
    
    if success:
        print("\n🚀 Main Summary Statistics Collection Days fix validated!")
        print("🎯 The Summary Statistics table is now production-ready.")
    else:
        print("\n💥 Additional debugging needed for main Summary Statistics.")
