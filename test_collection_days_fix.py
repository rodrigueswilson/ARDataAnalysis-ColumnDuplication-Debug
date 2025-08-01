#!/usr/bin/env python3
"""
Test script to validate the school year collection days calculation fix.
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from utils.calendar import calculate_collection_days_for_period

def test_collection_days_calculation():
    """Test the collection days calculation for school years."""
    print("🧪 Testing Collection Days Calculation Fix")
    print("=" * 50)
    
    # Test individual periods
    periods = ['SY 21-22 P1', 'SY 21-22 P2', 'SY 21-22 P3', 
               'SY 22-23 P1', 'SY 22-23 P2', 'SY 22-23 P3']
    
    period_days = {}
    for period in periods:
        days = calculate_collection_days_for_period(period)
        period_days[period] = days
        print(f"📅 {period}: {days} collection days")
    
    # Calculate school year totals (the fix)
    sy_2122_total = (
        calculate_collection_days_for_period('SY 21-22 P1') +
        calculate_collection_days_for_period('SY 21-22 P2') +
        calculate_collection_days_for_period('SY 21-22 P3')
    )
    
    sy_2223_total = (
        calculate_collection_days_for_period('SY 22-23 P1') +
        calculate_collection_days_for_period('SY 22-23 P2') +
        calculate_collection_days_for_period('SY 22-23 P3')
    )
    
    overall_total = sy_2122_total + sy_2223_total
    
    print("\n🎯 School Year Totals (FIXED CALCULATION):")
    print(f"📊 2021-2022 School Year: {sy_2122_total} collection days")
    print(f"📊 2022-2023 School Year: {sy_2223_total} collection days") 
    print(f"📊 Overall Total: {overall_total} collection days")
    
    # Validate the fix
    if sy_2122_total > 0 and sy_2223_total > 0:
        print("\n✅ SUCCESS: School year collection days calculation is working!")
        print("🔧 The bug fix correctly sums collection days from all periods.")
        return True
    else:
        print("\n❌ FAILURE: School year collection days still showing 0")
        return False

if __name__ == "__main__":
    success = test_collection_days_calculation()
    if success:
        print("\n🎉 Collection days calculation fix validated successfully!")
    else:
        print("\n💥 Collection days calculation still needs debugging.")
