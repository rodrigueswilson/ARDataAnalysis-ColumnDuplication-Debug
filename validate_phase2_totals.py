#!/usr/bin/env python3
"""
Validate Phase 2 Totals Implementation
=====================================

This script validates that all Phase 2 high-priority totals have been
successfully implemented and are working correctly.
"""

import pandas as pd
from pathlib import Path
import json
from datetime import datetime

def find_latest_report():
    """Find the most recent Excel report file."""
    report_files = list(Path('.').glob('AR_Analysis_Report_*.xlsx'))
    if not report_files:
        raise FileNotFoundError("No Excel report files found")
    
    latest_file = max(report_files, key=lambda f: f.stat().st_mtime)
    return latest_file

def validate_phase2_totals():
    """
    Validate all Phase 2 high-priority totals implementations.
    """
    print("🔍 VALIDATING PHASE 2 TOTALS IMPLEMENTATION")
    print("="*60)
    
    # Phase 2 high-priority sheets to validate
    phase2_sheets = {
        'Weekly Counts': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 5,
            'description': '5 numeric columns, weekly analysis'
        },
        'Day of Week Counts': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 4,
            'description': '4 numeric columns, day-of-week analysis'
        },
        'Activity Counts': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 4,
            'description': '4 numeric columns, activity analysis'
        },
        'File Size by Day': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 3,
            'description': '3 numeric columns, daily file size analysis'
        },
        'SY Days': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 4,
            'description': '4 numeric columns, school year days analysis'
        },
        'Data Quality': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 11,
            'description': '11 numeric columns, data quality metrics'
        }
    }
    
    # Find latest report
    try:
        report_file = find_latest_report()
        print(f"📊 Analyzing report: {report_file}")
        print(f"📅 Report timestamp: {datetime.fromtimestamp(report_file.stat().st_mtime)}")
    except Exception as e:
        print(f"❌ Error finding report: {e}")
        return False
    
    # Validation results
    validation_results = {}
    total_validations = 0
    successful_validations = 0
    
    print(f"\n🎯 VALIDATING {len(phase2_sheets)} PHASE 2 HIGH-PRIORITY SHEETS")
    print("-" * 60)
    
    for sheet_name, expected in phase2_sheets.items():
        total_validations += 1
        print(f"\n📋 Validating: {sheet_name}")
        print(f"   Expected: {expected['description']}")
        
        try:
            # Read the sheet
            df = pd.read_excel(report_file, sheet_name=sheet_name)
            
            # Basic validation
            if df.empty:
                validation_results[sheet_name] = {
                    'status': 'FAILED',
                    'reason': 'Sheet is empty',
                    'details': {}
                }
                print(f"   ❌ FAILED: Sheet is empty")
                continue
            
            # Check for totals row (should be last row with TOTAL label)
            last_row = df.iloc[-1]
            has_total_label = False
            total_label_found = None
            
            # Check if any cell in the last row contains the expected label
            for col in df.columns:
                cell_value = str(last_row[col]).strip()
                if cell_value == expected['expected_label']:
                    has_total_label = True
                    total_label_found = col
                    break
            
            # Count numeric columns
            numeric_columns = df.select_dtypes(include=['number']).columns.tolist()
            numeric_count = len(numeric_columns)
            
            # Validate totals calculations for numeric columns
            totals_correct = True
            totals_details = {}
            
            if has_total_label and len(df) > 1:
                data_rows = df.iloc[:-1]  # All rows except the last (totals) row
                totals_row = df.iloc[-1]
                
                for col in numeric_columns:
                    if col in df.columns:
                        expected_total = data_rows[col].sum()
                        actual_total = totals_row[col]
                        
                        # Handle potential NaN values
                        if pd.isna(actual_total):
                            totals_details[col] = {
                                'expected': expected_total,
                                'actual': 'NaN',
                                'correct': False
                            }
                            totals_correct = False
                        else:
                            difference = abs(expected_total - actual_total)
                            tolerance = max(0.01, abs(expected_total) * 0.001)  # 0.1% tolerance
                            is_correct = difference <= tolerance
                            
                            totals_details[col] = {
                                'expected': expected_total,
                                'actual': actual_total,
                                'difference': difference,
                                'correct': is_correct
                            }
                            
                            if not is_correct:
                                totals_correct = False
            
            # Overall validation
            validation_passed = (
                has_total_label and 
                numeric_count >= expected['min_numeric_columns'] and
                totals_correct
            )
            
            if validation_passed:
                successful_validations += 1
                validation_results[sheet_name] = {
                    'status': 'PASSED',
                    'details': {
                        'total_label_found': total_label_found,
                        'numeric_columns_count': numeric_count,
                        'numeric_columns': numeric_columns,
                        'totals_calculations': totals_details,
                        'rows_count': len(df)
                    }
                }
                print(f"   ✅ PASSED: {numeric_count} numeric columns, totals correct")
            else:
                issues = []
                if not has_total_label:
                    issues.append("No TOTAL label found in last row")
                if numeric_count < expected['min_numeric_columns']:
                    issues.append(f"Only {numeric_count} numeric columns (expected {expected['min_numeric_columns']})")
                if not totals_correct:
                    issues.append("Totals calculations incorrect")
                
                validation_results[sheet_name] = {
                    'status': 'FAILED',
                    'reason': '; '.join(issues),
                    'details': {
                        'total_label_found': total_label_found,
                        'numeric_columns_count': numeric_count,
                        'numeric_columns': numeric_columns,
                        'totals_calculations': totals_details,
                        'rows_count': len(df)
                    }
                }
                print(f"   ❌ FAILED: {'; '.join(issues)}")
        
        except Exception as e:
            validation_results[sheet_name] = {
                'status': 'ERROR',
                'reason': f'Exception during validation: {str(e)}',
                'details': {}
            }
            print(f"   🚨 ERROR: {str(e)}")
    
    # Summary
    print(f"\n" + "="*60)
    print(f"📊 PHASE 2 VALIDATION SUMMARY")
    print(f"="*60)
    print(f"✅ Successful validations: {successful_validations}/{total_validations}")
    print(f"📈 Success rate: {successful_validations/total_validations*100:.1f}%")
    
    if successful_validations == total_validations:
        print(f"🎉 PHASE 2 IMPLEMENTATION: 100% SUCCESS!")
        print(f"🚀 All high-priority totals are working correctly")
    else:
        print(f"⚠️  Some validations failed - review details above")
    
    # Save detailed results
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    results_file = f"phase2_totals_validation_{timestamp}.json"
    
    validation_summary = {
        'validation_timestamp': timestamp,
        'report_file': str(report_file),
        'total_sheets_validated': total_validations,
        'successful_validations': successful_validations,
        'success_rate': successful_validations/total_validations*100,
        'phase2_status': 'COMPLETE' if successful_validations == total_validations else 'PARTIAL',
        'sheet_results': validation_results
    }
    
    with open(results_file, 'w') as f:
        json.dump(validation_summary, f, indent=2, default=str)
    
    print(f"\n💾 Detailed results saved: {results_file}")
    
    return successful_validations == total_validations

def check_overall_totals_coverage():
    """
    Check the overall totals coverage across all sheets.
    """
    print(f"\n📊 CHECKING OVERALL TOTALS COVERAGE")
    print("="*40)
    
    try:
        report_file = find_latest_report()
        
        # Get all sheet names
        xl_file = pd.ExcelFile(report_file)
        all_sheets = xl_file.sheet_names
        
        sheets_with_totals = 0
        sheets_analyzed = 0
        
        for sheet_name in all_sheets:
            # Skip certain sheets that shouldn't have totals
            skip_sheets = ['Dashboard', 'Summary Statistics', 'Data Cleaning']
            if any(skip in sheet_name for skip in skip_sheets):
                continue
                
            sheets_analyzed += 1
            
            try:
                df = pd.read_excel(report_file, sheet_name=sheet_name)
                if not df.empty and len(df) > 1:
                    last_row = df.iloc[-1]
                    
                    # Check if last row contains totals
                    has_totals = False
                    for col in df.columns:
                        cell_value = str(last_row[col]).strip()
                        if cell_value in ['TOTAL', 'Row Total', 'Grand Total']:
                            has_totals = True
                            break
                    
                    if has_totals:
                        sheets_with_totals += 1
                        print(f"   ✅ {sheet_name}: Has totals")
                    else:
                        print(f"   ⚪ {sheet_name}: No totals")
            except Exception as e:
                print(f"   🚨 {sheet_name}: Error reading sheet")
        
        coverage_percentage = (sheets_with_totals / sheets_analyzed * 100) if sheets_analyzed > 0 else 0
        
        print(f"\n📈 OVERALL COVERAGE:")
        print(f"   Sheets with totals: {sheets_with_totals}/{sheets_analyzed}")
        print(f"   Coverage percentage: {coverage_percentage:.1f}%")
        
        return coverage_percentage
        
    except Exception as e:
        print(f"❌ Error checking coverage: {e}")
        return 0

def main():
    """
    Main validation function.
    """
    print("🚀 PHASE 2 TOTALS VALIDATION")
    print("="*50)
    
    try:
        # Validate Phase 2 implementations
        phase2_success = validate_phase2_totals()
        
        # Check overall coverage
        coverage = check_overall_totals_coverage()
        
        print(f"\n" + "="*50)
        print(f"🎯 VALIDATION COMPLETE")
        print(f"="*50)
        
        if phase2_success:
            print(f"✅ Phase 2 implementation: SUCCESS")
        else:
            print(f"⚠️  Phase 2 implementation: NEEDS ATTENTION")
        
        print(f"📊 Overall totals coverage: {coverage:.1f}%")
        
        if coverage >= 80:
            print(f"🎉 Excellent totals coverage achieved!")
        elif coverage >= 60:
            print(f"📈 Good totals coverage - consider remaining sheets")
        else:
            print(f"📋 More totals implementation recommended")
            
    except Exception as e:
        print(f"❌ Validation error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
