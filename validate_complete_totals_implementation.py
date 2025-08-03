#!/usr/bin/env python3
"""
Complete Totals Implementation Validation
=========================================

This script validates the entire systematic totals implementation across
all phases (Phase 1 + Phase 2 + Phase 3) to confirm comprehensive coverage.
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

def read_pipeline_sheet_correctly(file_path, sheet_name):
    """
    Read a pipeline sheet correctly, accounting for the header structure.
    """
    try:
        # Read with header at row 2 (0-indexed), which is row 3 in Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)
        
        # Clean up any unnamed columns or NaN column names
        df.columns = [f"Col_{i}" if pd.isna(col) or str(col).startswith('Unnamed') 
                     else str(col) for i, col in enumerate(df.columns)]
        
        # Remove any completely empty rows
        df = df.dropna(how='all')
        
        # Convert numeric columns properly
        for col in df.columns:
            if len(df) > 1:
                data_part = df[col].iloc[:-1]
                try:
                    numeric_data = pd.to_numeric(data_part, errors='coerce')
                    if not numeric_data.isna().all():
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                except:
                    pass
        
        return df
    except Exception as e:
        print(f"Error reading {sheet_name}: {e}")
        return None

def read_base_sheet_correctly(file_path, sheet_name):
    """
    Read base sheets (like Summary Statistics) with different structure.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except Exception as e:
        print(f"Error reading {sheet_name}: {e}")
        return None

def validate_complete_totals_implementation():
    """
    Validate the complete totals implementation across all phases.
    """
    print("🔍 COMPLETE TOTALS IMPLEMENTATION VALIDATION")
    print("="*60)
    
    # Complete implementation targets across all phases
    all_target_sheets = {
        # PHASE 1: Original high-priority (already validated)
        'Monthly Capture Volume': {
            'phase': 1,
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 1,
            'sheet_type': 'pipeline'
        },
        'File Size Stats': {
            'phase': 1,
            'expected_totals': 'full_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 2,
            'sheet_type': 'pipeline'
        },
        'Time of Day': {
            'phase': 1,
            'expected_totals': 'row_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 1,
            'sheet_type': 'pipeline'
        },
        'Weekday by Period': {
            'phase': 1,
            'expected_totals': 'row_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 1,
            'sheet_type': 'pipeline'
        },
        
        # PHASE 2: High-priority remaining
        'Weekly Counts': {
            'phase': 2,
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 3,
            'sheet_type': 'pipeline'
        },
        'Day of Week Counts': {
            'phase': 2,
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 3,
            'sheet_type': 'pipeline'
        },
        'Activity Counts': {
            'phase': 2,
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 3,
            'sheet_type': 'pipeline'
        },
        'File Size by Day': {
            'phase': 2,
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 2,
            'sheet_type': 'pipeline'
        },
        'SY Days': {
            'phase': 2,
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 3,
            'sheet_type': 'pipeline'
        },
        'Data Quality': {
            'phase': 2,
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 5,
            'sheet_type': 'pipeline'
        },
        
        # PHASE 3: Medium and low-priority
        'Collection Periods': {
            'phase': 3,
            'expected_totals': 'row_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 1,
            'sheet_type': 'pipeline'
        },
        'Camera Usage by Year': {
            'phase': 3,
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 1,
            'sheet_type': 'pipeline'
        },
        'Camera Usage Dates': {
            'phase': 3,
            'expected_totals': 'row_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 1,
            'sheet_type': 'pipeline'
        },
        'Camera Usage by Year Range': {
            'phase': 3,
            'expected_totals': 'row_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 1,
            'sheet_type': 'pipeline'
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
    
    # Validation results by phase
    phase_results = {1: {'total': 0, 'passed': 0}, 2: {'total': 0, 'passed': 0}, 3: {'total': 0, 'passed': 0}}
    validation_results = {}
    
    print(f"\n🎯 VALIDATING {len(all_target_sheets)} TOTAL IMPLEMENTATION TARGETS")
    print("-" * 60)
    
    for sheet_name, expected in all_target_sheets.items():
        phase = expected['phase']
        phase_results[phase]['total'] += 1
        
        print(f"\n📋 Phase {phase} - Validating: {sheet_name}")
        print(f"   Expected: {expected['expected_totals']} with {expected['min_numeric_columns']}+ numeric columns")
        
        # Read the sheet based on type
        if expected['sheet_type'] == 'pipeline':
            df = read_pipeline_sheet_correctly(report_file, sheet_name)
        else:
            df = read_base_sheet_correctly(report_file, sheet_name)
        
        if df is None or df.empty:
            validation_results[sheet_name] = {
                'phase': phase,
                'status': 'FAILED',
                'reason': 'Sheet is empty or could not be read'
            }
            print(f"   ❌ FAILED: Sheet is empty or could not be read")
            continue
        
        # Check for totals
        has_total_label = False
        total_label_found = None
        
        # Check last row for column totals
        if len(df) > 1:
            last_row = df.iloc[-1]
            for col in df.columns:
                cell_value = str(last_row[col]).strip()
                if cell_value == expected['expected_label']:
                    has_total_label = True
                    total_label_found = col
                    break
        
        # Check for row totals (rightmost column)
        has_row_totals = False
        if len(df.columns) > 1:
            rightmost_col = df.columns[-1]
            if 'total' in str(rightmost_col).lower() or 'row total' in str(rightmost_col).lower():
                has_row_totals = True
        
        # Count numeric columns
        numeric_columns = df.select_dtypes(include=['number']).columns.tolist()
        numeric_count = len(numeric_columns)
        
        # Determine if validation passes based on expected totals type
        validation_passed = False
        
        if expected['expected_totals'] == 'column_totals':
            validation_passed = has_total_label and numeric_count >= expected['min_numeric_columns']
        elif expected['expected_totals'] == 'row_totals':
            validation_passed = (has_total_label or has_row_totals) and numeric_count >= expected['min_numeric_columns']
        elif expected['expected_totals'] == 'full_totals':
            validation_passed = has_total_label and numeric_count >= expected['min_numeric_columns']
        
        if validation_passed:
            phase_results[phase]['passed'] += 1
            validation_results[sheet_name] = {
                'phase': phase,
                'status': 'PASSED',
                'details': {
                    'numeric_columns_count': numeric_count,
                    'has_column_totals': has_total_label,
                    'has_row_totals': has_row_totals,
                    'sheet_shape': df.shape
                }
            }
            print(f"   ✅ PASSED: {numeric_count} numeric columns, totals detected")
        else:
            issues = []
            if not has_total_label and not has_row_totals:
                issues.append("No totals found")
            if numeric_count < expected['min_numeric_columns']:
                issues.append(f"Only {numeric_count} numeric columns")
            
            validation_results[sheet_name] = {
                'phase': phase,
                'status': 'FAILED',
                'reason': '; '.join(issues),
                'details': {
                    'numeric_columns_count': numeric_count,
                    'has_column_totals': has_total_label,
                    'has_row_totals': has_row_totals,
                    'sheet_shape': df.shape
                }
            }
            print(f"   ❌ FAILED: {'; '.join(issues)}")
    
    # Phase-by-phase summary
    print(f"\n" + "="*60)
    print(f"📊 PHASE-BY-PHASE VALIDATION SUMMARY")
    print(f"="*60)
    
    total_passed = 0
    total_sheets = 0
    
    for phase in [1, 2, 3]:
        passed = phase_results[phase]['passed']
        total = phase_results[phase]['total']
        total_passed += passed
        total_sheets += total
        
        success_rate = (passed / total * 100) if total > 0 else 0
        print(f"Phase {phase}: {passed}/{total} sheets ({success_rate:.1f}% success)")
    
    overall_success_rate = (total_passed / total_sheets * 100) if total_sheets > 0 else 0
    
    print(f"\n🎯 OVERALL IMPLEMENTATION SUMMARY")
    print(f"✅ Total successful: {total_passed}/{total_sheets}")
    print(f"📈 Overall success rate: {overall_success_rate:.1f}%")
    
    if overall_success_rate == 100:
        print(f"🎉 COMPLETE SUCCESS! All systematic totals implemented correctly!")
    elif overall_success_rate >= 90:
        print(f"🚀 Excellent implementation! Minor issues to address.")
    elif overall_success_rate >= 75:
        print(f"✅ Good implementation! Some sheets need attention.")
    else:
        print(f"📋 Implementation needs work - review failed validations.")
    
    # Save comprehensive results
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    results_file = f"complete_totals_validation_{timestamp}.json"
    
    validation_summary = {
        'validation_timestamp': timestamp,
        'report_file': str(report_file),
        'total_sheets_validated': total_sheets,
        'successful_validations': total_passed,
        'overall_success_rate': overall_success_rate,
        'phase_results': phase_results,
        'implementation_status': 'COMPLETE' if overall_success_rate == 100 else 'PARTIAL',
        'sheet_results': validation_results
    }
    
    with open(results_file, 'w') as f:
        json.dump(validation_summary, f, indent=2, default=str)
    
    print(f"\n💾 Comprehensive results saved: {results_file}")
    
    return overall_success_rate >= 90

def check_final_coverage():
    """
    Check the final totals coverage across all sheets in the report.
    """
    print(f"\n📊 FINAL TOTALS COVERAGE ANALYSIS")
    print("="*40)
    
    try:
        report_file = find_latest_report()
        xl_file = pd.ExcelFile(report_file)
        all_sheets = xl_file.sheet_names
        
        sheets_with_totals = 0
        sheets_analyzed = 0
        coverage_details = []
        
        # Skip certain sheets that shouldn't have totals
        skip_sheets = ['Dashboard', 'Summary Statistics', 'Data Cleaning', 'Raw Data']
        
        for sheet_name in all_sheets:
            if any(skip in sheet_name for skip in skip_sheets):
                continue
                
            sheets_analyzed += 1
            
            try:
                # Try pipeline sheet reading first
                df = read_pipeline_sheet_correctly(report_file, sheet_name)
                if df is None:
                    df = read_base_sheet_correctly(report_file, sheet_name)
                
                if df is not None and not df.empty and len(df) > 1:
                    has_totals = False
                    
                    # Check last row for totals
                    last_row = df.iloc[-1]
                    for col in df.columns:
                        cell_value = str(last_row[col]).strip()
                        if cell_value in ['TOTAL', 'Row Total', 'Grand Total']:
                            has_totals = True
                            break
                    
                    # Check for row totals in column names
                    if not has_totals:
                        for col in df.columns:
                            if 'total' in str(col).lower():
                                has_totals = True
                                break
                    
                    if has_totals:
                        sheets_with_totals += 1
                        coverage_details.append(f"   ✅ {sheet_name}: Has totals")
                    else:
                        coverage_details.append(f"   ⚪ {sheet_name}: No totals")
                else:
                    coverage_details.append(f"   🚨 {sheet_name}: Could not analyze")
            except Exception as e:
                coverage_details.append(f"   🚨 {sheet_name}: Error reading sheet")
        
        # Print coverage details
        for detail in coverage_details:
            print(detail)
        
        coverage_percentage = (sheets_with_totals / sheets_analyzed * 100) if sheets_analyzed > 0 else 0
        
        print(f"\n📈 FINAL COVERAGE SUMMARY:")
        print(f"   Sheets with totals: {sheets_with_totals}/{sheets_analyzed}")
        print(f"   Coverage percentage: {coverage_percentage:.1f}%")
        
        # Coverage assessment
        if coverage_percentage >= 80:
            print(f"🎉 EXCELLENT COVERAGE! Target achieved.")
        elif coverage_percentage >= 60:
            print(f"✅ GOOD COVERAGE! Close to target.")
        else:
            print(f"📋 MORE WORK NEEDED for comprehensive coverage.")
        
        return coverage_percentage
        
    except Exception as e:
        print(f"❌ Error checking coverage: {e}")
        return 0

def main():
    """
    Main comprehensive validation function.
    """
    print("🚀 COMPLETE TOTALS IMPLEMENTATION VALIDATION")
    print("="*55)
    
    try:
        # Validate complete implementation
        implementation_success = validate_complete_totals_implementation()
        
        # Check final coverage
        final_coverage = check_final_coverage()
        
        print(f"\n" + "="*55)
        print(f"🎯 COMPREHENSIVE VALIDATION COMPLETE")
        print(f"="*55)
        
        if implementation_success:
            print(f"✅ Systematic implementation: SUCCESS")
        else:
            print(f"📋 Systematic implementation: NEEDS ATTENTION")
        
        print(f"📊 Final totals coverage: {final_coverage:.1f}%")
        
        if implementation_success and final_coverage >= 80:
            print(f"🎉 MISSION ACCOMPLISHED!")
            print(f"🚀 Comprehensive totals system fully operational!")
        else:
            print(f"📋 Review validation details for remaining work.")
            
    except Exception as e:
        print(f"❌ Validation error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
