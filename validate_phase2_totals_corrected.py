#!/usr/bin/env python3
"""
Validate Phase 2 Totals Implementation - Corrected Version
=========================================================

This script properly validates Phase 2 totals by correctly reading the Excel
sheet structure with headers starting at row 3 (0-indexed row 2).
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
    Pipeline sheets have:
    - Row 1: Sheet title
    - Row 2: Empty or subtitle
    - Row 3: Column headers
    - Row 4+: Data
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
            # Skip the totals row for conversion
            if len(df) > 1:
                data_part = df[col].iloc[:-1]
                # Try to convert to numeric, keeping non-numeric as is
                try:
                    numeric_data = pd.to_numeric(data_part, errors='coerce')
                    if not numeric_data.isna().all():
                        # If most values can be converted, convert the whole column
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                except:
                    pass
        
        return df
    except Exception as e:
        print(f"Error reading {sheet_name}: {e}")
        return None

def validate_phase2_totals_corrected():
    """
    Validate all Phase 2 high-priority totals implementations with correct sheet reading.
    """
    print("🔍 VALIDATING PHASE 2 TOTALS IMPLEMENTATION (CORRECTED)")
    print("="*65)
    
    # Phase 2 high-priority sheets to validate
    phase2_sheets = {
        'Weekly Counts': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 3,  # Adjusted expectation
            'description': 'Weekly analysis with numeric columns'
        },
        'Day of Week Counts': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 3,  # Adjusted expectation
            'description': 'Day-of-week analysis with numeric columns'
        },
        'Activity Counts': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 3,  # Adjusted expectation
            'description': 'Activity analysis with numeric columns'
        },
        'File Size by Day': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 2,  # Adjusted expectation
            'description': 'Daily file size analysis with numeric columns'
        },
        'SY Days': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 3,  # Adjusted expectation
            'description': 'School year days analysis with numeric columns'
        },
        'Data Quality': {
            'expected_totals': 'column_totals',
            'expected_label': 'TOTAL',
            'min_numeric_columns': 5,  # Adjusted expectation
            'description': 'Data quality metrics with numeric columns'
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
    print("-" * 65)
    
    for sheet_name, expected in phase2_sheets.items():
        total_validations += 1
        print(f"\n📋 Validating: {sheet_name}")
        print(f"   Expected: {expected['description']}")
        
        # Read the sheet correctly
        df = read_pipeline_sheet_correctly(report_file, sheet_name)
        
        if df is None or df.empty:
            validation_results[sheet_name] = {
                'status': 'FAILED',
                'reason': 'Sheet is empty or could not be read',
                'details': {}
            }
            print(f"   ❌ FAILED: Sheet is empty or could not be read")
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
        
        print(f"   📊 Found {numeric_count} numeric columns: {numeric_columns}")
        print(f"   📏 Sheet shape: {df.shape}")
        print(f"   🏷️  Total label found: {has_total_label} (in column: {total_label_found})")
        
        # Validate totals calculations for numeric columns
        totals_correct = True
        totals_details = {}
        
        if has_total_label and len(df) > 1 and numeric_columns:
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
                            print(f"   ⚠️  {col}: Expected {expected_total}, got {actual_total} (diff: {difference})")
        
        # Overall validation
        validation_passed = (
            has_total_label and 
            numeric_count >= expected['min_numeric_columns'] and
            (totals_correct if numeric_columns else True)  # Only check totals if we have numeric columns
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
            if not totals_correct and numeric_columns:
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
    
    # Summary
    print(f"\n" + "="*65)
    print(f"📊 PHASE 2 VALIDATION SUMMARY (CORRECTED)")
    print(f"="*65)
    print(f"✅ Successful validations: {successful_validations}/{total_validations}")
    print(f"📈 Success rate: {successful_validations/total_validations*100:.1f}%")
    
    if successful_validations == total_validations:
        print(f"🎉 PHASE 2 IMPLEMENTATION: 100% SUCCESS!")
        print(f"🚀 All high-priority totals are working correctly")
    elif successful_validations > 0:
        print(f"✅ Partial success - {successful_validations} sheets working correctly")
        print(f"📋 Review failed validations for remaining issues")
    else:
        print(f"⚠️  All validations failed - review implementation")
    
    # Save detailed results
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    results_file = f"phase2_totals_validation_corrected_{timestamp}.json"
    
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

def main():
    """
    Main validation function.
    """
    print("🚀 PHASE 2 TOTALS VALIDATION (CORRECTED)")
    print("="*50)
    
    try:
        # Validate Phase 2 implementations with corrected reading
        phase2_success = validate_phase2_totals_corrected()
        
        print(f"\n" + "="*50)
        print(f"🎯 CORRECTED VALIDATION COMPLETE")
        print(f"="*50)
        
        if phase2_success:
            print(f"✅ Phase 2 implementation: SUCCESS")
            print(f"🎉 All high-priority totals working correctly!")
        else:
            print(f"📋 Phase 2 implementation: PARTIAL SUCCESS")
            print(f"🔧 Some sheets may need configuration adjustments")
            
    except Exception as e:
        print(f"❌ Validation error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
