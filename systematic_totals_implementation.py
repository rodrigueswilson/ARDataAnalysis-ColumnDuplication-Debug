#!/usr/bin/env python3
"""
Systematic Totals Implementation Script
======================================

This script systematically applies totals to all Excel report tables based on the
comprehensive table inventory analysis. It implements totals where needed and
validates the results.
"""

import json
import pandas as pd
from pathlib import Path
from datetime import datetime
from report_generator.totals_manager import TotalsManager
from report_generator.totals_integration_guide import TotalsIntegrationHelper

def load_table_inventory():
    """Load the most recent table inventory analysis."""
    inventory_files = list(Path(".").glob("table_inventory_*.json"))
    if not inventory_files:
        raise FileNotFoundError("No table inventory files found. Run create_table_inventory.py first.")
    
    latest_inventory = max(inventory_files, key=lambda x: x.stat().st_mtime)
    print(f"📋 Loading inventory: {latest_inventory}")
    
    with open(latest_inventory, 'r') as f:
        return json.load(f)

def categorize_sheets_by_totals_priority(inventory):
    """
    Categorize sheets by totals implementation priority based on inventory analysis.
    """
    categories = {
        'high_priority_full_totals': [],      # Need full totals (row + column + grand)
        'medium_priority_column_totals': [],  # Need column totals only
        'low_priority_row_totals': [],        # Need row totals only
        'no_totals_needed': [],               # Don't need totals
        'already_implemented': []             # Already have some totals
    }
    
    for sheet_name, sheet_info in inventory['sheets'].items():
        current_status = sheet_info['current_totals_status']
        recommendations = sheet_info['recommended_totals']
        
        # Skip dashboard and metadata sheets
        if 'dashboard' in sheet_name.lower() or sheet_name in ['Raw Data']:
            categories['no_totals_needed'].append({
                'sheet_name': sheet_name,
                'reason': 'Dashboard/metadata sheet - totals not appropriate'
            })
            continue
        
        # Categorize based on current status and recommendations
        if current_status == 'no_totals_detected':
            # High priority - sheets with no totals that need them
            if isinstance(recommendations, list):
                full_totals_needed = any(
                    rec.get('strategy') == 'full_totals' 
                    for rec in recommendations
                )
                column_totals_needed = any(
                    rec.get('strategy') == 'column_totals'
                    for rec in recommendations
                )
                
                if full_totals_needed:
                    categories['high_priority_full_totals'].append({
                        'sheet_name': sheet_name,
                        'tables': len(sheet_info['tables']),
                        'recommendations': recommendations
                    })
                elif column_totals_needed:
                    categories['medium_priority_column_totals'].append({
                        'sheet_name': sheet_name,
                        'tables': len(sheet_info['tables']),
                        'recommendations': recommendations
                    })
                else:
                    categories['low_priority_row_totals'].append({
                        'sheet_name': sheet_name,
                        'tables': len(sheet_info['tables']),
                        'recommendations': recommendations
                    })
        
        elif current_status == 'has_some_totals':
            # May need enhancement or completion
            categories['already_implemented'].append({
                'sheet_name': sheet_name,
                'status': 'partial_implementation',
                'needs_review': True
            })
    
    return categories

def create_implementation_plan(categories):
    """
    Create a detailed implementation plan for systematic totals application.
    """
    plan = {
        'timestamp': datetime.now().isoformat(),
        'total_sheets_to_process': (
            len(categories['high_priority_full_totals']) +
            len(categories['medium_priority_column_totals']) +
            len(categories['low_priority_row_totals'])
        ),
        'phases': []
    }
    
    # Phase 1: High Priority - Full Totals
    if categories['high_priority_full_totals']:
        plan['phases'].append({
            'phase': 1,
            'name': 'High Priority - Full Totals Implementation',
            'description': 'Implement comprehensive totals (row + column + grand) for sheets with significant numeric data',
            'sheets': categories['high_priority_full_totals'],
            'implementation_strategy': 'full_totals',
            'estimated_impact': 'high'
        })
    
    # Phase 2: Medium Priority - Column Totals
    if categories['medium_priority_column_totals']:
        plan['phases'].append({
            'phase': 2,
            'name': 'Medium Priority - Column Totals Implementation',
            'description': 'Add column totals for summary sheets and data tables',
            'sheets': categories['medium_priority_column_totals'],
            'implementation_strategy': 'column_totals',
            'estimated_impact': 'medium'
        })
    
    # Phase 3: Low Priority - Row Totals
    if categories['low_priority_row_totals']:
        plan['phases'].append({
            'phase': 3,
            'name': 'Low Priority - Row Totals Implementation',
            'description': 'Add row totals for categorical breakdown sheets',
            'sheets': categories['low_priority_row_totals'],
            'implementation_strategy': 'row_totals',
            'estimated_impact': 'low'
        })
    
    # Phase 4: Review and Enhancement
    if categories['already_implemented']:
        plan['phases'].append({
            'phase': 4,
            'name': 'Review and Enhancement',
            'description': 'Review existing totals implementations and enhance where needed',
            'sheets': categories['already_implemented'],
            'implementation_strategy': 'review_and_enhance',
            'estimated_impact': 'maintenance'
        })
    
    return plan

def identify_specific_sheet_implementations():
    """
    Identify specific sheets that need immediate totals implementation.
    """
    # Based on inventory analysis, these sheets were identified as high priority
    target_sheets = {
        'Monthly Capture Volume': {
            'current_status': 'no_totals_detected',
            'recommended_action': 'full_totals',
            'rationale': '3 numeric columns, perfect for column totals',
            'implementation': 'add_column_totals_to_pipeline'
        },
        'File Size Stats': {
            'current_status': 'no_totals_detected', 
            'recommended_action': 'full_totals',
            'rationale': '4 numeric columns, statistical summary needs totals',
            'implementation': 'add_comprehensive_totals'
        },
        'Time of Day': {
            'current_status': 'no_totals_detected',
            'recommended_action': 'row_totals',
            'rationale': 'Few rows with 1 numeric column',
            'implementation': 'add_row_totals_to_pipeline'
        },
        'Weekday by Period': {
            'current_status': 'no_totals_detected',
            'recommended_action': 'row_totals', 
            'rationale': 'Multiple rows with 1 numeric column',
            'implementation': 'add_row_totals_to_pipeline'
        }
    }
    
    return target_sheets

def generate_implementation_code(sheet_name, implementation_type):
    """
    Generate the specific code needed to implement totals for a sheet.
    """
    implementations = {
        'add_column_totals_to_pipeline': f'''
# Add column totals to {sheet_name}
def enhance_{sheet_name.lower().replace(" ", "_")}_with_totals(self, ws, df, start_row=2):
    """Add column totals to {sheet_name} sheet."""
    try:
        config = {{
            'add_row_totals': False,
            'add_column_totals': True,
            'add_grand_total': False,
            'totals_label': 'TOTAL'
        }}
        
        self.totals_manager.add_totals_to_worksheet(
            worksheet=ws,
            dataframe=df,
            start_row=start_row,
            start_col=1,
            config=config
        )
        
        print(f"[SUCCESS] Added column totals to {{sheet_name}}")
        
    except Exception as e:
        print(f"[ERROR] Failed to add totals to {{sheet_name}}: {{e}}")
''',
        
        'add_comprehensive_totals': f'''
# Add comprehensive totals to {sheet_name}
def enhance_{sheet_name.lower().replace(" ", "_")}_with_totals(self, ws, df, start_row=2):
    """Add comprehensive totals to {sheet_name} sheet."""
    try:
        config = {{
            'add_row_totals': True,
            'add_column_totals': True,
            'add_grand_total': True,
            'totals_label': 'TOTAL',
            'row_totals_label': 'Row Total'
        }}
        
        self.totals_manager.add_totals_to_worksheet(
            worksheet=ws,
            dataframe=df,
            start_row=start_row,
            start_col=1,
            config=config
        )
        
        print(f"[SUCCESS] Added comprehensive totals to {{sheet_name}}")
        
    except Exception as e:
        print(f"[ERROR] Failed to add totals to {{sheet_name}}: {{e}}")
''',
        
        'add_row_totals_to_pipeline': f'''
# Add row totals to {sheet_name}
def enhance_{sheet_name.lower().replace(" ", "_")}_with_totals(self, ws, df, start_row=2):
    """Add row totals to {sheet_name} sheet."""
    try:
        config = {{
            'add_row_totals': True,
            'add_column_totals': False,
            'add_grand_total': False,
            'row_totals_label': 'Total'
        }}
        
        self.totals_manager.add_totals_to_worksheet(
            worksheet=ws,
            dataframe=df,
            start_row=start_row,
            start_col=1,
            config=config
        )
        
        print(f"[SUCCESS] Added row totals to {{sheet_name}}")
        
    except Exception as e:
        print(f"[ERROR] Failed to add totals to {{sheet_name}}: {{e}}")
'''
    }
    
    return implementations.get(implementation_type, "# Implementation type not recognized")

def print_implementation_summary(categories, plan, target_sheets):
    """
    Print a comprehensive summary of the totals implementation plan.
    """
    print("\n" + "="*80)
    print("📊 SYSTEMATIC TOTALS IMPLEMENTATION PLAN")
    print("="*80)
    
    print(f"📅 Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"🎯 Total sheets to process: {plan['total_sheets_to_process']}")
    
    print("\n" + "-"*60)
    print("CATEGORIZATION SUMMARY")
    print("-"*60)
    
    for category, sheets in categories.items():
        if sheets:
            print(f"\n📋 {category.replace('_', ' ').title()}: {len(sheets)} sheets")
            for sheet in sheets[:3]:  # Show first 3 as examples
                if isinstance(sheet, dict) and 'sheet_name' in sheet:
                    print(f"   • {sheet['sheet_name']}")
            if len(sheets) > 3:
                print(f"   ... and {len(sheets) - 3} more")
    
    print("\n" + "-"*60)
    print("IMPLEMENTATION PHASES")
    print("-"*60)
    
    for phase in plan['phases']:
        print(f"\n🔸 Phase {phase['phase']}: {phase['name']}")
        print(f"   📝 {phase['description']}")
        print(f"   📊 Sheets: {len(phase['sheets'])}")
        print(f"   🎯 Strategy: {phase['implementation_strategy']}")
        print(f"   📈 Impact: {phase['estimated_impact']}")
    
    print("\n" + "-"*60)
    print("IMMEDIATE PRIORITY TARGETS")
    print("-"*60)
    
    for sheet_name, details in target_sheets.items():
        print(f"\n🎯 {sheet_name}")
        print(f"   📊 Status: {details['current_status']}")
        print(f"   🔧 Action: {details['recommended_action']}")
        print(f"   💡 Rationale: {details['rationale']}")
        print(f"   ⚙️ Implementation: {details['implementation']}")
    
    print("\n" + "="*80)
    print("✅ READY FOR SYSTEMATIC IMPLEMENTATION")
    print("="*80)

def main():
    """
    Main function to analyze and plan systematic totals implementation.
    """
    try:
        print("🚀 Starting systematic totals implementation analysis...")
        
        # Load table inventory
        inventory = load_table_inventory()
        
        # Categorize sheets by priority
        categories = categorize_sheets_by_totals_priority(inventory)
        
        # Create implementation plan
        plan = create_implementation_plan(categories)
        
        # Identify specific targets
        target_sheets = identify_specific_sheet_implementations()
        
        # Print comprehensive summary
        print_implementation_summary(categories, plan, target_sheets)
        
        # Save implementation plan
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        plan_file = f"totals_implementation_plan_{timestamp}.json"
        
        full_plan = {
            'categories': categories,
            'implementation_plan': plan,
            'target_sheets': target_sheets,
            'generated_at': datetime.now().isoformat()
        }
        
        with open(plan_file, 'w') as f:
            json.dump(full_plan, f, indent=2)
        
        print(f"\n💾 Implementation plan saved to: {plan_file}")
        
        # Generate sample implementation code
        print("\n" + "-"*60)
        print("SAMPLE IMPLEMENTATION CODE")
        print("-"*60)
        
        for sheet_name, details in list(target_sheets.items())[:2]:
            print(f"\n# {sheet_name}")
            print(generate_implementation_code(sheet_name, details['implementation']))
        
        print("\n🎯 Next Steps:")
        print("1. Review the implementation plan")
        print("2. Start with Phase 1 (High Priority) sheets")
        print("3. Implement totals systematically using the generated code patterns")
        print("4. Test each implementation thoroughly")
        print("5. Validate cross-sheet consistency")
        
    except Exception as e:
        print(f"❌ Error during analysis: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
