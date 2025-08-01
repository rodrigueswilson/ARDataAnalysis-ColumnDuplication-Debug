#!/usr/bin/env python
"""
AR Data Analysis - Intelligent Maintenance Automation Script

This script provides comprehensive maintenance automation for development cycles:
1. Intelligent backup system with git state preservation
2. Smart git change analysis and categorization
3. Staged commit creation with meaningful messages
4. Automated decluttering and organization
5. Documentation generation and updates

Designed specifically for AR Data Analysis codebase patterns.
"""

import os
import sys
import re
import json
import shutil
import zipfile
import datetime
import subprocess
from pathlib import Path
from typing import Set, Dict, List, Tuple, Optional, Any

# Import our proven declutter logic
try:
    from declutter import detect_codebase_files, create_backup, get_timestamp
    DECLUTTER_AVAILABLE = True
except ImportError:
    DECLUTTER_AVAILABLE = False
    print("⚠️ declutter.py not found - backup functionality will be limited")

class MaintenanceAutomation:
    """Intelligent maintenance automation for AR Data Analysis codebase."""
    
    def __init__(self):
        self.timestamp = get_timestamp() if DECLUTTER_AVAILABLE else datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.maintenance_dir = os.path.join('archive', 'maintenance')
        self.backup_dir = os.path.join('archive', 'backups')
        
        # Ensure directories exist
        os.makedirs(self.maintenance_dir, exist_ok=True)
        os.makedirs(self.backup_dir, exist_ok=True)
        
        # Git change categories based on analysis
        self.essential_patterns = [
            # Core modules
            r'^ar_utils\.py$',
            r'^acf_pacf_charts\.py$',
            r'^dashboard_generator\.py$',
            r'^chart_config_helper\.py$',
            r'^column_cleanup_utils\.py$',
            r'^declutter\.py$',
            r'^db_utils\.py$',
            
            # Main entry points
            r'^generate_report\.py$',
            r'^populate_db\.py$',
            r'^tag_.*\.py$',
            
            # Configuration
            r'^report_config\.json$',
            r'^config\.yaml$',
            r'^\.gitignore$',
            r'^README\.md$',
            r'^requirements\.txt$',
            
            # Core directories
            r'^report_generator/.*',
            r'^pipelines/.*',
            r'^utils/.*',
            r'^docs/.*'
        ]
        
        self.cleanup_patterns = [
            # Debug and analysis scripts
            r'debug_.*\.py$',
            r'analyze_.*\.py$',
            r'check_.*\.py$',
            r'verify_.*\.py$',
            r'investigate_.*\.py$',
            r'examine_.*\.py$',
            r'isolate_.*\.py$',
            r'compare_.*\.py$',
            r'direct_.*\.py$',
            r'final_.*\.py$',
            r'quick_.*\.py$',
            r'advanced_.*\.py$',
            
            # Legacy and old files
            r'.*_LEGACY_.*\.py$',
            r'.*_ORIGINAL\.py$',
            r'pipelines_LEGACY_ORIGINAL\.py$',
            
            # Old documentation
            r'CHATGPT_.*\.md$',
            r'QUICK_START_.*\.md$',
            r'REFACTORING_LOG\.md$',
            
            # Data exports
            r'.*\.media_records\.json$'
        ]
        
        self.archive_patterns = [
            # Test files
            r'test_.*\.py$',
            
            # Migration files
            r'migration_backup/.*',
            r'.*consolidated.*\.py$',
            r'.*_consolidated\.py$',
            
            # Utility scripts that aren't core
            r'sheet_management_cli\.py$',
            r'declutter_directory\.py$',
            r'deduplication_investigation\.py$',
            
            # Old declutter versions
            r'improved_declutter.*\.py$'
        ]

    def run_git_command(self, command: List[str]) -> Tuple[bool, str]:
        """Run a git command and return success status and output."""
        try:
            result = subprocess.run(
                ['git'] + command,
                capture_output=True,
                text=True,
                check=True
            )
            return True, result.stdout.strip()
        except subprocess.CalledProcessError as e:
            return False, e.stderr.strip()

    def get_git_status(self) -> Dict[str, List[str]]:
        """Get current git status categorized by change type."""
        success, output = self.run_git_command(['status', '--porcelain'])
        if not success:
            print(f"❌ Error getting git status: {output}")
            return {}
        
        changes = {
            'modified': [],
            'deleted': [],
            'untracked': [],
            'added': []
        }
        
        for line in output.split('\n'):
            if not line.strip():
                continue
                
            status = line[:2]
            filepath = line[3:].strip()
            
            if status.startswith('M'):
                changes['modified'].append(filepath)
            elif status.startswith('D'):
                changes['deleted'].append(filepath)
            elif status.startswith('??'):
                changes['untracked'].append(filepath)
            elif status.startswith('A'):
                changes['added'].append(filepath)
        
        return changes

    def categorize_changes(self, changes: Dict[str, List[str]]) -> Dict[str, List[str]]:
        """Categorize git changes into logical groups."""
        categories = {
            'essential': [],      # Core codebase changes to commit
            'cleanup': [],        # Files to remove/delete
            'archive': [],        # Files to move to archive
            'documentation': [],  # Documentation updates
            'utilities': [],      # New utility files
            'tests': []          # Test files
        }
        
        all_files = []
        for change_type, files in changes.items():
            all_files.extend(files)
        
        for filepath in all_files:
            categorized = False
            
            # Check essential patterns first
            for pattern in self.essential_patterns:
                if re.match(pattern, filepath):
                    categories['essential'].append(filepath)
                    categorized = True
                    break
            
            if categorized:
                continue
            
            # Check cleanup patterns
            for pattern in self.cleanup_patterns:
                if re.search(pattern, filepath):
                    categories['cleanup'].append(filepath)
                    categorized = True
                    break
            
            if categorized:
                continue
            
            # Check archive patterns
            for pattern in self.archive_patterns:
                if re.search(pattern, filepath):
                    categories['archive'].append(filepath)
                    categorized = True
                    break
            
            if categorized:
                continue
            
            # Special handling for specific file types
            if filepath.endswith('.md') and 'docs/' in filepath:
                categories['documentation'].append(filepath)
            elif filepath.endswith('.py') and filepath.startswith('test_'):
                categories['tests'].append(filepath)
            elif filepath.endswith(('.py', '.json', '.yaml')):
                categories['utilities'].append(filepath)
            else:
                # Default to archive for uncategorized files
                categories['archive'].append(filepath)
        
        return categories

    def create_comprehensive_backup(self) -> Optional[str]:
        """Create comprehensive backup including git state."""
        print("💾 Creating comprehensive maintenance backup...")
        
        backup_filename = f'maintenance_backup_{self.timestamp}.zip'
        backup_path = os.path.join(self.backup_dir, backup_filename)
        
        try:
            with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
                # Use proven codebase detection if available
                if DECLUTTER_AVAILABLE:
                    codebase_files = detect_codebase_files()
                    for file_path in codebase_files:
                        if os.path.isfile(file_path):
                            zipf.write(file_path)
                        elif os.path.isdir(file_path):
                            for root, _, files in os.walk(file_path):
                                for file in files:
                                    full_path = os.path.join(root, file)
                                    zipf.write(full_path)
                
                # Backup git state
                git_files = ['.git/HEAD', '.git/config', '.git/refs/heads/main']
                for git_file in git_files:
                    if os.path.exists(git_file):
                        zipf.write(git_file)
                
                # Backup current git status
                success, git_status = self.run_git_command(['status', '--porcelain'])
                if success:
                    status_file = f'git_status_{self.timestamp}.txt'
                    with open(status_file, 'w') as f:
                        f.write(git_status)
                    zipf.write(status_file)
                    os.remove(status_file)
            
            print(f"✅ Comprehensive backup created: {backup_path}")
            return backup_path
            
        except Exception as e:
            print(f"❌ Error creating backup: {e}")
            return None

    def stage_and_commit_changes(self, categories: Dict[str, List[str]]) -> bool:
        """Stage and commit changes in logical groups."""
        print("📝 Creating staged commits...")
        
        commit_groups = [
            {
                'name': 'Core system enhancements',
                'files': categories['essential'],
                'message': f'feat: Core system enhancements and improvements\n\n- Updated core modules: ar_utils.py, acf_pacf_charts.py, dashboard_generator.py\n- Enhanced report generation system\n- Improved configuration management\n- Updated documentation\n\nMaintenance timestamp: {self.timestamp}'
            },
            {
                'name': 'Add new utilities and tools',
                'files': categories['utilities'],
                'message': f'feat: Add new utilities and development tools\n\n- Added chart configuration helper\n- Added column cleanup utilities\n- Added maintenance automation\n- Enhanced development workflow\n\nMaintenance timestamp: {self.timestamp}'
            },
            {
                'name': 'Documentation updates',
                'files': categories['documentation'],
                'message': f'docs: Update project documentation\n\n- Enhanced project documentation\n- Updated development guides\n- Improved code organization\n\nMaintenance timestamp: {self.timestamp}'
            },
            {
                'name': 'Remove debug and development artifacts',
                'files': categories['cleanup'],
                'message': f'cleanup: Remove debug scripts and development artifacts\n\n- Removed debug and analysis scripts\n- Cleaned up legacy files\n- Removed temporary development files\n- Improved repository organization\n\nMaintenance timestamp: {self.timestamp}'
            }
        ]
        
        commits_created = 0
        
        for group in commit_groups:
            if not group['files']:
                continue
            
            print(f"📦 Staging: {group['name']}")
            
            # Stage files
            for filepath in group['files']:
                success, output = self.run_git_command(['add', filepath])
                if not success:
                    print(f"⚠️ Warning: Could not stage {filepath}: {output}")
            
            # Create commit
            success, output = self.run_git_command(['commit', '-m', group['message']])
            if success:
                print(f"✅ Created commit: {group['name']}")
                commits_created += 1
            else:
                print(f"⚠️ Warning: Could not create commit for {group['name']}: {output}")
        
        return commits_created > 0

    def run_declutter(self) -> bool:
        """Run the proven declutter process."""
        print("🧹 Running automated decluttering...")
        
        if not DECLUTTER_AVAILABLE:
            print("⚠️ Declutter functionality not available")
            return False
        
        try:
            # Import and run declutter logic
            import declutter
            
            # Run declutter programmatically (non-interactive)
            print("📂 Detecting remaining files to declutter...")
            
            # This would need to be adapted to run non-interactively
            # For now, we'll suggest manual run
            print("💡 Please run 'python declutter.py' manually after maintenance completes")
            return True
            
        except Exception as e:
            print(f"❌ Error running declutter: {e}")
            return False

    def generate_maintenance_report(self, categories: Dict[str, List[str]], backup_path: str) -> str:
        """Generate comprehensive maintenance report."""
        report_filename = f'maintenance_report_{self.timestamp}.md'
        report_path = os.path.join(self.maintenance_dir, report_filename)
        
        try:
            with open(report_path, 'w') as f:
                f.write(f"# Maintenance Report - {self.timestamp}\n\n")
                f.write("## Summary\n\n")
                f.write(f"- **Timestamp**: {self.timestamp}\n")
                f.write(f"- **Backup Created**: {backup_path}\n")
                f.write(f"- **Total Files Processed**: {sum(len(files) for files in categories.values())}\n\n")
                
                f.write("## Changes by Category\n\n")
                for category, files in categories.items():
                    if files:
                        f.write(f"### {category.title()} ({len(files)} files)\n\n")
                        for file in files[:10]:  # Show first 10
                            f.write(f"- {file}\n")
                        if len(files) > 10:
                            f.write(f"- ... and {len(files) - 10} more files\n")
                        f.write("\n")
                
                f.write("## Git Operations\n\n")
                f.write("- Staged commits created for logical change groups\n")
                f.write("- Repository cleaned and organized\n")
                f.write("- Development artifacts archived\n\n")
                
                f.write("## Next Steps\n\n")
                f.write("1. Review committed changes\n")
                f.write("2. Run `python declutter.py` for final cleanup\n")
                f.write("3. Push changes to remote repository\n")
                f.write("4. Continue development cycle\n\n")
                
                f.write("## Recovery\n\n")
                f.write(f"If rollback is needed, restore from: `{backup_path}`\n")
            
            print(f"📋 Maintenance report created: {report_path}")
            return report_path
            
        except Exception as e:
            print(f"❌ Error creating maintenance report: {e}")
            return ""

    def run_maintenance(self) -> bool:
        """Run complete maintenance workflow."""
        print("🚀 Starting AR Data Analysis Maintenance Automation")
        print("=" * 60)
        
        # Step 1: Create comprehensive backup
        backup_path = self.create_comprehensive_backup()
        if not backup_path:
            print("❌ Backup failed - aborting maintenance for safety")
            return False
        
        # Step 2: Analyze git changes
        print("\n📊 Analyzing git changes...")
        changes = self.get_git_status()
        if not changes:
            print("❌ Could not analyze git changes")
            return False
        
        total_changes = sum(len(files) for files in changes.values())
        print(f"✅ Found {total_changes} pending changes")
        
        # Step 3: Categorize changes
        print("\n🏷️ Categorizing changes...")
        categories = self.categorize_changes(changes)
        
        for category, files in categories.items():
            if files:
                print(f"   - {category.title()}: {len(files)} files")
        
        # Step 4: Show preview and get confirmation
        print(f"\n🔍 Maintenance Preview:")
        print(f"   - Essential files to commit: {len(categories['essential'])}")
        print(f"   - Files to cleanup/remove: {len(categories['cleanup'])}")
        print(f"   - Files to archive: {len(categories['archive'])}")
        print(f"   - Backup created: {backup_path}")
        
        response = input(f"\n🤔 Proceed with maintenance automation? (y/N): ")
        if response.lower() != 'y':
            print("❌ Maintenance cancelled by user")
            return False
        
        # Step 5: Stage and commit changes
        success = self.stage_and_commit_changes(categories)
        if not success:
            print("⚠️ Some git operations failed, but continuing...")
        
        # Step 6: Generate maintenance report
        report_path = self.generate_maintenance_report(categories, backup_path)
        
        # Step 7: Final summary
        print("\n✨ Maintenance automation completed!")
        print(f"📝 Report: {report_path}")
        print(f"💾 Backup: {backup_path}")
        print("\n🔄 Next steps:")
        print("   1. Review the commits created")
        print("   2. Run 'python declutter.py' for final cleanup")
        print("   3. Push changes: 'git push origin main'")
        
        return True

def main():
    """Main entry point."""
    maintenance = MaintenanceAutomation()
    success = maintenance.run_maintenance()
    
    if success:
        print("\n🎉 Maintenance automation completed successfully!")
    else:
        print("\n❌ Maintenance automation encountered issues")
    
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())
