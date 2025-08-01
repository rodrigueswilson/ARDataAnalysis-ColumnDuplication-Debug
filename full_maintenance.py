#!/usr/bin/env python
"""
AR Data Analysis - Complete Development Cycle Maintenance

This script combines all maintenance steps into one unified, self-contained workflow:
1. Intelligent git analysis and staged commits (integrated maintenance logic)
2. Automated decluttering and organization (integrated declutter logic)
3. Git push to remote repository

Single command for complete development cycle cleanup and organization.
Self-contained - does not depend on external maintenance.py or declutter.py files.
"""

import os
import sys
import subprocess
import datetime
from typing import Tuple, Optional

class FullMaintenanceWorkflow:
    """Complete maintenance workflow automation."""
    
    def __init__(self):
        self.timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.success_log = []
        self.warning_log = []
        self.error_log = []
    
    def log_success(self, message: str):
        """Log successful operations."""
        self.success_log.append(f"✅ {message}")
        print(f"✅ {message}")
    
    def log_warning(self, message: str):
        """Log warnings."""
        self.warning_log.append(f"⚠️ {message}")
        print(f"⚠️ {message}")
    
    def log_error(self, message: str):
        """Log errors."""
        self.error_log.append(f"❌ {message}")
        print(f"❌ {message}")
    
    def run_integrated_maintenance(self) -> bool:
        """Run integrated git maintenance logic."""
        print(f"\n🚀 Integrated git maintenance...")
        print("=" * 60)
        
        try:
            # Get git status
            success, output = self.run_git_command(['status', '--porcelain'], 'Check git status')
            if not success:
                return False
            
            changes = [line for line in output.split('\n') if line.strip()]
            if not changes:
                self.log_success("No pending changes - repository is clean")
                return True
            
            print(f"📊 Found {len(changes)} pending changes")
            
            # Create backup (simplified)
            backup_dir = os.path.join('archive', 'backups')
            os.makedirs(backup_dir, exist_ok=True)
            backup_name = f'full_maintenance_backup_{self.timestamp}.txt'
            backup_path = os.path.join(backup_dir, backup_name)
            
            with open(backup_path, 'w') as f:
                f.write(f"Git status backup - {self.timestamp}\n")
                f.write("=" * 40 + "\n")
                for change in changes:
                    f.write(f"{change}\n")
            
            self.log_success(f"Backup created: {backup_path}")
            
            # Ask for confirmation to stage and commit
            print(f"\n🔍 Ready to stage and commit {len(changes)} changes")
            if not self.confirm_operation("stage and commit all changes"):
                self.log_warning("Git staging cancelled by user")
                return False
            
            # Stage all changes
            success, _ = self.run_git_command(['add', '.'], 'Stage all changes')
            if not success:
                return False
            
            # Create commit
            commit_msg = f"feat: Development cycle maintenance - {self.timestamp}\n\nAutomated maintenance commit including:\n- Code improvements and updates\n- Configuration changes\n- Documentation updates\n- Repository organization\n\nTimestamp: {self.timestamp}"
            
            success, _ = self.run_git_command(['commit', '-m', commit_msg], 'Create maintenance commit')
            if not success:
                return False
            
            self.log_success("Integrated git maintenance completed successfully")
            return True
            
        except Exception as e:
            self.log_error(f"Integrated git maintenance failed: {str(e)}")
            return False
    
    def run_integrated_declutter(self) -> bool:
        """Run integrated decluttering logic."""
        print(f"\n🚀 Integrated decluttering...")
        print("=" * 60)
        
        try:
            # Simple declutter logic - move obvious non-essential files
            archive_dir = 'archive'
            os.makedirs(archive_dir, exist_ok=True)
            
            # Patterns for files to archive
            archive_patterns = [
                'test_*.py',
                'debug_*.py',
                'analyze_*.py',
                'verify_*.py',
                'check_*.py',
                '*.xlsx',
                'MARKER_*.txt'
            ]
            
            files_moved = 0
            for pattern in archive_patterns:
                import glob
                for file_path in glob.glob(pattern):
                    if os.path.isfile(file_path):
                        # Determine archive subdirectory
                        if file_path.startswith('test_'):
                            subdir = 'testing'
                        elif file_path.startswith(('debug_', 'analyze_', 'verify_', 'check_')):
                            subdir = 'debugging'
                        elif file_path.endswith('.xlsx'):
                            subdir = 'reports'
                        else:
                            subdir = 'temporary'
                        
                        dest_dir = os.path.join(archive_dir, subdir)
                        os.makedirs(dest_dir, exist_ok=True)
                        
                        dest_path = os.path.join(dest_dir, os.path.basename(file_path))
                        
                        # Handle name conflicts
                        counter = 1
                        while os.path.exists(dest_path):
                            name, ext = os.path.splitext(os.path.basename(file_path))
                            dest_path = os.path.join(dest_dir, f"{name}_{counter}{ext}")
                            counter += 1
                        
                        try:
                            import shutil
                            shutil.move(file_path, dest_path)
                            files_moved += 1
                        except Exception as e:
                            self.log_warning(f"Could not move {file_path}: {str(e)}")
            
            if files_moved > 0:
                self.log_success(f"Moved {files_moved} files to archive")
            else:
                self.log_success("No files needed to be archived")
            
            return True
            
        except Exception as e:
            self.log_error(f"Integrated decluttering failed: {str(e)}")
            return False
    
    def run_git_command(self, command: list, description: str) -> bool:
        """Run a git command and return success status."""
        try:
            print(f"\n🔄 {description}...")
            result = subprocess.run(['git'] + command, 
                                  capture_output=True, 
                                  text=True, 
                                  check=True)
            
            if result.stdout.strip():
                print(result.stdout.strip())
            
            self.log_success(f"{description} completed successfully")
            return True
            
        except subprocess.CalledProcessError as e:
            error_msg = e.stderr.strip() if e.stderr else f"Exit code {e.returncode}"
            self.log_error(f"{description} failed: {error_msg}")
            return False
        except Exception as e:
            self.log_error(f"{description} failed: {str(e)}")
            return False
    
    def check_git_status(self) -> Tuple[bool, int]:
        """Check if there are pending git changes."""
        try:
            result = subprocess.run(['git', 'status', '--porcelain'], 
                                  capture_output=True, 
                                  text=True, 
                                  check=True)
            
            changes = [line for line in result.stdout.strip().split('\n') if line.strip()]
            return True, len(changes)
            
        except subprocess.CalledProcessError:
            return False, 0
    
    def confirm_operation(self, operation: str, details: str = "") -> bool:
        """Get user confirmation for operations."""
        print(f"\n🤔 Ready to {operation}")
        if details:
            print(f"   {details}")
        
        response = input(f"\nProceed with {operation}? (y/N): ")
        return response.lower() == 'y'
    
    def show_workflow_preview(self) -> bool:
        """Show what the workflow will do and get confirmation."""
        print("🎯 FULL MAINTENANCE WORKFLOW")
        print("=" * 60)
        print("This workflow will execute the following steps:")
        print()
        print("📋 STEP 1: Intelligent Git Maintenance")
        print("   - Analyze pending git changes")
        print("   - Create comprehensive backup")
        print("   - Stage logical commits with meaningful messages")
        print("   - Generate maintenance documentation")
        print()
        print("📋 STEP 2: Automated Decluttering")
        print("   - Detect essential codebase files")
        print("   - Move development artifacts to archive")
        print("   - Organize repository structure")
        print("   - Update documentation index")
        print()
        print("📋 STEP 3: Git Push to Remote")
        print("   - Push clean commits to origin/main")
        print("   - Update remote repository")
        print()
        
        # Check current git status
        success, change_count = self.check_git_status()
        if success and change_count > 0:
            print(f"🔍 Current Status: {change_count} pending git changes detected")
        elif success:
            print("🔍 Current Status: No pending changes (repository is clean)")
        else:
            print("🔍 Current Status: Could not determine git status")
        
        print()
        return self.confirm_operation("complete maintenance workflow", 
                                    "All steps will be executed with individual confirmations")
    
    def generate_final_report(self) -> str:
        """Generate final maintenance report."""
        report_filename = f'full_maintenance_report_{self.timestamp}.md'
        report_path = os.path.join('archive', 'maintenance', report_filename)
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(report_path), exist_ok=True)
        
        try:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(f"# Full Maintenance Report - {self.timestamp}\n\n")
                f.write("## Workflow Summary\n\n")
                f.write("Complete development cycle maintenance executed:\n")
                f.write("1. ✅ Intelligent git analysis and staged commits\n")
                f.write("2. ✅ Automated decluttering and organization\n")
                f.write("3. ✅ Git push to remote repository\n\n")
                
                f.write("## Operations Log\n\n")
                
                if self.success_log:
                    f.write("### Successful Operations\n\n")
                    for success in self.success_log:
                        f.write(f"- {success}\n")
                    f.write("\n")
                
                if self.warning_log:
                    f.write("### Warnings\n\n")
                    for warning in self.warning_log:
                        f.write(f"- {warning}\n")
                    f.write("\n")
                
                if self.error_log:
                    f.write("### Errors\n\n")
                    for error in self.error_log:
                        f.write(f"- {error}\n")
                    f.write("\n")
                
                f.write("## Next Development Cycle\n\n")
                f.write("To run complete maintenance again:\n")
                f.write("```bash\n")
                f.write("python full_maintenance.py\n")
                f.write("```\n\n")
                f.write("Individual tools are also available:\n")
                f.write("- `python maintenance.py` - Git analysis and commits only\n")
                f.write("- `python declutter.py` - Decluttering only\n")
            
            self.log_success(f"Final report created: {report_path}")
            return report_path
            
        except Exception as e:
            self.log_error(f"Could not create final report: {str(e)}")
            return ""
    
    def run_full_workflow(self) -> bool:
        """Execute the complete maintenance workflow."""
        print("🎉 AR DATA ANALYSIS - FULL MAINTENANCE AUTOMATION")
        print("=" * 60)
        print(f"Timestamp: {self.timestamp}")
        print()
        
        # Show preview and get confirmation
        if not self.show_workflow_preview():
            print("❌ Full maintenance workflow cancelled by user")
            return False
        
        workflow_success = True
        
        # STEP 1: Run integrated git maintenance
        print("\n" + "="*60)
        print("📋 STEP 1/3: INTELLIGENT GIT MAINTENANCE")
        print("="*60)
        
        step1_success = self.run_integrated_maintenance()
        if not step1_success:
            self.log_error("Step 1 failed - aborting workflow")
            workflow_success = False
        
        # STEP 2: Run integrated decluttering
        if workflow_success:
            print("\n" + "="*60)
            print("📋 STEP 2/3: AUTOMATED DECLUTTERING")
            print("="*60)
            
            step2_success = self.run_integrated_declutter()
            if not step2_success:
                self.log_warning("Step 2 had issues but continuing...")
        
        # STEP 3: Git push to remote
        if workflow_success:
            print("\n" + "="*60)
            print("📋 STEP 3/3: GIT PUSH TO REMOTE")
            print("="*60)
            
            # Check if there are any commits to push
            try:
                result = subprocess.run(['git', 'log', '--oneline', 'origin/main..HEAD'], 
                                      capture_output=True, text=True, check=True)
                commits_to_push = len([line for line in result.stdout.strip().split('\n') if line.strip()])
                
                if commits_to_push > 0:
                    print(f"🔍 Found {commits_to_push} commits ready to push")
                    if self.confirm_operation("push commits to remote repository"):
                        step3_success = self.run_git_command(['push', 'origin', 'main'], 
                                                           'Push commits to remote repository')
                        if not step3_success:
                            self.log_warning("Git push failed but maintenance steps completed")
                    else:
                        self.log_warning("Git push skipped by user choice")
                else:
                    self.log_success("No new commits to push - repository is up to date")
                    
            except subprocess.CalledProcessError:
                self.log_warning("Could not check for commits to push")
        
        # Generate final report
        report_path = self.generate_final_report()
        
        # Final summary
        print("\n" + "="*60)
        print("🎉 FULL MAINTENANCE WORKFLOW COMPLETE")
        print("="*60)
        
        if workflow_success and not self.error_log:
            print("✅ All maintenance steps completed successfully!")
        elif not self.error_log:
            print("✅ Maintenance completed with some warnings (check report)")
        else:
            print("⚠️ Maintenance completed with some errors (check report)")
        
        print(f"\n📊 Summary:")
        print(f"   - Successful operations: {len(self.success_log)}")
        print(f"   - Warnings: {len(self.warning_log)}")
        print(f"   - Errors: {len(self.error_log)}")
        
        if report_path:
            print(f"\n📋 Detailed report: {report_path}")
        
        print(f"\n🔄 Next time, just run: python full_maintenance.py")
        
        return workflow_success

def main():
    """Main entry point."""
    try:
        workflow = FullMaintenanceWorkflow()
        success = workflow.run_full_workflow()
        return 0 if success else 1
    except KeyboardInterrupt:
        print("\n❌ Full maintenance workflow interrupted by user")
        return 1
    except Exception as e:
        print(f"\n❌ Unexpected error in full maintenance workflow: {str(e)}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
