"""
Main automation engine for ZERF Data Automation System
"""

import time
from pathlib import Path
from datetime import datetime
from typing import Optional, Callable

from ..utils.logger import get_logger, create_progress_logger
from ..utils.config_manager import ConfigManager
from ..utils.exceptions import WorkflowError, ConfigurationError
from ..integrations.sap_integration import SAPIntegration
from ..integrations.sharepoint_client import SharePointClient
from .data_processor import DataProcessor
from .file_handler import FileHandler
from .scheduler import WorkflowScheduler

logger = get_logger(__name__)

class ZERFAutomationEngine:
    """Main automation engine that orchestrates the complete workflow"""
    
    def __init__(self, config_file: Optional[str] = None, progress_callback: Optional[Callable] = None):
        """Initialize the automation engine"""
        self.config_manager = ConfigManager(config_file)
        self.progress_callback = progress_callback
        
        # Initialize components
        self.sap_integration = SAPIntegration(self.config_manager)
        self.sharepoint_client = SharePointClient(self.config_manager)
        self.data_processor = DataProcessor(self.config_manager)
        self.file_handler = FileHandler(self.config_manager)
        self.scheduler = WorkflowScheduler(self.config_manager, self.run_full_workflow)
        
        self._is_running = False
        
        logger.info("ZERF Automation Engine initialized")
    
    def _update_progress(self, message: str, step: int = None, total: int = None):
        """Update progress via callback if available"""
        if self.progress_callback:
            self.progress_callback(message, step, total)
    
    def validate_configuration(self) -> bool:
        """Validate system configuration"""
        try:
            logger.info("Validating configuration...")
            is_valid, errors = self.config_manager.validate_configuration()
            
            if not is_valid:
                logger.error(f"Configuration validation failed: {'; '.join(errors)}")
                return False
            
            logger.info("✅ Configuration validation passed")
            return True
        except Exception as e:
            logger.error(f"Configuration validation error: {e}")
            return False
    
    def test_sharepoint_connection(self) -> bool:
        """Test SharePoint connection"""
        try:
            logger.info("Testing SharePoint connection...")
            return self.sharepoint_client.test_connection()
        except Exception as e:
            logger.error(f"SharePoint connection test failed: {e}")
            return False
    
    def run_full_workflow(self, override_dates: dict = None) -> bool:
        """Execute the complete automation workflow"""
        workflow_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        logger.info("="*60)
        logger.info(f"Starting ZERF automation workflow [{workflow_id}]")
        
        if override_dates:
            start_date = override_dates.get('start_date', self.config_manager.get_start_date())
            end_date = override_dates.get('end_date', self.config_manager.get_end_date())
        else:
            start_date = self.config_manager.get_start_date()
            end_date = self.config_manager.get_end_date()
        
        logger.info(f"Date range: {start_date} to {end_date}")
        logger.info("="*60)
        
        # Create progress tracker
        progress = create_progress_logger(logger, 7, "ZERF Workflow")
        
        try:
            self._is_running = True
            
            # Step 1: Validate configuration
            progress.step("Validating configuration...")
            self._update_progress("Validating configuration...", 1, 7)
            
            if not self.validate_configuration():
                raise WorkflowError("Configuration validation failed", "validation")
            
            # Step 2: Generate and execute VBS script
            progress.step("Executing SAP data extraction...")
            self._update_progress("Executing SAP data extraction...", 2, 7)
            
            success = self.sap_integration.extract_data(start_date, end_date)
            if not success:
                raise WorkflowError("SAP data extraction failed", "sap_extraction")
            
            # Step 3: Wait for file to be completely written
            progress.step("Waiting for file completion...")
            self._update_progress("Waiting for file completion...", 3, 7)
            time.sleep(10)  # Give SAP time to finish writing the file
            
            # Step 4: Find and validate downloaded file
            progress.step("Locating downloaded file...")
            self._update_progress("Locating downloaded file...", 4, 7)
            
            downloaded_file = self.file_handler.find_latest_download()
            if not downloaded_file:
                raise WorkflowError("No downloaded file found", "file_detection")
            
            logger.info(f"Found downloaded file: {downloaded_file}")
            
            # Step 5: Process and clean the data
            progress.step("Processing and cleaning data...")
            self._update_progress("Processing and cleaning data...", 5, 7)
            
            cleaned_file = self.data_processor.process_file(downloaded_file)
            if not cleaned_file:
                raise WorkflowError("Data processing failed", "data_processing")
            
            logger.info(f"Data cleaned successfully: {cleaned_file}")
            
            # Step 6: Upload to SharePoint (if configured)
            progress.step("Uploading to SharePoint...")
            self._update_progress("Uploading to SharePoint...", 6, 7)
            
            if self._should_upload_to_sharepoint():
                upload_success = self.sharepoint_client.upload_file(cleaned_file)
                if upload_success:
                    logger.info("✅ SharePoint upload successful")
                else:
                    logger.warning("⚠️ SharePoint upload failed (continuing workflow)")
            else:
                logger.info("⏭️ SharePoint upload skipped (not configured)")
            
            # Step 7: Create backups
            progress.step("Creating backups...")
            self._update_progress("Creating backups...", 7, 7)
            
            self.file_handler.backup_file(downloaded_file, "original")
            self.file_handler.backup_file(cleaned_file, "processed")
            
            # Workflow completed successfully
            logger.info("="*60)
            logger.info(f"✅ ZERF automation workflow [{workflow_id}] completed successfully!")
            logger.info("="*60)
            
            self._update_progress("Workflow completed successfully!", 7, 7)
            return True
            
        except WorkflowError as e:
            logger.error(f"❌ Workflow failed at {e.step}: {e}")
            if hasattr(e, 'details') and e.details:
                logger.error(f"Details: {e.details}")
            progress.error(str(e))
            return False
            
        except Exception as e:
            logger.error(f"❌ Unexpected workflow error: {e}", exc_info=True)
            progress.error(f"Unexpected error: {e}")
            return False
            
        finally:
            self._is_running = False
    
    def _should_upload_to_sharepoint(self) -> bool:
        """Check if SharePoint upload should be performed"""
        url = self.config_manager.get_sharepoint_url()
        username = self.config_manager.get_sharepoint_username()
        password = self.config_manager.get_sharepoint_password()
        
        return all([url, username, password]) and url not in ['', 'n/A']
    
    def run_data_processing_only(self, file_path: str) -> Optional[Path]:
        """Run only the data processing step for a specific file"""
        try:
            logger.info(f"Processing file: {file_path}")
            
            if not Path(file_path).exists():
                raise WorkflowError(f"File not found: {file_path}", "file_validation")
            
            # Process the file
            cleaned_file = self.data_processor.process_file(Path(file_path))
            
            if cleaned_file:
                # Create backup
                self.file_handler.backup_file(Path(file_path), "original")
                self.file_handler.backup_file(cleaned_file, "processed")
                
                logger.info(f"✅ File processing completed: {cleaned_file}")
                return cleaned_file
            else:
                logger.error("❌ File processing failed")
                return None
                
        except Exception as e:
            logger.error(f"❌ File processing error: {e}", exc_info=True)
            return None
    
    def test_file_detection(self) -> Optional[Path]:
        """Test the file detection functionality"""
        try:
            logger.info("Testing file detection...")
            found_file = self.file_handler.find_latest_download(max_wait_minutes=1)
            
            if found_file:
                logger.info(f"✅ File detection test successful: {found_file}")
                return found_file
            else:
                logger.warning("⚠️ No recent files found during test")
                return None
                
        except Exception as e:
            logger.error(f"❌ File detection test failed: {e}")
            return None
    
    def get_system_status(self) -> dict:
        """Get current system status"""
        return {
            'is_running': self._is_running,
            'scheduler_active': self.scheduler.is_active() if self.scheduler else False,
            'config_valid': self.validate_configuration(),
            'sharepoint_configured': self._should_upload_to_sharepoint(),
            'last_run': getattr(self, '_last_run', None),
            'next_scheduled_run': self.scheduler.get_next_run_time() if self.scheduler else None
        }
    
    def start_scheduler(self):
        """Start the scheduled automation"""
        try:
            logger.info(f"Starting scheduler - will run daily at {self.config_manager.get_run_time()}")
            self.scheduler.start()
        except Exception as e:
            logger.error(f"Failed to start scheduler: {e}")
            raise
    
    def stop_scheduler(self):
        """Stop the scheduled automation"""
        try:
            logger.info("Stopping scheduler...")
            if self.scheduler:
                self.scheduler.stop()
            self._is_running = False
            logger.info("✅ Scheduler stopped")
        except Exception as e:
            logger.error(f"Error stopping scheduler: {e}")
    
    def stop(self):
        """Stop all operations"""
        self.stop_scheduler()
    
    def cleanup_old_files(self, days_to_keep: int = 30):
        """Clean up old backup and log files"""
        try:
            logger.info(f"Cleaning up files older than {days_to_keep} days...")
            
            # Clean up backup files
            backup_folder = self.config_manager.get_backup_folder()
            cleaned_backups = self.file_handler.cleanup_old_files(backup_folder, days_to_keep)
            
            # Clean up log files
            log_folder = Path("logs")
            cleaned_logs = self.file_handler.cleanup_old_files(log_folder, days_to_keep)
            
            total_cleaned = cleaned_backups + cleaned_logs
            logger.info(f"✅ Cleanup completed: {total_cleaned} files removed")
            
        except Exception as e:
            logger.error(f"Cleanup failed: {e}")
    
    def export_configuration(self, export_path: str, include_passwords: bool = False):
        """Export current configuration to file"""
        try:
            self.config_manager.export_config(Path(export_path), include_passwords)
            logger.info(f"✅ Configuration exported to {export_path}")
        except Exception as e:
            logger.error(f"Failed to export configuration: {e}")
            raise
    
    def get_recent_logs(self, lines: int = 100) -> list:
        """Get recent log entries"""
        try:
            log_files = list(Path("logs").glob("*.log"))
            if not log_files:
                return []
            
            latest_log = max(log_files, key=lambda f: f.stat().st_mtime)
            
            with open(latest_log, 'r', encoding='utf-8') as f:
                all_lines = f.readlines()
                return all_lines[-lines:] if len(all_lines) > lines else all_lines
                
        except Exception as e:
            logger.error(f"Failed to read logs: {e}")
            return []