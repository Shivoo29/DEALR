"""
Workflow scheduler for ZERF Automation System
"""

import time
import threading
from datetime import datetime, timedelta
from typing import Callable, Optional
import schedule

from ..utils.logger import get_logger
from ..utils.exceptions import WorkflowError

logger = get_logger(__name__)

class WorkflowScheduler:
    """Enhanced scheduler for automated workflow execution"""
    
    def __init__(self, config_manager, workflow_function: Callable):
        self.config_manager = config_manager
        self.workflow_function = workflow_function
        self.is_running = False
        self.scheduler_thread = None
        self.last_run_time = None
        self.last_run_success = None
        self.next_run_time = None
        
        # Schedule configuration
        self.run_time = config_manager.get_run_time()
        self.check_interval = config_manager.get_check_interval()
        
        logger.info(f"Scheduler initialized - daily run at {self.run_time}")
    
    def start(self):
        """Start the scheduler in a background thread"""
        if self.is_running:
            logger.warning("Scheduler is already running")
            return
        
        try:
            # Clear any existing schedule
            schedule.clear()
            
            # Schedule daily execution
            schedule.every().day.at(self.run_time).do(self._execute_workflow)
            
            # Calculate next run time
            self._update_next_run_time()
            
            # Start scheduler thread
            self.is_running = True
            self.scheduler_thread = threading.Thread(target=self._scheduler_loop, daemon=True)
            self.scheduler_thread.start()
            
            logger.info(f"✅ Scheduler started - next run: {self.next_run_time}")
            
        except Exception as e:
            logger.error(f"Failed to start scheduler: {e}")
            self.is_running = False
            raise
    
    def stop(self):
        """Stop the scheduler"""
        if not self.is_running:
            logger.info("Scheduler is not running")
            return
        
        try:
            self.is_running = False
            schedule.clear()
            
            # Wait for scheduler thread to finish (with timeout)
            if self.scheduler_thread and self.scheduler_thread.is_alive():
                self.scheduler_thread.join(timeout=5)
            
            logger.info("✅ Scheduler stopped")
            
        except Exception as e:
            logger.error(f"Error stopping scheduler: {e}")
    
    def _scheduler_loop(self):
        """Main scheduler loop running in background thread"""
        logger.info("Scheduler loop started")
        
        while self.is_running:
            try:
                # Run pending scheduled tasks
                schedule.run_pending()
                
                # Update next run time
                self._update_next_run_time()
                
                # Sleep for check interval
                time.sleep(self.check_interval)
                
            except Exception as e:
                logger.error(f"Scheduler loop error: {e}")
                time.sleep(60)  # Wait a minute before retrying
        
        logger.info("Scheduler loop ended")
    
    def _execute_workflow(self):
        """Execute the workflow function with error handling and logging"""
        execution_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        logger.info("="*60)
        logger.info(f"🕐 SCHEDULED WORKFLOW EXECUTION [{execution_id}]")
        logger.info("="*60)
        
        start_time = datetime.now()
        
        try:
            # Execute the workflow
            success = self.workflow_function()
            
            end_time = datetime.now()
            duration = end_time - start_time
            
            # Record execution results
            self.last_run_time = start_time
            self.last_run_success = success
            
            if success:
                logger.info("="*60)
                logger.info(f"✅ SCHEDULED WORKFLOW COMPLETED SUCCESSFULLY [{execution_id}]")
                logger.info(f"⏱️  Duration: {duration.total_seconds():.1f} seconds")
                logger.info("="*60)
            else:
                logger.error("="*60)
                logger.error(f"❌ SCHEDULED WORKFLOW FAILED [{execution_id}]")
                logger.error(f"⏱️  Duration: {duration.total_seconds():.1f} seconds")
                logger.error("="*60)
            
            # Send notification if configured
            self._send_notification(success, duration, execution_id)
            
        except Exception as e:
            end_time = datetime.now()
            duration = end_time - start_time
            
            self.last_run_time = start_time
            self.last_run_success = False
            
            logger.error("="*60)
            logger.error(f"❌ SCHEDULED WORKFLOW CRASHED [{execution_id}]")
            logger.error(f"💥 Error: {e}")
            logger.error(f"⏱️  Duration: {duration.total_seconds():.1f} seconds")
            logger.error("="*60, exc_info=True)
            
            # Send error notification
            self._send_notification(False, duration, execution_id, error=str(e))
    
    def _update_next_run_time(self):
        """Update the next scheduled run time"""
        try:
            jobs = schedule.get_jobs()
            if jobs:
                self.next_run_time = min(job.next_run for job in jobs)
            else:
                # Calculate next run time manually
                now = datetime.now()
                run_hour, run_minute = map(int, self.run_time.split(':'))
                
                # Next run today or tomorrow
                next_run = now.replace(hour=run_hour, minute=run_minute, second=0, microsecond=0)
                if next_run <= now:
                    next_run += timedelta(days=1)
                
                self.next_run_time = next_run
                
        except Exception as e:
            logger.error(f"Failed to update next run time: {e}")
            self.next_run_time = None
    
    def _send_notification(self, success: bool, duration: timedelta, execution_id: str, error: str = None):
        """Send notification about workflow execution"""
        try:
            # This is a placeholder for notification functionality
            # Could be extended to send emails, Slack messages, etc.
            
            status = "SUCCESS" if success else "FAILED"
            message = f"ZERF Automation [{execution_id}]: {status}"
            
            if not success and error:
                message += f" - {error}"
            
            message += f" (Duration: {duration.total_seconds():.1f}s)"
            
            # Log the notification (placeholder for actual notification sending)
            logger.info(f"📧 Notification: {message}")
            
            # TODO: Implement actual notification sending here
            # - Email notifications
            # - Slack/Teams integration
            # - System notifications
            
        except Exception as e:
            logger.error(f"Failed to send notification: {e}")
    
    def run_now(self) -> bool:
        """Execute workflow immediately (outside of schedule)"""
        logger.info("🚀 Manual workflow execution requested")
        
        try:
            success = self.workflow_function()
            self.last_run_time = datetime.now()
            self.last_run_success = success
            return success
            
        except Exception as e:
            logger.error(f"Manual workflow execution failed: {e}")
            self.last_run_time = datetime.now()
            self.last_run_success = False
            return False
    
    def is_active(self) -> bool:
        """Check if scheduler is currently active"""
        return self.is_running and (
            self.scheduler_thread is not None and 
            self.scheduler_thread.is_alive()
        )
    
    def get_status(self) -> dict:
        """Get comprehensive scheduler status"""
        return {
            'is_running': self.is_running,
            'is_active': self.is_active(),
            'run_time': self.run_time,
            'check_interval': self.check_interval,
            'next_run_time': self.next_run_time.isoformat() if self.next_run_time else None,
            'last_run_time': self.last_run_time.isoformat() if self.last_run_time else None,
            'last_run_success': self.last_run_success,
            'scheduled_jobs': len(schedule.get_jobs()),
            'thread_alive': self.scheduler_thread.is_alive() if self.scheduler_thread else False
        }
    
    def get_next_run_time(self) -> Optional[datetime]:
        """Get the next scheduled run time"""
        return self.next_run_time
    
    def get_time_until_next_run(self) -> Optional[timedelta]:
        """Get time remaining until next scheduled run"""
        if self.next_run_time:
            return self.next_run_time - datetime.now()
        return None
    
    def reschedule(self, new_run_time: str):
        """Reschedule the workflow to a new time"""
        try:
            # Validate time format
            hour, minute = map(int, new_run_time.split(':'))
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError("Invalid time format")
            
            # Update configuration
            self.config_manager.set('Schedule', 'run_time', new_run_time)
            self.config_manager.save_config()
            self.run_time = new_run_time
            
            # Restart scheduler with new time
            if self.is_running:
                self.stop()
                time.sleep(1)  # Brief pause
                self.start()
            
            logger.info(f"✅ Rescheduled to {new_run_time}")
            
        except Exception as e:
            logger.error(f"Failed to reschedule: {e}")
            raise
    
    def get_execution_history(self, days: int = 7) -> list:
        """Get execution history for the last N days"""
        # This is a placeholder - in a real implementation, 
        # you'd store execution history in a database or file
        
        history = []
        
        if self.last_run_time and self.last_run_success is not None:
            history.append({
                'timestamp': self.last_run_time.isoformat(),
                'success': self.last_run_success,
                'type': 'scheduled'
            })
        
        return history
    
    def cleanup_old_logs(self, days_to_keep: int = 30):
        """Clean up old scheduler logs"""
        try:
            # This would typically clean up execution logs, 
            # notification history, etc.
            logger.info(f"Cleaning up scheduler logs older than {days_to_keep} days")
            # Implementation would go here
            
        except Exception as e:
            logger.error(f"Failed to cleanup old logs: {e}")

class AdvancedScheduler(WorkflowScheduler):
    """Advanced scheduler with additional features"""
    
    def __init__(self, config_manager, workflow_function: Callable):
        super().__init__(config_manager, workflow_function)
        self.retry_count = 0
        self.max_retries = config_manager.get_max_retries()
        self.retry_delay_minutes = 15
        
    def _execute_workflow(self):
        """Enhanced workflow execution with retry logic"""
        success = False
        retry_count = 0
        
        while retry_count <= self.max_retries and not success:
            try:
                if retry_count > 0:
                    logger.info(f"🔄 Retry attempt {retry_count}/{self.max_retries}")
                    time.sleep(self.retry_delay_minutes * 60)  # Wait before retry
                
                success = self.workflow_function()
                
                if success:
                    self.retry_count = 0  # Reset retry count on success
                    logger.info("✅ Workflow completed successfully")
                else:
                    retry_count += 1
                    if retry_count <= self.max_retries:
                        logger.warning(f"⚠️ Workflow failed, will retry in {self.retry_delay_minutes} minutes")
                    
            except Exception as e:
                retry_count += 1
                logger.error(f"❌ Workflow crashed on attempt {retry_count}: {e}")
                
                if retry_count <= self.max_retries:
                    logger.info(f"🔄 Will retry in {self.retry_delay_minutes} minutes")
        
        # Record final result
        self.last_run_time = datetime.now()
        self.last_run_success = success
        self.retry_count = retry_count - 1 if not success else 0
        
        if not success:
            logger.error(f"❌ Workflow failed after {self.max_retries} retries")
    
    def set_retry_parameters(self, max_retries: int, delay_minutes: int):
        """Configure retry parameters"""
        self.max_retries = max_retries
        self.retry_delay_minutes = delay_minutes
        logger.info(f"Retry parameters updated: {max_retries} retries, {delay_minutes}min delay")