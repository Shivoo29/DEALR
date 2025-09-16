"""
Control tab for ZERF Automation System GUI
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import threading

from ..utils.logger import get_logger

logger = get_logger(__name__)

class ControlTab:
    """Control tab for the GUI"""
    
    def __init__(self, parent, main_app):
        self.parent = parent
        self.main_app = main_app
        self.frame = ttk.Frame(parent)
        
        self.setup_ui()
        self.update_status_timer()
        
        logger.debug("Control tab initialized")
    
    def setup_ui(self):
        """Setup the user interface"""
        # Main container with padding
        main_container = ttk.Frame(self.frame, padding="15")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Setup sections
        self.setup_control_section(main_container)
        self.setup_status_section(main_container)
        self.setup_progress_section(main_container)
        self.setup_realtime_logs_section(main_container)
    
    def setup_control_section(self, parent):
        """Setup control buttons section"""
        control_frame = ttk.LabelFrame(parent, text="🎮 Workflow Control", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Primary action buttons
        primary_frame = ttk.Frame(control_frame)
        primary_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Run now button
        self.run_now_btn = ttk.Button(
            primary_frame,
            text="▶️ Run Workflow Now",
            command=self.run_workflow_now,
            style='Accent.TButton'
        )
        self.run_now_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Start scheduler button
        self.start_scheduler_btn = ttk.Button(
            primary_frame,
            text="⏰ Start Scheduler",
            command=self.start_scheduler
        )
        self.start_scheduler_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Stop button
        self.stop_btn = ttk.Button(
            primary_frame,
            text="⏹️ Stop System",
            command=self.stop_system,
            style='Toolbutton'
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Secondary action buttons
        secondary_frame = ttk.Frame(control_frame)
        secondary_frame.pack(fill=tk.X)
        
        # Test buttons
        ttk.Button(
            secondary_frame,
            text="🔍 Test File Detection",
            command=self.test_file_detection
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            secondary_frame,
            text="📁 Process File Manually",
            command=self.process_file_manually
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            secondary_frame,
            text="☁️ Test SharePoint",
            command=self.test_sharepoint
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            secondary_frame,
            text="🧹 Cleanup Old Files",
            command=self.cleanup_old_files
        ).pack(side=tk.LEFT)
    
    def setup_status_section(self, parent):
        """Setup system status section"""
        status_frame = ttk.LabelFrame(parent, text="📊 System Status", padding="10")
        status_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Create status grid
        status_grid = ttk.Frame(status_frame)
        status_grid.pack(fill=tk.X)
        
        # Configure grid columns
        for i in range(6):
            status_grid.grid_columnconfigure(i, weight=1)
        
        # System status
        ttk.Label(status_grid, text="System:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.system_status_label = ttk.Label(status_grid, text="🔴 Stopped", foreground="red")
        self.system_status_label.grid(row=0, column=1, sticky=tk.W, padx=(0, 15))
        
        # Scheduler status
        ttk.Label(status_grid, text="Scheduler:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.scheduler_status_label = ttk.Label(status_grid, text="🔴 Inactive", foreground="red")
        self.scheduler_status_label.grid(row=0, column=3, sticky=tk.W, padx=(0, 15))
        
        # Configuration status
        ttk.Label(status_grid, text="Config:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
        self.config_status_label = ttk.Label(status_grid, text="❓ Unknown", foreground="gray")
        self.config_status_label.grid(row=0, column=5, sticky=tk.W)
        
        # Next run time
        ttk.Label(status_grid, text="Next Run:", font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.next_run_label = ttk.Label(status_grid, text="Not scheduled", foreground="gray")
        self.next_run_label.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=(0, 15), pady=(5, 0))
        
        # Last run time
        ttk.Label(status_grid, text="Last Run:", font=('Segoe UI', 9, 'bold')).grid(row=1, column=3, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.last_run_label = ttk.Label(status_grid, text="Never", foreground="gray")
        self.last_run_label.grid(row=1, column=4, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # Current date range
        ttk.Label(status_grid, text="Date Range:", font=('Segoe UI', 9, 'bold')).grid(row=2, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.date_range_label = ttk.Label(status_grid, text="Not configured", foreground="gray")
        self.date_range_label.grid(row=2, column=1, columnspan=5, sticky=tk.W, pady=(5, 0))
    
    def setup_progress_section(self, parent):
        """Setup progress tracking section"""
        progress_frame = ttk.LabelFrame(parent, text="📈 Progress", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        # Progress label
        self.progress_label = ttk.Label(progress_frame, text="Ready", foreground="green")
        self.progress_label.pack(anchor=tk.W)
        
        # Progress details
        self.progress_details = ttk.Label(progress_frame, text="", font=('Segoe UI', 8), foreground="gray")
        self.progress_details.pack(anchor=tk.W)
    
    def setup_realtime_logs_section(self, parent):
        """Setup real-time logs section"""
        logs_frame = ttk.LabelFrame(parent, text="📋 Real-time Activity Logs", padding="10")
        logs_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create text widget with scrollbar
        text_frame = ttk.Frame(logs_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(
            text_frame,
            height=12,
            wrap=tk.WORD,
            font=('Consolas', 9),
            background='#f8f9fa',
            foreground='#212529'
        )
        
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure text tags for different log levels
        self.log_text.tag_configure("INFO", foreground="#28a745")
        self.log_text.tag_configure("WARNING", foreground="#ffc107")
        self.log_text.tag_configure("ERROR", foreground="#dc3545")
        self.log_text.tag_configure("SUCCESS", foreground="#20c997")
        self.log_text.tag_configure("DEBUG", foreground="#6c757d")
        
        # Control buttons for logs
        log_controls = ttk.Frame(logs_frame)
        log_controls.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(log_controls, text="🗑️ Clear Logs", command=self.clear_logs).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(log_controls, text="💾 Export Logs", command=self.export_logs).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(log_controls, text="🔄 Refresh", command=self.refresh_logs).pack(side=tk.LEFT)
        
        # Auto-scroll checkbox
        self.auto_scroll_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            log_controls,
            text="Auto-scroll",
            variable=self.auto_scroll_var
        ).pack(side=tk.RIGHT)
    
    # Event handlers and utility methods
    
    def run_workflow_now(self):
        """Run workflow immediately"""
        try:
            self.add_log_message("🚀 Starting workflow execution...", "INFO")
            self.update_progress("Starting workflow...", 0, 100)
            
            # Disable button during execution
            self.run_now_btn.config(state='disabled')
            
            def workflow_callback(success):
                # Re-enable button
                self.run_now_btn.config(state='normal')
                
                if success:
                    self.add_log_message("✅ Workflow completed successfully!", "SUCCESS")
                    self.update_progress("Workflow completed successfully", 100, 100)
                else:
                    self.add_log_message("❌ Workflow failed", "ERROR")
                    self.update_progress("Workflow failed", 0, 100)
            
            # Run workflow in background
            self.main_app.run_workflow_async(workflow_callback)
            
        except Exception as e:
            logger.error(f"Failed to start workflow: {e}")
            self.add_log_message(f"❌ Failed to start workflow: {e}", "ERROR")
            self.run_now_btn.config(state='normal')
    
    def start_scheduler(self):
        """Start the scheduler"""
        try:
            self.add_log_message("⏰ Starting scheduler...", "INFO")
            self.main_app.start_scheduler_async()
            self.add_log_message("✅ Scheduler started successfully", "SUCCESS")
            
        except Exception as e:
            logger.error(f"Failed to start scheduler: {e}")
            self.add_log_message(f"❌ Failed to start scheduler: {e}", "ERROR")
    
    def stop_system(self):
        """Stop the system"""
        try:
            result = messagebox.askyesno(
                "Confirm Stop",
                "Are you sure you want to stop the automation system?\n\nThis will stop any running workflows and the scheduler."
            )
            
            if result:
                self.add_log_message("⏹️ Stopping system...", "INFO")
                self.main_app.stop_system()
                self.add_log_message("✅ System stopped", "INFO")
                self.update_progress("System stopped", 0, 100)
                
        except Exception as e:
            logger.error(f"Failed to stop system: {e}")
            self.add_log_message(f"❌ Failed to stop system: {e}", "ERROR")
    
    def test_file_detection(self):
        """Test file detection functionality"""
        try:
            self.add_log_message("🔍 Testing file detection...", "INFO")
            
            def test_thread():
                try:
                    if self.main_app.automation_engine:
                        found_file = self.main_app.automation_engine.test_file_detection()
                        
                        if found_file:
                            message = f"✅ File detection test successful: {found_file.name}"
                            level = "SUCCESS"
                        else:
                            message = "⚠️ No recent files found during test"
                            level = "WARNING"
                    else:
                        message = "❌ Automation engine not available"
                        level = "ERROR"
                    
                    self.main_app.root.after(0, lambda: self.add_log_message(message, level))
                    
                except Exception as e:
                    error_msg = f"❌ File detection test failed: {e}"
                    self.main_app.root.after(0, lambda: self.add_log_message(error_msg, "ERROR"))
            
            threading.Thread(target=test_thread, daemon=True).start()
            
        except Exception as e:
            logger.error(f"File detection test error: {e}")
            self.add_log_message(f"❌ File detection test error: {e}", "ERROR")
    
    def process_file_manually(self):
        """Process a file manually"""
        try:
            file_path = filedialog.askopenfilename(
                title="Select Excel file to process",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            
            if not file_path:
                return
            
            self.add_log_message(f"📁 Processing selected file: {Path(file_path).name}", "INFO")
            
            def process_thread():
                try:
                    if self.main_app.automation_engine:
                        result = self.main_app.automation_engine.run_data_processing_only(file_path)
                        
                        if result:
                            message = f"✅ File processed successfully: {result.name}"
                            level = "SUCCESS"
                        else:
                            message = "❌ File processing failed"
                            level = "ERROR"
                    else:
                        message = "❌ Automation engine not available"
                        level = "ERROR"
                    
                    self.main_app.root.after(0, lambda: self.add_log_message(message, level))
                    
                except Exception as e:
                    error_msg = f"❌ File processing failed: {e}"
                    self.main_app.root.after(0, lambda: self.add_log_message(error_msg, "ERROR"))
            
            threading.Thread(target=process_thread, daemon=True).start()
            
        except Exception as e:
            logger.error(f"Manual file processing error: {e}")
            self.add_log_message(f"❌ Manual processing error: {e}", "ERROR")
    
    def test_sharepoint(self):
        """Test SharePoint connection"""
        try:
            self.add_log_message("☁️ Testing SharePoint connection...", "INFO")
            
            def test_thread():
                try:
                    if self.main_app.automation_engine:
                        success = self.main_app.automation_engine.test_sharepoint_connection()
                        
                        if success:
                            message = "✅ SharePoint connection test successful"
                            level = "SUCCESS"
                        else:
                            message = "❌ SharePoint connection test failed"
                            level = "ERROR"
                    else:
                        message = "❌ Automation engine not available"
                        level = "ERROR"
                    
                    self.main_app.root.after(0, lambda: self.add_log_message(message, level))
                    
                except Exception as e:
                    error_msg = f"❌ SharePoint test failed: {e}"
                    self.main_app.root.after(0, lambda: self.add_log_message(error_msg, "ERROR"))
            
            threading.Thread(target=test_thread, daemon=True).start()
            
        except Exception as e:
            logger.error(f"SharePoint test error: {e}")
            self.add_log_message(f"❌ SharePoint test error: {e}", "ERROR")
    
    def cleanup_old_files(self):
        """Clean up old files"""
        try:
            result = messagebox.askyesno(
                "Confirm Cleanup",
                "This will remove backup and log files older than 30 days.\n\nContinue?"
            )
            
            if result:
                self.add_log_message("🧹 Cleaning up old files...", "INFO")
                
                def cleanup_thread():
                    try:
                        if self.main_app.automation_engine:
                            self.main_app.automation_engine.cleanup_old_files()
                            message = "✅ File cleanup completed"
                            level = "SUCCESS"
                        else:
                            message = "❌ Automation engine not available"
                            level = "ERROR"
                        
                        self.main_app.root.after(0, lambda: self.add_log_message(message, level))
                        
                    except Exception as e:
                        error_msg = f"❌ Cleanup failed: {e}"
                        self.main_app.root.after(0, lambda: self.add_log_message(error_msg, "ERROR"))
                
                threading.Thread(target=cleanup_thread, daemon=True).start()
                
        except Exception as e:
            logger.error(f"Cleanup error: {e}")
            self.add_log_message(f"❌ Cleanup error: {e}", "ERROR")
    
    def add_log_message(self, message: str, level: str = "INFO"):
        """Add message to real-time log display"""
        try:
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_entry = f"[{timestamp}] {message}\n"
            
            # Insert message
            self.log_text.insert(tk.END, log_entry, level)
            
            # Auto-scroll if enabled
            if self.auto_scroll_var.get():
                self.log_text.see(tk.END)
            
            # Limit log size (keep last 1000 lines)
            line_count = int(self.log_text.index('end-1c').split('.')[0])
            if line_count > 1000:
                self.log_text.delete('1.0', f'{line_count-1000}.0')
            
            # Update display
            self.log_text.update_idletasks()
            
        except Exception as e:
            logger.error(f"Failed to add log message: {e}")
    
    def update_progress(self, message: str, step: int = None, total: int = None):
        """Update progress display"""
        try:
            self.progress_label.config(text=message)
            
            if step is not None and total is not None and total > 0:
                progress_percent = (step / total) * 100
                self.progress_var.set(progress_percent)
                self.progress_details.config(text=f"Step {step} of {total} ({progress_percent:.1f}%)")
            else:
                self.progress_details.config(text="")
            
        except Exception as e:
            logger.error(f"Failed to update progress: {e}")
    
    def update_status_timer(self):
        """Update status displays periodically"""
        try:
            self.update_status_displays()
            # Schedule next update in 5 seconds
            self.main_app.root.after(5000, self.update_status_timer)
        except Exception as e:
            logger.error(f"Status update error: {e}")
    
    def update_status_displays(self):
        """Update all status displays"""
        try:
            if not self.main_app.automation_engine:
                return
            
            status = self.main_app.automation_engine.get_system_status()
            
            # System status
            if status.get('is_running'):
                self.system_status_label.config(text="🟢 Running", foreground="green")
            else:
                self.system_status_label.config(text="🔴 Stopped", foreground="red")
            
            # Scheduler status
            if status.get('scheduler_active'):
                self.scheduler_status_label.config(text="🟢 Active", foreground="green")
            else:
                self.scheduler_status_label.config(text="🔴 Inactive", foreground="red")
            
            # Configuration status
            if status.get('config_valid'):
                self.config_status_label.config(text="✅ Valid", foreground="green")
            else:
                self.config_status_label.config(text="❌ Invalid", foreground="red")
            
            # Next run time
            next_run = status.get('next_scheduled_run')
            if next_run:
                try:
                    next_run_dt = datetime.fromisoformat(next_run)
                    self.next_run_label.config(text=next_run_dt.strftime("%Y-%m-%d %H:%M"), foreground="blue")
                except:
                    self.next_run_label.config(text="Parse error", foreground="red")
            else:
                self.next_run_label.config(text="Not scheduled", foreground="gray")
            
            # Last run time
            last_run = status.get('last_run')
            if last_run:
                try:
                    last_run_dt = datetime.fromisoformat(last_run)
                    self.last_run_label.config(text=last_run_dt.strftime("%Y-%m-%d %H:%M"), foreground="black")
                except:
                    self.last_run_label.config(text="Parse error", foreground="red")
            else:
                self.last_run_label.config(text="Never", foreground="gray")
            
            # Date range
            config_manager = self.main_app.automation_engine.config_manager
            start_date = config_manager.get_start_date()
            end_date = config_manager.get_end_date()
            self.date_range_label.config(text=f"{start_date} to {end_date}", foreground="black")
            
        except Exception as e:
            logger.error(f"Failed to update status displays: {e}")
    
    def clear_logs(self):
        """Clear the log display"""
        try:
            self.log_text.delete(1.0, tk.END)
            self.add_log_message("🗑️ Log display cleared", "INFO")
        except Exception as e:
            logger.error(f"Failed to clear logs: {e}")
    
    def export_logs(self):
        """Export logs to file"""
        try:
            self.main_app.export_logs()
        except Exception as e:
            logger.error(f"Failed to export logs: {e}")
            self.add_log_message(f"❌ Failed to export logs: {e}", "ERROR")
    
    def refresh_logs(self):
        """Refresh log display"""
        try:
            self.add_log_message("🔄 Logs refreshed", "INFO")
        except Exception as e:
            logger.error(f"Failed to refresh logs: {e}")