"""
Configuration tab for ZERF Automation System GUI
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path

try:
    from tkcalendar import DateEntry
    TKCALENDAR_AVAILABLE = True
except ImportError:
    TKCALENDAR_AVAILABLE = False

from ..utils.logger import get_logger
from ..utils.validators import Validators

logger = get_logger(__name__)

class ConfigTab:
    """Configuration tab for the GUI"""
    
    def __init__(self, parent, main_app):
        self.parent = parent
        self.main_app = main_app
        self.frame = ttk.Frame(parent)
        
        # Configuration variables
        self.setup_variables()
        self.setup_ui()
        
        logger.debug("Configuration tab initialized")
    
    def setup_variables(self):
        """Setup tkinter variables for configuration"""
        # Date range variables
        self.start_date_var = tk.StringVar(value="08/03/2025")
        self.end_date_var = tk.StringVar(value=datetime.now().strftime("%m/%d/%Y"))
        
        # SharePoint variables
        self.sp_url_var = tk.StringVar()
        self.sp_username_var = tk.StringVar()
        self.sp_password_var = tk.StringVar()
        self.sp_folder_var = tk.StringVar(value="ERF Reporting_Data Analytics & Power BI")
        
        # Path variables
        self.download_folder_var = tk.StringVar(value="downloads")
        self.backup_folder_var = tk.StringVar(value="backup")
        
        # Schedule variables
        self.run_time_var = tk.StringVar(value="08:00")
        self.check_interval_var = tk.StringVar(value="30")
        
        # Settings variables
        self.auto_start_var = tk.BooleanVar(value=True)
        self.cleanup_old_files_var = tk.BooleanVar(value=True)
        self.max_retries_var = tk.StringVar(value="3")
        self.timeout_minutes_var = tk.StringVar(value="10")
        self.log_level_var = tk.StringVar(value="INFO")
        
        # Advanced variables
        self.use_keyring_var = tk.BooleanVar(value=True)
        self.system_notifications_var = tk.BooleanVar(value=True)
        self.debug_mode_var = tk.BooleanVar(value=False)
    
    def setup_ui(self):
        """Setup the user interface"""
        # Create scrollable frame
        canvas = tk.Canvas(self.frame)
        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Setup configuration sections
        self.setup_date_range_section()
        self.setup_sharepoint_section()
        self.setup_paths_section()
        self.setup_schedule_section()
        self.setup_settings_section()
        self.setup_advanced_section()
        self.setup_action_buttons()
        
        # Pack scrollable components
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel to canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def setup_date_range_section(self):
        """Setup date range configuration section"""
        section_frame = ttk.LabelFrame(self.scrollable_frame, text="📅 Date Range Configuration", padding="10")
        section_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Start date
        ttk.Label(section_frame, text="Start Date:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        if TKCALENDAR_AVAILABLE:
            try:
                start_date_obj = datetime.strptime("08/03/2025", "%m/%d/%Y").date()
                self.start_date_entry = DateEntry(
                    section_frame, 
                    textvariable=self.start_date_var,
                    date_pattern='mm/dd/yyyy',
                    width=12
                )
                self.start_date_entry.set_date(start_date_obj)
                self.start_date_entry.grid(row=0, column=1, pady=5, padx=(10, 5), sticky=tk.W)
            except:
                self.start_date_entry = ttk.Entry(section_frame, textvariable=self.start_date_var, width=15)
                self.start_date_entry.grid(row=0, column=1, pady=5, padx=(10, 5), sticky=tk.W)
                ttk.Label(section_frame, text="(MM/DD/YYYY)", font=("Arial", 8)).grid(row=0, column=2, pady=5)
        else:
            self.start_date_entry = ttk.Entry(section_frame, textvariable=self.start_date_var, width=15)
            self.start_date_entry.grid(row=0, column=1, pady=5, padx=(10, 5), sticky=tk.W)
            ttk.Label(section_frame, text="(MM/DD/YYYY)", font=("Arial", 8)).grid(row=0, column=2, pady=5)
        
        # End date
        ttk.Label(section_frame, text="End Date:", font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        if TKCALENDAR_AVAILABLE:
            try:
                self.end_date_entry = DateEntry(
                    section_frame,
                    textvariable=self.end_date_var,
                    date_pattern='mm/dd/yyyy',
                    width=12
                )
                self.end_date_entry.set_date(datetime.now().date())
                self.end_date_entry.grid(row=1, column=1, pady=5, padx=(10, 5), sticky=tk.W)
            except:
                self.end_date_entry = ttk.Entry(section_frame, textvariable=self.end_date_var, width=15)
                self.end_date_entry.grid(row=1, column=1, pady=5, padx=(10, 5), sticky=tk.W)
                ttk.Label(section_frame, text="(MM/DD/YYYY)", font=("Arial", 8)).grid(row=1, column=2, pady=5)
        else:
            self.end_date_entry = ttk.Entry(section_frame, textvariable=self.end_date_var, width=15)
            self.end_date_entry.grid(row=1, column=1, pady=5, padx=(10, 5), sticky=tk.W)
            ttk.Label(section_frame, text="(MM/DD/YYYY)", font=("Arial", 8)).grid(row=1, column=2, pady=5)
        
        # Set to today button
        ttk.Button(
            section_frame, 
            text="Set End Date to Today", 
            command=self.set_end_date_to_today
        ).grid(row=1, column=3, pady=5, padx=(5, 0))
        
        # Date range validation
        self.date_validation_label = ttk.Label(section_frame, text="", foreground="green")
        self.date_validation_label.grid(row=2, column=0, columnspan=4, pady=5, sticky=tk.W)
        
        # Bind validation to date changes
        self.start_date_var.trace('w', self.validate_date_range)
        self.end_date_var.trace('w', self.validate_date_range)
    
    def setup_sharepoint_section(self):
        """Setup SharePoint configuration section"""
        section_frame = ttk.LabelFrame(self.scrollable_frame, text="☁️ SharePoint Configuration", padding="10")
        section_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Site URL
        ttk.Label(section_frame, text="Site URL:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        url_entry = ttk.Entry(section_frame, textvariable=self.sp_url_var, width=70)
        url_entry.grid(row=0, column=1, columnspan=2, pady=5, padx=(10, 5), sticky=tk.EW)
        ttk.Label(section_frame, text="Example: https://company.sharepoint.com/sites/sitename", 
                 font=("Arial", 8), foreground="gray").grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=(10, 0))
        
        # Username
        ttk.Label(section_frame, text="Username:", font=('Segoe UI', 9, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5)
        username_entry = ttk.Entry(section_frame, textvariable=self.sp_username_var, width=50)
        username_entry.grid(row=2, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        
        # Password
        ttk.Label(section_frame, text="Password:", font=('Segoe UI', 9, 'bold')).grid(row=3, column=0, sticky=tk.W, pady=5)
        password_entry = ttk.Entry(section_frame, textvariable=self.sp_password_var, show="*", width=50)
        password_entry.grid(row=3, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        
        # Folder path
        ttk.Label(section_frame, text="Folder Path:", font=('Segoe UI', 9, 'bold')).grid(row=4, column=0, sticky=tk.W, pady=5)
        folder_entry = ttk.Entry(section_frame, textvariable=self.sp_folder_var, width=70)
        folder_entry.grid(row=4, column=1, columnspan=2, pady=5, padx=(10, 5), sticky=tk.EW)
        
        # Test connection button
        ttk.Button(
            section_frame, 
            text="Test SharePoint Connection", 
            command=self.test_sharepoint_connection
        ).grid(row=5, column=1, pady=10, sticky=tk.W, padx=(10, 0))
        
        # Connection status
        self.sp_status_label = ttk.Label(section_frame, text="", foreground="gray")
        self.sp_status_label.grid(row=5, column=2, pady=10, sticky=tk.W, padx=(10, 0))
        
        # Configure column weights
        section_frame.grid_columnconfigure(1, weight=1)
    
    def setup_paths_section(self):
        """Setup paths configuration section"""
        section_frame = ttk.LabelFrame(self.scrollable_frame, text="📁 Paths Configuration", padding="10")
        section_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Download folder
        ttk.Label(section_frame, text="Download Folder:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        download_entry = ttk.Entry(section_frame, textvariable=self.download_folder_var, width=50)
        download_entry.grid(row=0, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        ttk.Button(section_frame, text="Browse", command=lambda: self.browse_folder(self.download_folder_var)).grid(row=0, column=2, pady=5)
        
        # Backup folder
        ttk.Label(section_frame, text="Backup Folder:", font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5)
        backup_entry = ttk.Entry(section_frame, textvariable=self.backup_folder_var, width=50)
        backup_entry.grid(row=1, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        ttk.Button(section_frame, text="Browse", command=lambda: self.browse_folder(self.backup_folder_var)).grid(row=1, column=2, pady=5)
    
    def setup_schedule_section(self):
        """Setup schedule configuration section"""
        section_frame = ttk.LabelFrame(self.scrollable_frame, text="⏰ Schedule Configuration", padding="10")
        section_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Run time
        ttk.Label(section_frame, text="Daily Run Time:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        time_entry = ttk.Entry(section_frame, textvariable=self.run_time_var, width=10)
        time_entry.grid(row=0, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        ttk.Label(section_frame, text="(HH:MM format)", font=("Arial", 8)).grid(row=0, column=2, pady=5, padx=(5, 0))
        
        # Check interval
        ttk.Label(section_frame, text="Check Interval:", font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5)
        interval_entry = ttk.Entry(section_frame, textvariable=self.check_interval_var, width=10)
        interval_entry.grid(row=1, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        ttk.Label(section_frame, text="seconds", font=("Arial", 8)).grid(row=1, column=2, pady=5, padx=(5, 0))
    
    def setup_settings_section(self):
        """Setup general settings section"""
        section_frame = ttk.LabelFrame(self.scrollable_frame, text="⚙️ General Settings", padding="10")
        section_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Auto start
        ttk.Checkbutton(
            section_frame, 
            text="Auto-start scheduler when application starts", 
            variable=self.auto_start_var
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        # Cleanup old files
        ttk.Checkbutton(
            section_frame, 
            text="Automatically cleanup old backup files", 
            variable=self.cleanup_old_files_var
        ).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        # Max retries
        ttk.Label(section_frame, text="Max Retries:", font=('Segoe UI', 9, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5)
        retries_entry = ttk.Entry(section_frame, textvariable=self.max_retries_var, width=10)
        retries_entry.grid(row=2, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        
        # Timeout
        ttk.Label(section_frame, text="Timeout (minutes):", font=('Segoe UI', 9, 'bold')).grid(row=3, column=0, sticky=tk.W, pady=5)
        timeout_entry = ttk.Entry(section_frame, textvariable=self.timeout_minutes_var, width=10)
        timeout_entry.grid(row=3, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        
        # Log level
        ttk.Label(section_frame, text="Log Level:", font=('Segoe UI', 9, 'bold')).grid(row=4, column=0, sticky=tk.W, pady=5)
        log_level_combo = ttk.Combobox(
            section_frame, 
            textvariable=self.log_level_var, 
            values=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
            state="readonly",
            width=12
        )
        log_level_combo.grid(row=4, column=1, pady=5, padx=(10, 5), sticky=tk.W)
    
    def setup_advanced_section(self):
        """Setup advanced settings section"""
        section_frame = ttk.LabelFrame(self.scrollable_frame, text="🔧 Advanced Settings", padding="10")
        section_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Use keyring
        ttk.Checkbutton(
            section_frame, 
            text="Use secure password storage (recommended)", 
            variable=self.use_keyring_var
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        # System notifications
        ttk.Checkbutton(
            section_frame, 
            text="Enable system notifications", 
            variable=self.system_notifications_var
        ).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        # Debug mode
        ttk.Checkbutton(
            section_frame, 
            text="Enable debug mode (verbose logging)", 
            variable=self.debug_mode_var
        ).grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=2)
    
    def setup_action_buttons(self):
        """Setup action buttons"""
        button_frame = ttk.Frame(self.scrollable_frame)
        button_frame.pack(fill=tk.X, pady=20)
        
        # Save configuration
        ttk.Button(
            button_frame, 
            text="💾 Save Configuration", 
            command=self.save_configuration,
            style='Accent.TButton'
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        # Load configuration
        ttk.Button(
            button_frame, 
            text="📂 Load Configuration", 
            command=self.load_configuration
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        # Reset to defaults
        ttk.Button(
            button_frame, 
            text="🔄 Reset to Defaults", 
            command=self.reset_to_defaults
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        # Export configuration
        ttk.Button(
            button_frame, 
            text="📤 Export Config", 
            command=self.export_configuration
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        # Validate configuration
        ttk.Button(
            button_frame, 
            text="✅ Validate Config", 
            command=self.validate_configuration
        ).pack(side=tk.LEFT)
    
    # Event handlers and utility methods
    
    def set_end_date_to_today(self):
        """Set end date to today"""
        today = datetime.now().strftime("%m/%d/%Y")
        self.end_date_var.set(today)
        if TKCALENDAR_AVAILABLE and hasattr(self, 'end_date_entry'):
            try:
                self.end_date_entry.set_date(datetime.now().date())
            except:
                pass
    
    def validate_date_range(self, *args):
        """Validate date range and update status"""
        try:
            start_date = self.start_date_var.get()
            end_date = self.end_date_var.get()
            
            if not start_date or not end_date:
                self.date_validation_label.config(text="", foreground="gray")
                return
            
            # Validate formats
            if not Validators.validate_date_format(start_date):
                self.date_validation_label.config(text="❌ Invalid start date format", foreground="red")
                return
            
            if not Validators.validate_date_format(end_date):
                self.date_validation_label.config(text="❌ Invalid end date format", foreground="red")
                return
            
            # Validate range
            if not Validators.validate_date_range(start_date, end_date):
                self.date_validation_label.config(text="❌ Start date must be before end date", foreground="red")
                return
            
            # All good
            self.date_validation_label.config(text="✅ Date range is valid", foreground="green")
            
        except Exception as e:
            self.date_validation_label.config(text=f"❌ Date validation error: {e}", foreground="red")
    
    def browse_folder(self, var):
        """Browse for folder and update variable"""
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            var.set(folder)
    
    def test_sharepoint_connection(self):
        """Test SharePoint connection"""
        try:
            self.sp_status_label.config(text="🔄 Testing...", foreground="blue")
            self.parent.update()
            
            # Save current config temporarily
            self.save_configuration()
            
            # Test connection using automation engine
            if self.main_app.automation_engine:
                success = self.main_app.automation_engine.test_sharepoint_connection()
                if success:
                    self.sp_status_label.config(text="✅ Connection successful", foreground="green")
                else:
                    self.sp_status_label.config(text="❌ Connection failed", foreground="red")
            else:
                self.sp_status_label.config(text="❌ Engine not initialized", foreground="red")
                
        except Exception as e:
            logger.error(f"SharePoint connection test failed: {e}")
            self.sp_status_label.config(text=f"❌ Test failed: {e}", foreground="red")
    
    def save_configuration(self):
        """Save current configuration"""
        try:
            if not self.main_app.automation_engine:
                messagebox.showerror("Error", "Automation engine not initialized")
                return
            
            config_manager = self.main_app.automation_engine.config_manager
            
            # Update configuration from GUI
            config_manager.set('DateRange', 'start_date', self.start_date_var.get())
            config_manager.set('DateRange', 'end_date', self.end_date_var.get())
            
            # SharePoint settings with secure storage
            if self.use_keyring_var.get():
                config_manager.set_sharepoint_credentials(
                    self.sp_url_var.get(),
                    self.sp_username_var.get(),
                    self.sp_password_var.get(),
                    use_keyring=True
                )
            else:
                config_manager.set('SharePoint', 'site_url', self.sp_url_var.get())
                config_manager.set('SharePoint', 'username', self.sp_username_var.get())
                config_manager.set('SharePoint', 'password', self.sp_password_var.get())
            
            config_manager.set('SharePoint', 'folder_path', self.sp_folder_var.get())
            
            # Paths
            config_manager.set('Paths', 'download_folder', self.download_folder_var.get())
            config_manager.set('Paths', 'backup_folder', self.backup_folder_var.get())
            
            # Schedule
            config_manager.set('Schedule', 'run_time', self.run_time_var.get())
            config_manager.set('Schedule', 'check_interval', self.check_interval_var.get())
            
            # Settings
            config_manager.set('Settings', 'auto_start', str(self.auto_start_var.get()).lower())
            config_manager.set('Settings', 'cleanup_old_files', str(self.cleanup_old_files_var.get()).lower())
            config_manager.set('Settings', 'max_retries', self.max_retries_var.get())
            config_manager.set('Settings', 'timeout_minutes', self.timeout_minutes_var.get())
            config_manager.set('Settings', 'log_level', self.log_level_var.get())
            
            # Advanced settings
            config_manager.set('Security', 'use_keyring', str(self.use_keyring_var.get()).lower())
            config_manager.set('Notifications', 'system_notifications', str(self.system_notifications_var.get()).lower())
            config_manager.set('Advanced', 'debug_mode', str(self.debug_mode_var.get()).lower())
            
            # Save to file
            config_manager.save_config()
            
            messagebox.showinfo("Success", "Configuration saved successfully!")
            logger.info("Configuration saved from GUI")
            
        except Exception as e:
            logger.error(f"Failed to save configuration: {e}")
            messagebox.showerror("Error", f"Failed to save configuration:\n{e}")
    
    def load_configuration(self):
        """Load configuration into GUI"""
        try:
            if not self.main_app.automation_engine:
                logger.warning("Automation engine not initialized, using defaults")
                return
            
            config_manager = self.main_app.automation_engine.config_manager
            
            # Load date range
            self.start_date_var.set(config_manager.get_start_date())
            self.end_date_var.set(config_manager.get_end_date())
            
            # Load SharePoint settings
            self.sp_url_var.set(config_manager.get_sharepoint_url())
            self.sp_username_var.set(config_manager.get_sharepoint_username())
            self.sp_password_var.set(config_manager.get_sharepoint_password())
            self.sp_folder_var.set(config_manager.get_sharepoint_folder())
            
            # Load paths
            self.download_folder_var.set(str(config_manager.get_download_folder()))
            self.backup_folder_var.set(str(config_manager.get_backup_folder()))
            
            # Load schedule
            self.run_time_var.set(config_manager.get_run_time())
            self.check_interval_var.set(str(config_manager.get_check_interval()))
            
            # Load settings
            self.auto_start_var.set(config_manager.get_bool('Settings', 'auto_start', True))
            self.cleanup_old_files_var.set(config_manager.get_bool('Settings', 'cleanup_old_files', True))
            self.max_retries_var.set(str(config_manager.get_max_retries()))
            self.timeout_minutes_var.set(str(config_manager.get_timeout_minutes()))
            self.log_level_var.set(config_manager.get_log_level())
            
            # Load advanced settings
            self.use_keyring_var.set(config_manager.get_bool('Security', 'use_keyring', True))
            self.system_notifications_var.set(config_manager.get_bool('Notifications', 'system_notifications', True))
            self.debug_mode_var.set(config_manager.get_bool('Advanced', 'debug_mode', False))
            
            logger.info("Configuration loaded into GUI")
            
        except Exception as e:
            logger.error(f"Failed to load configuration: {e}")
            messagebox.showerror("Error", f"Failed to load configuration:\n{e}")
    
    def reset_to_defaults(self):
        """Reset configuration to defaults"""
        result = messagebox.askyesno(
            "Confirm Reset",
            "Are you sure you want to reset all settings to default values?\n\nThis will overwrite your current configuration."
        )
        
        if result:
            try:
                # Reset all variables to defaults
                self.start_date_var.set("08/03/2025")
                self.end_date_var.set(datetime.now().strftime("%m/%d/%Y"))
                self.sp_url_var.set("")
                self.sp_username_var.set("")
                self.sp_password_var.set("")
                self.sp_folder_var.set("ERF Reporting_Data Analytics & Power BI")
                self.download_folder_var.set("downloads")
                self.backup_folder_var.set("backup")
                self.run_time_var.set("08:00")
                self.check_interval_var.set("30")
                self.auto_start_var.set(True)
                self.cleanup_old_files_var.set(True)
                self.max_retries_var.set("3")
                self.timeout_minutes_var.set("10")
                self.log_level_var.set("INFO")
                self.use_keyring_var.set(True)
                self.system_notifications_var.set(True)
                self.debug_mode_var.set(False)
                
                messagebox.showinfo("Success", "Configuration reset to defaults!")
                logger.info("Configuration reset to defaults")
                
            except Exception as e:
                logger.error(f"Failed to reset configuration: {e}")
                messagebox.showerror("Error", f"Failed to reset configuration:\n{e}")
    
    def export_configuration(self):
        """Export configuration to file"""
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".ini",
                filetypes=[("INI files", "*.ini"), ("All files", "*.*")],
                title="Export Configuration"
            )
            
            if file_path and self.main_app.automation_engine:
                self.main_app.automation_engine.export_configuration(file_path, include_passwords=False)
                messagebox.showinfo("Success", f"Configuration exported to:\n{file_path}")
                
        except Exception as e:
            logger.error(f"Failed to export configuration: {e}")
            messagebox.showerror("Error", f"Failed to export configuration:\n{e}")
    
    def validate_configuration(self):
        """Validate current configuration"""
        try:
            # First validate GUI inputs
            gui_errors = []
            
            # Validate dates
            if not Validators.validate_date_format(self.start_date_var.get()):
                gui_errors.append("Invalid start date format")
            
            if not Validators.validate_date_format(self.end_date_var.get()):
                gui_errors.append("Invalid end date format")
            
            if not Validators.validate_date_range(self.start_date_var.get(), self.end_date_var.get()):
                gui_errors.append("Start date must be before end date")
            
            # Validate time format
            if not Validators.validate_time_format(self.run_time_var.get()):
                gui_errors.append("Invalid run time format (use HH:MM)")
            
            # Validate SharePoint URL if provided
            sp_url = self.sp_url_var.get().strip()
            if sp_url and sp_url != "":
                if not Validators.validate_sharepoint_url(sp_url):
                    gui_errors.append("Invalid SharePoint URL format")
            
            # Validate numeric fields
            try:
                int(self.check_interval_var.get())
            except ValueError:
                gui_errors.append("Check interval must be a number")
            
            try:
                int(self.max_retries_var.get())
            except ValueError:
                gui_errors.append("Max retries must be a number")
            
            try:
                int(self.timeout_minutes_var.get())
            except ValueError:
                gui_errors.append("Timeout must be a number")
            
            if gui_errors:
                error_msg = "Configuration validation failed:\n\n" + "\n".join(f"• {error}" for error in gui_errors)
                messagebox.showerror("Validation Failed", error_msg)
                return False
            
            # Test with automation engine if available
            if self.main_app.automation_engine:
                # Save current config temporarily for validation
                self.save_configuration()
                
                is_valid, errors = self.main_app.automation_engine.config_manager.validate_configuration()
                
                if is_valid:
                    messagebox.showinfo("Validation Success", "✅ Configuration is valid and ready to use!")
                    return True
                else:
                    error_msg = "Configuration validation failed:\n\n" + "\n".join(f"• {error}" for error in errors)
                    messagebox.showerror("Validation Failed", error_msg)
                    return False
            else:
                messagebox.showinfo("Validation Success", "✅ GUI validation passed!\n\nNote: Engine validation not available.")
                return True
                
        except Exception as e:
            logger.error(f"Configuration validation error: {e}")
            messagebox.showerror("Validation Error", f"Validation failed with error:\n{e}")
            return False