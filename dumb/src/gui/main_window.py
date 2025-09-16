"""
Main GUI window for ZERF Automation System
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from pathlib import Path

from ..core.automation_engine import ZERFAutomationEngine
from ..utils.logger import get_logger, GUILogHandler
from .config_tab import ConfigTab
from .control_tab import ControlTab
from .logs_tab import LogsTab

logger = get_logger(__name__)

class ZERFAutomationGUI:
    """Main GUI application for ZERF Automation System"""
    
    def __init__(self, config_file: str = None):
        self.config_file = config_file
        self.automation_engine = None
        self.gui_log_handler = None
        
        # Initialize GUI
        self.root = tk.Tk()
        self.setup_main_window()
        self.setup_components()
        self.initialize_engine()
        
        logger.info("ZERF Automation GUI initialized")
    
    def setup_main_window(self):
        """Setup the main window properties"""
        self.root.title("ZERF Data Automation System v2.0")
        self.root.geometry("1200x900")
        self.root.minsize(800, 600)
        
        # Set window icon (if available)
        try:
            icon_path = Path("assets/icon.ico")
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except Exception:
            pass  # Icon not critical
        
        # Configure grid weights
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')  # Modern theme
        
        # Configure styles
        style.configure('Title.TLabel', font=('Segoe UI', 12, 'bold'))
        style.configure('Header.TLabel', font=('Segoe UI', 10, 'bold'))
        style.configure('Success.TLabel', foreground='green')
        style.configure('Error.TLabel', foreground='red')
        style.configure('Warning.TLabel', foreground='orange')
    
    def setup_components(self):
        """Setup GUI components"""
        # Create main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Title bar
        self.setup_title_bar(main_frame)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        
        # Initialize tabs
        self.config_tab = ConfigTab(self.notebook, self)
        self.control_tab = ControlTab(self.notebook, self)
        self.logs_tab = LogsTab(self.notebook, self)
        
        # Add tabs to notebook
        self.notebook.add(self.config_tab.frame, text="⚙️ Configuration")
        self.notebook.add(self.control_tab.frame, text="🎮 Control")
        self.notebook.add(self.logs_tab.frame, text="📋 Logs")
        
        # Status bar
        self.setup_status_bar(main_frame)
        
        # Bind events
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_title_bar(self, parent):
        """Setup the title bar with branding"""
        title_frame = ttk.Frame(parent)
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        title_frame.grid_columnconfigure(1, weight=1)
        
        # Logo/Icon placeholder
        icon_label = ttk.Label(title_frame, text="🏭", font=('Segoe UI', 20))
        icon_label.grid(row=0, column=0, padx=(0, 10))
        
        # Title and subtitle
        title_container = ttk.Frame(title_frame)
        title_container.grid(row=0, column=1, sticky="w")
        
        title_label = ttk.Label(
            title_container, 
            text="ZERF Data Automation System", 
            style='Title.TLabel'
        )
        title_label.grid(row=0, column=0, sticky="w")
        
        subtitle_label = ttk.Label(
            title_container, 
            text="SAP Data Extraction & Processing Automation for Lam Research",
            font=('Segoe UI', 9),
            foreground='gray'
        )
        subtitle_label.grid(row=1, column=0, sticky="w")
        
        # Version info
        version_label = ttk.Label(
            title_frame, 
            text="v2.0",
            font=('Segoe UI', 9),
            foreground='gray'
        )
        version_label.grid(row=0, column=2, padx=(10, 0))
    
    def setup_status_bar(self, parent):
        """Setup the status bar"""
        self.status_frame = ttk.Frame(parent, relief='sunken', borderwidth=1)
        self.status_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.status_frame.grid_columnconfigure(1, weight=1)
        
        # Status indicator
        self.status_indicator = ttk.Label(self.status_frame, text="●", foreground="red")
        self.status_indicator.grid(row=0, column=0, padx=(5, 0))
        
        # Status text
        self.status_text = ttk.Label(self.status_frame, text="System Stopped")
        self.status_text.grid(row=0, column=1, sticky="w", padx=(5, 0))
        
        # Connection status
        self.connection_status = ttk.Label(self.status_frame, text="●", foreground="gray")
        self.connection_status.grid(row=0, column=2, padx=(0, 5))
        
        # Time display
        self.time_label = ttk.Label(self.status_frame, text="")
        self.time_label.grid(row=0, column=3, padx=(0, 5))
        
        # Update time every second
        self.update_time()
    
    def initialize_engine(self):
        """Initialize the automation engine"""
        try:
            self.automation_engine = ZERFAutomationEngine(
                config_file=self.config_file,
                progress_callback=self.update_progress
            )
            
            # Setup GUI log handler
            self.gui_log_handler = GUILogHandler(self.log_to_gui)
            self.automation_engine.sap_integration.vbs_generator.logger = get_logger(__name__)
            
            # Load initial configuration into GUI
            self.config_tab.load_configuration()
            
            # Update status
            self.update_status("Ready", "green")
            
            logger.info("Automation engine initialized successfully")
            
        except Exception as e:
            logger.error(f"Failed to initialize automation engine: {e}")
            messagebox.showerror("Initialization Error", f"Failed to initialize system:\n{e}")
            self.update_status("Error", "red")
    
    def log_to_gui(self, message: str, level: str = "INFO"):
        """Log message to GUI (called from log handler)"""
        if hasattr(self, 'control_tab'):
            self.control_tab.add_log_message(message, level)
    
    def update_progress(self, message: str, step: int = None, total: int = None):
        """Update progress display"""
        if hasattr(self, 'control_tab'):
            self.control_tab.update_progress(message, step, total)
    
    def update_status(self, status_text: str, color: str = "black"):
        """Update the status bar"""
        try:
            self.status_text.config(text=status_text)
            self.status_indicator.config(foreground=color)
        except Exception:
            pass  # Ignore errors if widgets don't exist yet
    
    def update_time(self):
        """Update the time display"""
        try:
            import datetime
            current_time = datetime.datetime.now().strftime("%H:%M:%S")
            self.time_label.config(text=current_time)
            self.root.after(1000, self.update_time)  # Update every second
        except Exception:
            pass
    
    def run_workflow_async(self, callback=None):
        """Run workflow in background thread"""
        def workflow_thread():
            try:
                self.update_status("Running Workflow...", "orange")
                
                # Add GUI log handler
                if self.gui_log_handler:
                    logger.addHandler(self.gui_log_handler)
                
                success = self.automation_engine.run_full_workflow()
                
                # Remove GUI log handler
                if self.gui_log_handler:
                    logger.removeHandler(self.gui_log_handler)
                
                # Update status on main thread
                self.root.after(0, lambda: self.update_status(
                    "Workflow Completed" if success else "Workflow Failed",
                    "green" if success else "red"
                ))
                
                # Call callback if provided
                if callback:
                    self.root.after(0, lambda: callback(success))
                    
            except Exception as e:
                logger.error(f"Workflow thread error: {e}")
                self.root.after(0, lambda: self.update_status("Workflow Error", "red"))
                if callback:
                    self.root.after(0, lambda: callback(False))
        
        thread = threading.Thread(target=workflow_thread, daemon=True)
        thread.start()
    
    def start_scheduler_async(self):
        """Start scheduler in background thread"""
        def scheduler_thread():
            try:
                self.automation_engine.start_scheduler()
            except Exception as e:
                logger.error(f"Scheduler error: {e}")
                self.root.after(0, lambda: self.update_status("Scheduler Error", "red"))
        
        thread = threading.Thread(target=scheduler_thread, daemon=True)
        thread.start()
        self.update_status("Scheduler Running", "blue")
    
    def stop_system(self):
        """Stop all operations"""
        try:
            if self.automation_engine:
                self.automation_engine.stop()
            self.update_status("System Stopped", "red")
            logger.info("System stopped by user")
        except Exception as e:
            logger.error(f"Error stopping system: {e}")
    
    def test_configuration(self) -> bool:
        """Test current configuration"""
        try:
            if not self.automation_engine:
                return False
            
            # Test configuration validation
            config_valid = self.automation_engine.validate_configuration()
            
            # Test SharePoint connection if configured
            sharepoint_valid = True
            if self.automation_engine._should_upload_to_sharepoint():
                sharepoint_valid = self.automation_engine.test_sharepoint_connection()
            
            return config_valid and sharepoint_valid
            
        except Exception as e:
            logger.error(f"Configuration test failed: {e}")
            return False
    
    def get_system_status(self) -> dict:
        """Get comprehensive system status"""
        if self.automation_engine:
            return self.automation_engine.get_system_status()
        return {'error': 'Engine not initialized'}
    
    def export_logs(self):
        """Export logs to file"""
        try:
            from tkinter import filedialog
            
            # Get recent logs
            logs = self.automation_engine.get_recent_logs(1000) if self.automation_engine else []
            
            if not logs:
                messagebox.showinfo("Export Logs", "No logs available to export")
                return
            
            # Ask user for save location
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                title="Export Logs"
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.writelines(logs)
                
                messagebox.showinfo("Export Logs", f"Logs exported to:\n{file_path}")
                logger.info(f"Logs exported to {file_path}")
                
        except Exception as e:
            logger.error(f"Failed to export logs: {e}")
            messagebox.showerror("Export Error", f"Failed to export logs:\n{e}")
    
    def show_about(self):
        """Show about dialog"""
        about_text = """
ZERF Data Automation System v2.0

A comprehensive SAP data extraction and processing 
automation tool for Lam Research.

Features:
• Automated SAP ZERF data extraction
• Advanced Excel data cleaning
• SharePoint integration
• Scheduled automation
• Comprehensive logging

© 2025 Lam Research Development Team
        """.strip()
        
        messagebox.showinfo("About ZERF Automation System", about_text)
    
    def on_closing(self):
        """Handle window closing event"""
        try:
            if self.automation_engine and self.automation_engine.get_system_status().get('is_running'):
                result = messagebox.askyesno(
                    "Confirm Exit",
                    "The automation system is currently running.\n\nAre you sure you want to exit?"
                )
                if not result:
                    return
            
            # Stop the system
            self.stop_system()
            
            # Destroy the window
            self.root.destroy()
            
            logger.info("GUI application closed")
            
        except Exception as e:
            logger.error(f"Error during closing: {e}")
            self.root.destroy()
    
    def run(self):
        """Start the GUI application"""
        try:
            logger.info("Starting ZERF Automation GUI")
            self.root.mainloop()
        except Exception as e:
            logger.error(f"GUI runtime error: {e}")
            raise

def main():
    """Main entry point for GUI application"""
    try:
        app = ZERFAutomationGUI()
        app.run()
    except Exception as e:
        logger.error(f"Failed to start GUI application: {e}")
        import tkinter.messagebox as mb
        mb.showerror("Startup Error", f"Failed to start application:\n{e}")

if __name__ == "__main__":
    main()