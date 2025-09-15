#!/usr/bin/env python3
"""
ZERF Data Automation System - With Date Range Configuration
===========================================================
Data Extraction Automation For Lam Research
===========================================================
Complete updated version with user-configurable date ranges and PGr filtering
"""

import os
import sys
import time
import subprocess
import schedule
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta
import logging
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import shutil
import json
import requests
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import configparser

# Optional date picker widget
try:
    from tkcalendar import DateEntry
    TKCALENDAR_AVAILABLE = True
except ImportError:
    TKCALENDAR_AVAILABLE = False
    print("Warning: tkcalendar not available - using basic date entry")

# SharePoint libraries (optional)
SHAREPOINT_AVAILABLE = False
try:
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.client_context import ClientContext
    SHAREPOINT_LIB = "office365"
    SHAREPOINT_AVAILABLE = True
except ImportError:
    try:
        import sharepy
        SHAREPOINT_LIB = "sharepy"
        SHAREPOINT_AVAILABLE = True
    except ImportError:
        print("Warning: No SharePoint library found")

class ZERFAutomationSystem:
    def __init__(self, config_file="zerf_config.ini"):
        self.config_file = config_file
        self.load_config()
        self.setup_logging()
        
        # Paths and settings
        self.download_folder = Path(self.config.get('Paths', 'download_folder', fallback='downloads'))
        self.vbs_script_path = Path(self.config.get('Paths', 'vbs_script', fallback='zerf_automation.vbs'))
        self.backup_folder = Path(self.config.get('Paths', 'backup_folder', fallback='backup'))
        
        # SharePoint settings
        self.sharepoint_url = self.config.get('SharePoint', 'site_url', fallback='')
        self.sharepoint_username = self.config.get('SharePoint', 'username', fallback='')
        self.sharepoint_password = self.config.get('SharePoint', 'password', fallback='')
        self.sharepoint_folder = self.config.get('SharePoint', 'folder_path', 
                                               fallback='ERF Reporting_Data Analytics & Power BI')
        
        # Date range settings
        self.start_date = self.config.get('DateRange', 'start_date', fallback='08/03/2025')
        self.end_date = self.config.get('DateRange', 'end_date', fallback=datetime.now().strftime("%m/%d/%Y"))
        
        # Timing settings
        self.run_time = self.config.get('Schedule', 'run_time', fallback='08:00')
        self.check_interval = int(self.config.get('Schedule', 'check_interval', fallback='30'))
        
        # Create directories
        self.download_folder.mkdir(exist_ok=True)
        self.backup_folder.mkdir(exist_ok=True)
        
        self.is_running = False
        self.logger.info("ZERF Automation System initialized")
    
    def load_config(self):
        """Load configuration from INI file"""
        self.config = configparser.RawConfigParser()
        
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)
        else:
            self.create_default_config()
    
    def create_default_config(self):
        """Create default configuration file"""
        self.config = configparser.RawConfigParser()
        
        self.config.add_section('Paths')
        self.config.set('Paths', 'download_folder', 'downloads')
        self.config.set('Paths', 'vbs_script', 'zerf_automation.vbs')
        self.config.set('Paths', 'backup_folder', 'backup')
        
        self.config.add_section('SharePoint')
        self.config.set('SharePoint', 'site_url', 'https://lamresearch.sharepoint.com/sites/ICELabRPMTeam-GovernanceandKPIs')
        self.config.set('SharePoint', 'username', '')
        self.config.set('SharePoint', 'password', '')
        self.config.set('SharePoint', 'folder_path', 'ERF Reporting_Data Analytics & Power BI')
        
        self.config.add_section('DateRange')
        self.config.set('DateRange', 'start_date', '08/03/2025')
        self.config.set('DateRange', 'end_date', datetime.now().strftime("%m/%d/%Y"))
        
        self.config.add_section('Schedule')
        self.config.set('Schedule', 'run_time', '08:00')
        self.config.set('Schedule', 'check_interval', '30')
        
        self.config.add_section('Settings')
        self.config.set('Settings', 'auto_start', 'true')
        self.config.set('Settings', 'cleanup_old_files', 'true')
        self.config.set('Settings', 'max_retries', '3')
        
        with open(self.config_file, 'w') as f:
            self.config.write(f)
        
        print(f"Created default config file: {self.config_file}")
    
    def setup_logging(self):
        """Setup logging system"""
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        log_file = log_dir / f"zerf_automation_{datetime.now().strftime('%Y%m%d')}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        self.logger = logging.getLogger(__name__)
    
    def get_today_filename(self, suffix="", extension=".xlsx"):
        """Generate filename with today's date"""
        today = datetime.now().strftime("%m-%d-%Y")
        return f"zerf_{today}{suffix}{extension}"
    
    def create_vbs_script(self):
        """Create VBS script with Windows dialog automation using configured date range"""
        # Use configured dates instead of hardcoded ones
        start_date = self.start_date
        end_date = self.end_date
        
        today_filename = f"zerf_{datetime.now().strftime('%m-%d-%Y')}"
        download_path = str(self.download_folder.absolute()).replace("/", "\\")
        
        vbs_content = f'''Dim SapGuiAuto, application, connection, session, WshShell
Dim downloadPath, fileName, fullPath

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children(0)
Set session = connection.Children(0)
Set WshShell = CreateObject("WScript.Shell")

downloadPath = "{download_path}"
fileName = "{today_filename}.xlsx"
fullPath = downloadPath & "\\" & fileName

' SAP Automation
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "zerf"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSP$00011-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtSP$00011-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtSP$00011-HIGH").caretPosition = 0
session.findById("wnd[0]/usr/btn%_SP$00011_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1010"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1020"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "1090"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "6100"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "6200"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6300"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtSP$00018-LOW").text = "{start_date}"
session.findById("wnd[0]/usr/ctxtSP$00018-HIGH").text = "{end_date}"
session.findById("wnd[0]/usr/ctxtSP$00018-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSP$00018-LOW").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press

' Start export process
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"

' Wait for SAP export dialog
WScript.Sleep 3000

' Handle SAP export dialog
On Error Resume Next
If Not (session.findById("wnd[1]") Is Nothing) Then
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
    WScript.Sleep 1000
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 3000
End If
On Error GoTo 0

' Handle Windows Save As dialog
Dim attempts, maxAttempts
maxAttempts = 20
attempts = 0

Do While attempts < maxAttempts
    If WshShell.AppActivate("Save As") Or WshShell.AppActivate("Export") Or WshShell.AppActivate("Save") Then
        WScript.Sleep 1000
        Exit Do
    End If
    WScript.Sleep 1000
    attempts = attempts + 1
Loop

If attempts < maxAttempts Then
    WshShell.SendKeys "^a"
    WScript.Sleep 500
    WshShell.SendKeys fullPath
    WScript.Sleep 1000
    WshShell.SendKeys "{{ENTER}}"
    WScript.Sleep 2000
    
    If WshShell.AppActivate("Confirm Save As") Then
        WScript.Sleep 500
        WshShell.SendKeys "{{ENTER}}"
    End If
End If

' Wait for file to be saved
WScript.Sleep 5000

' Ensure Excel process has finished writing the file
Dim attempts2, maxAttempts2
maxAttempts2 = 10
attempts2 = 0

Do While attempts2 < maxAttempts2
    Dim fso, fileExists
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(fullPath) Then
        On Error Resume Next
        Dim testFile
        Set testFile = fso.OpenTextFile(fullPath, 1)
        If Err.Number = 0 Then
            testFile.Close
            Exit Do
        End If
        On Error GoTo 0
    End If
    
    WScript.Sleep 2000
    attempts2 = attempts2 + 1
    Set fso = Nothing
Loop

Set WshShell = Nothing
Set session = Nothing
Set connection = Nothing
Set application = Nothing
Set SapGuiAuto = Nothing
'''
        
        with open(self.vbs_script_path, 'w') as f:
            f.write(vbs_content)
        
        self.logger.info(f"Created VBS script with date range: {start_date} to {end_date}")
        self.logger.info(f"Target file: {download_path}\\{today_filename}.xlsx")
    
    def run_vbs_script(self):
        """Execute the VBS script"""
        try:
            self.logger.info("Starting VBS script execution...")
            self.create_vbs_script()
            
            result = subprocess.run(
                ['cscript', str(self.vbs_script_path)],
                capture_output=True,
                text=True,
                timeout=600
            )
            
            if result.returncode == 0:
                self.logger.info("VBS script executed successfully")
                return True
            else:
                self.logger.error(f"VBS script failed: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            self.logger.error("VBS script timed out")
            return False
        except Exception as e:
            self.logger.error(f"Error running VBS script: {e}")
            return False
    
    def find_downloaded_file(self, max_wait_minutes=5):
        """Find the most recently downloaded Excel file, excluding temporary files"""
        try:
            search_locations = [
                self.download_folder,
                Path.home() / "Downloads",
                Path.home() / "Desktop"
            ]
            
            start_time = time.time()
            max_wait_seconds = max_wait_minutes * 60
            
            while time.time() - start_time < max_wait_seconds:
                all_excel_files = []
                
                for location in search_locations:
                    if location.exists():
                        self.logger.info(f"Searching in: {location}")
                        excel_files = list(location.glob("*.xlsx")) + list(location.glob("*.xls"))
                        
                        recent_files = []
                        current_time = time.time()
                        for file in excel_files:
                            try:
                                # Skip temporary Excel files
                                if file.name.startswith('~$'):
                                    self.logger.debug(f"Skipping temporary file: {file}")
                                    continue
                                
                                # Skip files that are currently being written to
                                try:
                                    with open(file, 'r+b'):
                                        pass
                                except (PermissionError, IOError):
                                    self.logger.debug(f"Skipping locked file: {file}")
                                    continue
                                
                                file_time = os.path.getctime(file)
                                if current_time - file_time < 3600:
                                    recent_files.append(file)
                                    self.logger.info(f"Found candidate file: {file} (age: {(current_time - file_time)/60:.1f} min)")
                            except OSError as e:
                                self.logger.debug(f"Error checking file {file}: {e}")
                                continue
                        
                        all_excel_files.extend(recent_files)
                
                if all_excel_files:
                    latest_file = max(all_excel_files, key=os.path.getctime)
                    
                    try:
                        with open(latest_file, 'rb') as test_file:
                            test_file.read(1)
                        self.logger.info(f"Selected accessible file: {latest_file}")
                        return latest_file
                    except (PermissionError, IOError) as e:
                        self.logger.warning(f"File not accessible: {latest_file} - {e}")
                        all_excel_files.remove(latest_file)
                        if all_excel_files:
                            continue
                
                time.sleep(10)
                self.logger.info(f"Waiting for files... ({(time.time() - start_time)/60:.1f} min elapsed)")
            
            self.logger.warning("No accessible Excel files found after waiting")
            return None
            
        except Exception as e:
            self.logger.error(f"Error finding downloaded file: {e}")
            return None
        
    def clean_excel_data(self, input_file_path):
        """Clean Excel data according to requirements"""
        try:
            output_filename = self.get_today_filename("_cleaned")
            output_file_path = self.download_folder / output_filename
            
            self.logger.info(f"Starting data cleaning: {input_file_path}")
            
            excel_file = pd.ExcelFile(input_file_path)
            cleaned_sheets = {}
            
            for sheet_name in excel_file.sheet_names:
                self.logger.info(f"Processing sheet: {sheet_name}")
                df = pd.read_excel(input_file_path, sheet_name=sheet_name)
                
                initial_rows = len(df)
            
            # Step 1: Create Unique_ID column
                if 'ERF Nr' in df.columns and 'Item' in df.columns:
                    erf_position = df.columns.get_loc('ERF Nr')
                    unique_id_values = df['ERF Nr'].astype(str) + '-' + df['Item'].astype(str)
                    # Insert Unique_ID column right before ERF Nr column
                    df.insert(erf_position, 'Unique_ID', unique_id_values)
                    self.logger.info("Created Unique_ID column")
                
                    # Debug: Print column names to verify Unique_ID is there
                    self.logger.info(f"Columns after Unique_ID creation: {list(df.columns)}")
                else:
                    self.logger.warning("ERF Nr or Item columns not found - cannot create Unique_ID")
            
                # Step 2: Remove duplicates
                if 'Unique_ID' in df.columns:
                    before_dedup = len(df)
                    df = df.drop_duplicates(subset=['Unique_ID'], keep='first')
                    after_dedup = len(df)
                    self.logger.info(f"Removed {before_dedup - after_dedup} duplicate rows based on Unique_ID")
            
                # Step 3: Remove specific statuses
                if 'Engineering Request Form Status' in df.columns:
                    status_to_remove = ['Draft', 'Presubmit', 'Submit']
                    before_status = len(df)
                    df = df[~df['Engineering Request Form Status'].isin(status_to_remove)]
                    after_status = len(df)
                    self.logger.info(f"Removed {before_status - after_status} rows with status: {status_to_remove}")
            
                # Step 4: Remove blank ERF Sched Line Status
                if 'ERF Sched Line Status' in df.columns:
                    before_blank = len(df)
                    df = df.dropna(subset=['ERF Sched Line Status'])
                    df = df[df['ERF Sched Line Status'] != '']
                    after_blank = len(df)
                    self.logger.info(f"Removed {before_blank - after_blank} rows with blank ERF Sched Line Status")
            
                # Step 5: Remove Indirect commodity types
                if 'Commodity Type' in df.columns:
                    before_commodity = len(df)
                    df = df[~df['Commodity Type'].astype(str).str.contains('Indirect', case=False, na=False)]
                    after_commodity = len(df)
                    self.logger.info(f"Removed {before_commodity - after_commodity} rows with Indirect commodity type")
            
                # Step 6: Keep only specific Ship-To-Plant values
                if 'Ship-To-Plant' in df.columns:
                    allowed_plants = [6100, 6200, 6300, '6100', '6200', '6300']
                    before_plant = len(df)
                    df = df[df['Ship-To-Plant'].isin(allowed_plants)]
                    after_plant = len(df)
                    self.logger.info(f"Kept only allowed plants, removed {before_plant - after_plant} rows")
            
                # Step 7: Remove W91 and Z05 from PGr column
                if 'PGr' in df.columns:
                    pgr_to_remove = ['W91', 'Z05']
                    initial_pgr_rows = len(df)
                    df = df[~df['PGr'].isin(pgr_to_remove)]
                    removed_pgr_rows = initial_pgr_rows - len(df)
                    self.logger.info(f"Removed {removed_pgr_rows} rows with PGr values W91 or Z05")
            
                final_rows = len(df)
                self.logger.info(f"Sheet {sheet_name}: {initial_rows} -> {final_rows} rows")
            
                # Final debug: Verify Unique_ID is still there
                if 'Unique_ID' in df.columns:
                    self.logger.info(f"✓ Unique_ID column preserved in final output")
                    self.logger.info(f"Sample Unique_ID values: {df['Unique_ID'].head(3).tolist()}")
                else:
                    self.logger.error("✗ Unique_ID column missing in final output!")
            
                cleaned_sheets[sheet_name] = df
        
            # Save cleaned data
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                for sheet_name, df in cleaned_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        
            self.logger.info(f"Data cleaning completed: {output_file_path}")
            return output_file_path
        
        except Exception as e:
            self.logger.error(f"Error cleaning data: {e}")
            return None
    
    def upload_to_sharepoint(self, file_path):
        """Upload file to SharePoint"""
        if not SHAREPOINT_AVAILABLE:
            self.logger.error("SharePoint library not available")
            return False
        
        try:
            self.logger.info(f"Starting SharePoint upload: {file_path}")
            
            if SHAREPOINT_LIB == "office365":
                return self._upload_office365(file_path)
            elif SHAREPOINT_LIB == "sharepy":
                return self._upload_sharepy(file_path)
                
        except Exception as e:
            self.logger.error(f"SharePoint upload failed: {e}")
            return False
    
    def _upload_office365(self, file_path):
        """Upload using Office365-REST-Python-Client"""
        try:
            credentials = UserCredential(self.sharepoint_username, self.sharepoint_password)
            ctx = ClientContext(self.sharepoint_url).with_credentials(credentials)
            
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            
            with open(file_path, 'rb') as file_content:
                file_name = Path(file_path).name
                target_folder = web.get_folder_by_server_relative_url(f"/sites/ICELabRPMTeam-GovernanceandKPIs/Shared Documents/{self.sharepoint_folder}")
                
                uploaded_file = target_folder.upload_file(file_name, file_content.read())
                ctx.execute_query()
            
            self.logger.info(f"Successfully uploaded to SharePoint: {file_name}")
            return True
            
        except Exception as e:
            self.logger.error(f"Office365 upload error: {e}")
            return False
    
    def _upload_sharepy(self, file_path):
        """Upload using sharepy library"""
        try:
            s = sharepy.connect(self.sharepoint_url, username=self.sharepoint_username, password=self.sharepoint_password)
            
            file_name = Path(file_path).name
            upload_url = f"{self.sharepoint_url}/Shared Documents/{self.sharepoint_folder}/{file_name}"
            
            with open(file_path, 'rb') as file_content:
                s.post(upload_url, files={'file': file_content})
            
            self.logger.info(f"Successfully uploaded to SharePoint: {file_name}")
            return True
            
        except Exception as e:
            self.logger.error(f"Sharepy upload error: {e}")
            return False
    
    def backup_file(self, file_path):
        """Create backup of processed files"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{Path(file_path).stem}_{timestamp}{Path(file_path).suffix}"
            backup_path = self.backup_folder / backup_name
            
            shutil.copy2(file_path, backup_path)
            self.logger.info(f"Backup created: {backup_path}")
            
        except Exception as e:
            self.logger.error(f"Backup failed: {e}")
    
    def run_full_workflow(self):
        """Execute the complete automation workflow"""
        self.logger.info("="*60)
        self.logger.info("Starting ZERF automation workflow")
        self.logger.info(f"Date range: {self.start_date} to {self.end_date}")
        self.logger.info("="*60)
        
        try:
            # Step 1: Run VBS script
            self.logger.info("Step 1: Running VBS script...")
            if not self.run_vbs_script():
                self.logger.error("VBS script failed - stopping workflow")
                return False
            
            # Step 2: Wait for file to be completely written
            self.logger.info("Step 2: Waiting for file to be completely saved...")
            time.sleep(10)
            
            # Step 3: Find downloaded file
            self.logger.info("Step 3: Searching for downloaded file...")
            downloaded_file = self.find_downloaded_file(max_wait_minutes=3)
            
            if not downloaded_file:
                self.logger.error("No downloaded file found")
                return False
            
            self.logger.info(f"Found downloaded file: {downloaded_file}")
            
            # Step 4: Clean the data
            self.logger.info("Step 4: Cleaning data...")
            cleaned_file = self.clean_excel_data(downloaded_file)
            if not cleaned_file:
                self.logger.error("Data cleaning failed")
                return False
            
            # Step 5: Upload to SharePoint
            self.logger.info("Step 5: Uploading to SharePoint...")
            if self.sharepoint_username and self.sharepoint_password:
                upload_success = self.upload_to_sharepoint(cleaned_file)
                if upload_success:
                    self.logger.info("SharePoint upload successful")
                else:
                    self.logger.warning("SharePoint upload failed")
            else:
                self.logger.warning("SharePoint credentials not configured")
            
            # Step 6: Backup files
            self.backup_file(downloaded_file)
            self.backup_file(cleaned_file)
            
            self.logger.info("="*60)
            self.logger.info("ZERF automation workflow completed successfully!")
            self.logger.info("="*60)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Workflow error: {e}")
            return False
    
    def start_scheduler(self):
        """Start the scheduled automation"""
        self.logger.info(f"Starting scheduler - will run daily at {self.run_time}")
        
        schedule.every().day.at(self.run_time).do(self.run_full_workflow)
        
        self.is_running = True
        
        while self.is_running:
            schedule.run_pending()
            time.sleep(60)
    
    def stop(self):
        """Stop the automation system"""
        self.is_running = False
        self.logger.info("ZERF automation system stopped")

class ZERFAutomationGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ZERF Automation System")
        self.root.geometry("900x800")
        
        self.automation_system = None
        self.setup_gui()
    
    def setup_gui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Configuration Tab
        config_frame = ttk.Frame(notebook)
        notebook.add(config_frame, text="Configuration")
        self.setup_config_tab(config_frame)
        
        # Control Tab
        control_frame = ttk.Frame(notebook)
        notebook.add(control_frame, text="Control")
        self.setup_control_tab(control_frame)
        
        # Logs Tab
        logs_frame = ttk.Frame(notebook)
        notebook.add(logs_frame, text="Logs")
        self.setup_logs_tab(logs_frame)
    
    def setup_config_tab(self, parent):
        # Create a scrollable frame
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Date Range Configuration
        date_frame = ttk.LabelFrame(scrollable_frame, text="Date Range Configuration", padding="10")
        date_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(date_frame, text="Start Date:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.start_date_var = tk.StringVar(value="08/03/2025")
        
        if TKCALENDAR_AVAILABLE:
            # Use DateEntry if tkcalendar is available
            try:
                start_date_obj = datetime.strptime("08/03/2025", "%m/%d/%Y").date()
                self.start_date_entry = DateEntry(date_frame, textvariable=self.start_date_var, 
                                                date_pattern='mm/dd/yyyy', width=12)
                self.start_date_entry.set_date(start_date_obj)
                self.start_date_entry.grid(row=0, column=1, pady=2, padx=5, sticky=tk.W)
            except:
                # Fallback to regular Entry
                self.start_date_entry = ttk.Entry(date_frame, textvariable=self.start_date_var, width=15)
                self.start_date_entry.grid(row=0, column=1, pady=2, padx=5, sticky=tk.W)
                ttk.Label(date_frame, text="(MM/DD/YYYY)", font=("Arial", 8)).grid(row=0, column=2, pady=2, padx=5)
        else:
            # Fallback to regular Entry if tkcalendar is not available
            self.start_date_entry = ttk.Entry(date_frame, textvariable=self.start_date_var, width=15)
            self.start_date_entry.grid(row=0, column=1, pady=2, padx=5, sticky=tk.W)
            ttk.Label(date_frame, text="(MM/DD/YYYY)", font=("Arial", 8)).grid(row=0, column=2, pady=2, padx=5)
        
        ttk.Label(date_frame, text="End Date:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.end_date_var = tk.StringVar(value=datetime.now().strftime("%m/%d/%Y"))
        
        if TKCALENDAR_AVAILABLE:
            # Use DateEntry if tkcalendar is available
            try:
                self.end_date_entry = DateEntry(date_frame, textvariable=self.end_date_var, 
                                              date_pattern='mm/dd/yyyy', width=12)
                self.end_date_entry.set_date(datetime.now().date())
                self.end_date_entry.grid(row=1, column=1, pady=2, padx=5, sticky=tk.W)
            except:
                # Fallback to regular Entry
                self.end_date_entry = ttk.Entry(date_frame, textvariable=self.end_date_var, width=15)
                self.end_date_entry.grid(row=1, column=1, pady=2, padx=5, sticky=tk.W)
                ttk.Label(date_frame, text="(MM/DD/YYYY)", font=("Arial", 8)).grid(row=1, column=2, pady=2, padx=5)
        else:
            # Fallback to regular Entry if tkcalendar is not available
            self.end_date_entry = ttk.Entry(date_frame, textvariable=self.end_date_var, width=15)
            self.end_date_entry.grid(row=1, column=1, pady=2, padx=5, sticky=tk.W)
            ttk.Label(date_frame, text="(MM/DD/YYYY)", font=("Arial", 8)).grid(row=1, column=2, pady=2, padx=5)
        
        # Add a button to set end date to today
        ttk.Button(date_frame, text="Set to Today", 
                  command=lambda: self.end_date_var.set(datetime.now().strftime("%m/%d/%Y"))).grid(row=1, column=3, pady=2, padx=5)
        
        # SharePoint Configuration
        sp_frame = ttk.LabelFrame(scrollable_frame, text="SharePoint Configuration", padding="10")
        sp_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(sp_frame, text="Site URL:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.sp_url_var = tk.StringVar()
        ttk.Entry(sp_frame, textvariable=self.sp_url_var, width=60).grid(row=0, column=1, pady=2, padx=5)
        
        ttk.Label(sp_frame, text="Username:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.sp_username_var = tk.StringVar()
        ttk.Entry(sp_frame, textvariable=self.sp_username_var, width=60).grid(row=1, column=1, pady=2, padx=5)
        
        ttk.Label(sp_frame, text="Password:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.sp_password_var = tk.StringVar()
        ttk.Entry(sp_frame, textvariable=self.sp_password_var, show="*", width=60).grid(row=2, column=1, pady=2, padx=5)
        
        # Paths Configuration
        paths_frame = ttk.LabelFrame(scrollable_frame, text="Paths Configuration", padding="10")
        paths_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(paths_frame, text="Download Folder:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.download_folder_var = tk.StringVar()
        ttk.Entry(paths_frame, textvariable=self.download_folder_var, width=50).grid(row=0, column=1, pady=2, padx=5)
        ttk.Button(paths_frame, text="Browse", command=self.browse_download_folder).grid(row=0, column=2, pady=2)
        
        # Schedule Configuration
        schedule_frame = ttk.LabelFrame(scrollable_frame, text="Schedule Configuration", padding="10")
        schedule_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(schedule_frame, text="Daily Run Time:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.run_time_var = tk.StringVar(value="08:00")
        ttk.Entry(schedule_frame, textvariable=self.run_time_var, width=10).grid(row=0, column=1, pady=2, padx=5)
        ttk.Label(schedule_frame, text="(HH:MM format)", font=("Arial", 8)).grid(row=0, column=2, pady=2, padx=5)
        
        ttk.Button(scrollable_frame, text="Save Configuration", command=self.save_config).pack(pady=10)
        
        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def setup_control_tab(self, parent):
        # Control buttons
        control_frame = ttk.Frame(parent)
        control_frame.pack(pady=20)
        
        ttk.Button(control_frame, text="Run Now", command=self.run_now).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Start Scheduler", command=self.start_scheduler).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Stop", command=self.stop_system).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Test File Detection", command=self.test_file_detection).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Process File Manually", command=self.process_file_manually).pack(side=tk.LEFT, padx=5)
        
        # Status
        status_frame = ttk.LabelFrame(parent, text="Status", padding="10")
        status_frame.pack(fill=tk.X, pady=10)
        
        self.status_var = tk.StringVar(value="Stopped")
        ttk.Label(status_frame, textvariable=self.status_var, font=("Arial", 12, "bold")).pack()
        
        # Current Date Range Display
        date_display_frame = ttk.LabelFrame(parent, text="Current Date Range", padding="10")
        date_display_frame.pack(fill=tk.X, pady=10)
        
        self.current_date_range_var = tk.StringVar(value="Not configured")
        ttk.Label(date_display_frame, textvariable=self.current_date_range_var, font=("Arial", 10)).pack()
        
        # Real-time log display
        log_frame = ttk.LabelFrame(parent, text="Real-time Logs", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        text_frame = ttk.Frame(log_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.realtime_log_text = tk.Text(text_frame, height=15, wrap=tk.WORD, font=("Consolas", 9))
        scrollbar_rt = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.realtime_log_text.yview)
        self.realtime_log_text.configure(yscrollcommand=scrollbar_rt.set)
        
        self.realtime_log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_rt.pack(side=tk.RIGHT, fill=tk.Y)
        
        ttk.Button(log_frame, text="Clear", command=self.clear_realtime_logs).pack(pady=5)
    
    def setup_logs_tab(self, parent):
        log_frame = ttk.Frame(parent)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        ttk.Button(parent, text="Refresh Logs", command=self.refresh_logs).pack(pady=5)
    
    def log_to_gui(self, message, level="INFO"):
        """Add log message to GUI real-time display"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {level}: {message}\n"
        
        self.realtime_log_text.insert(tk.END, log_entry)
        
        # Color coding
        if level == "ERROR":
            self.realtime_log_text.tag_add("error", f"end-{len(log_entry)+1}c", "end-1c")
            self.realtime_log_text.tag_config("error", foreground="red")
        elif level == "WARNING":
            self.realtime_log_text.tag_add("warning", f"end-{len(log_entry)+1}c", "end-1c")
            self.realtime_log_text.tag_config("warning", foreground="orange")
        elif level == "SUCCESS":
            self.realtime_log_text.tag_add("success", f"end-{len(log_entry)+1}c", "end-1c")
            self.realtime_log_text.tag_config("success", foreground="green")
        
        self.realtime_log_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_realtime_logs(self):
        """Clear the real-time log display"""
        self.realtime_log_text.delete(1.0, tk.END)
    
    def browse_download_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.download_folder_var.set(folder)
    
    def validate_date_format(self, date_string):
        """Validate date format MM/DD/YYYY"""
        try:
            datetime.strptime(date_string, "%m/%d/%Y")
            return True
        except ValueError:
            return False
    
    def save_config(self):
        try:
            # Validate date formats
            if not self.validate_date_format(self.start_date_var.get()):
                messagebox.showerror("Error", "Start date must be in MM/DD/YYYY format")
                return
            
            if not self.validate_date_format(self.end_date_var.get()):
                messagebox.showerror("Error", "End date must be in MM/DD/YYYY format")
                return
            
            # Validate date range
            start_date = datetime.strptime(self.start_date_var.get(), "%m/%d/%Y")
            end_date = datetime.strptime(self.end_date_var.get(), "%m/%d/%Y")
            
            if start_date > end_date:
                messagebox.showerror("Error", "Start date cannot be after end date")
                return
            
            if not self.automation_system:
                self.automation_system = ZERFAutomationSystem()
            
            # Use RawConfigParser to avoid interpolation issues
            self.automation_system.config = configparser.RawConfigParser()
            
            # Create sections
            for section in ['SharePoint', 'Paths', 'Schedule', 'Settings', 'DateRange']:
                if not self.automation_system.config.has_section(section):
                    self.automation_system.config.add_section(section)
            
            # Clean SharePoint URL
            site_url = self.sp_url_var.get().strip()
            if '/sites/' in site_url and ('?' in site_url or '/Shared' in site_url):
                # Extract base URL from full SharePoint link
                base_url = site_url.split('/Shared')[0].split('?')[0]
                if base_url.endswith('/'):
                    base_url = base_url[:-1]
                site_url = base_url
            
            # Set configuration values
            self.automation_system.config.set('SharePoint', 'site_url', site_url)
            self.automation_system.config.set('SharePoint', 'username', self.sp_username_var.get())
            self.automation_system.config.set('SharePoint', 'password', self.sp_password_var.get())
            self.automation_system.config.set('Paths', 'download_folder', self.download_folder_var.get())
            self.automation_system.config.set('Schedule', 'run_time', self.run_time_var.get())
            self.automation_system.config.set('DateRange', 'start_date', self.start_date_var.get())
            self.automation_system.config.set('DateRange', 'end_date', self.end_date_var.get())
            
            # Save to file
            with open(self.automation_system.config_file, 'w') as f:
                self.automation_system.config.write(f)
            
            # Update automation system settings
            self.automation_system.sharepoint_url = site_url
            self.automation_system.sharepoint_username = self.sp_username_var.get()
            self.automation_system.sharepoint_password = self.sp_password_var.get()
            self.automation_system.download_folder = Path(self.download_folder_var.get())
            self.automation_system.run_time = self.run_time_var.get()
            self.automation_system.start_date = self.start_date_var.get()
            self.automation_system.end_date = self.end_date_var.get()
            
            # Update current date range display
            self.current_date_range_var.set(f"From: {self.start_date_var.get()} To: {self.end_date_var.get()}")
            
            messagebox.showinfo("Success", "Configuration saved successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {e}")
    
    def load_existing_config(self):
        """Load existing configuration into GUI fields"""
        try:
            if not self.automation_system:
                self.automation_system = ZERFAutomationSystem()
            
            # Load values from config
            self.sp_url_var.set(self.automation_system.sharepoint_url)
            self.sp_username_var.set(self.automation_system.sharepoint_username)
            self.sp_password_var.set(self.automation_system.sharepoint_password)
            self.download_folder_var.set(str(self.automation_system.download_folder))
            self.run_time_var.set(self.automation_system.run_time)
            self.start_date_var.set(self.automation_system.start_date)
            self.end_date_var.set(self.automation_system.end_date)
            
            # Update current date range display
            self.current_date_range_var.set(f"From: {self.automation_system.start_date} To: {self.automation_system.end_date}")
            
        except Exception as e:
            self.log_to_gui(f"Error loading configuration: {e}", "ERROR")
    
    def test_file_detection(self):
        """Test the file detection functionality"""
        if not self.automation_system:
            messagebox.showerror("Error", "Please save configuration first!")
            return
        
        self.log_to_gui("Testing file detection...", "INFO")
        
        def test_detection():
            found_file = self.automation_system.find_downloaded_file(max_wait_minutes=1)
            if found_file:
                self.root.after(0, lambda: self.log_to_gui(f"Found file: {found_file}", "SUCCESS"))
            else:
                self.root.after(0, lambda: self.log_to_gui("No recent Excel files found", "WARNING"))
        
        threading.Thread(target=test_detection, daemon=True).start()
    
    def process_file_manually(self):
        """Allow user to manually select a file to process"""
        if not self.automation_system:
            messagebox.showerror("Error", "Please save configuration first!")
            return
        
        file_path = filedialog.askopenfilename(
            title="Select Excel file to process",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        self.log_to_gui(f"Processing selected file: {file_path}", "INFO")
        
        def process_file():
            try:
                cleaned_file = self.automation_system.clean_excel_data(file_path)
                
                if cleaned_file:
                    self.root.after(0, lambda: self.log_to_gui(f"File cleaned: {cleaned_file}", "SUCCESS"))
                    
                    if self.automation_system.sharepoint_username and self.automation_system.sharepoint_password:
                        upload_success = self.automation_system.upload_to_sharepoint(cleaned_file)
                        status = "SUCCESS" if upload_success else "WARNING"
                        msg = "SharePoint upload successful" if upload_success else "SharePoint upload failed"
                        self.root.after(0, lambda: self.log_to_gui(msg, status))
                    
                    self.automation_system.backup_file(Path(file_path))
                    self.automation_system.backup_file(cleaned_file)
                    
                    self.root.after(0, lambda: self.log_to_gui("Manual processing completed!", "SUCCESS"))
                else:
                    self.root.after(0, lambda: self.log_to_gui("File processing failed", "ERROR"))
                    
            except Exception as e:
                self.root.after(0, lambda: self.log_to_gui(f"Error processing file: {e}", "ERROR"))
        
        threading.Thread(target=process_file, daemon=True).start()
    
    def run_now(self):
        if not self.automation_system:
            messagebox.showerror("Error", "Please save configuration first!")
            return
        
        self.status_var.set("Running...")
        self.clear_realtime_logs()
        self.log_to_gui("Starting ZERF automation workflow...", "INFO")
        self.log_to_gui(f"Date range: {self.automation_system.start_date} to {self.automation_system.end_date}", "INFO")
        
        def run_workflow():
            try:
                class GUILogHandler(logging.Handler):
                    def __init__(self, gui_callback):
                        super().__init__()
                        self.gui_callback = gui_callback
                    
                    def emit(self, record):
                        msg = self.format(record)
                        level = "ERROR" if record.levelno >= 40 else "WARNING" if record.levelno >= 30 else "INFO"
                        self.gui_callback(msg, level)
                
                gui_handler = GUILogHandler(lambda msg, lvl: self.root.after(0, lambda: self.log_to_gui(msg, lvl)))
                gui_handler.setFormatter(logging.Formatter('%(message)s'))
                self.automation_system.logger.addHandler(gui_handler)
                
                success = self.automation_system.run_full_workflow()
                
                self.automation_system.logger.removeHandler(gui_handler)
                
                self.root.after(0, lambda: self.status_var.set("Completed" if success else "Failed"))
                self.root.after(0, lambda: self.log_to_gui("Workflow completed!" if success else "Workflow failed!", "SUCCESS" if success else "ERROR"))
                
            except Exception as e:
                self.root.after(0, lambda: self.log_to_gui(f"Error: {e}", "ERROR"))
                self.root.after(0, lambda: self.status_var.set("Error"))
        
        threading.Thread(target=run_workflow, daemon=True).start()
    
    def start_scheduler(self):
        if not self.automation_system:
            messagebox.showerror("Error", "Please save configuration first!")
            return
        
        self.status_var.set("Scheduled")
        self.log_to_gui(f"Scheduler started - will run daily at {self.automation_system.run_time}", "INFO")
        self.log_to_gui(f"Using date range: {self.automation_system.start_date} to {self.automation_system.end_date}", "INFO")
        
        def start_schedule():
            self.automation_system.start_scheduler()
        
        threading.Thread(target=start_schedule, daemon=True).start()
    
    def stop_system(self):
        if self.automation_system:
            self.automation_system.stop()
        self.status_var.set("Stopped")
        self.log_to_gui("System stopped", "INFO")
    
    def refresh_logs(self):
        try:
            log_files = list(Path("logs").glob("*.log"))
            if log_files:
                latest_log = max(log_files, key=os.path.getctime)
                with open(latest_log, 'r') as f:
                    content = f.read()
                    self.log_text.delete(1.0, tk.END)
                    self.log_text.insert(1.0, content)
                    self.log_text.see(tk.END)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load logs: {e}")
    
    def run(self):
        # Load existing config when GUI starts
        self.load_existing_config()
        self.root.mainloop()

def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='ZERF Data Automation System')
    parser.add_argument('--gui', action='store_true', help='Launch GUI interface')
    parser.add_argument('--background', action='store_true', help='Run in background mode')
    parser.add_argument('--run-now', action='store_true', help='Run workflow immediately')
    parser.add_argument('--start-date', help='Start date (MM/DD/YYYY format)')
    parser.add_argument('--end-date', help='End date (MM/DD/YYYY format)')
    
    args = parser.parse_args()
    
    if args.gui:
        app = ZERFAutomationGUI()
        app.run()
    elif args.run_now:
        automation = ZERFAutomationSystem()
        # Override dates if provided via command line
        if args.start_date:
            automation.start_date = args.start_date
        if args.end_date:
            automation.end_date = args.end_date
        automation.run_full_workflow()
    elif args.background:
        automation = ZERFAutomationSystem()
        try:
            automation.start_scheduler()
        except KeyboardInterrupt:
            automation.stop()
    else:
        parser.print_help()

if __name__ == "__main__":
    main()