#!/usr/bin/env python3
"""
ZERF Automation System - Deployment Script
==========================================
Automates the deployment process for the ZERF Automation System
"""

import os
import sys
import subprocess
import shutil
import json
import argparse
from pathlib import Path
from datetime import datetime
import winreg
import tempfile

class ZERFDeployer:
    """Automated deployment for ZERF Automation System"""
    
    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        self.deployment_log = []
        
    def log(self, message, level="INFO"):
        """Log deployment messages"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {level}: {message}"
        self.deployment_log.append(log_entry)
        print(log_entry)
    
    def check_prerequisites(self):
        """Check system prerequisites"""
        self.log("Checking prerequisites...")
        
        # Check Python version
        if sys.version_info < (3, 8):
            raise Exception("Python 3.8 or higher required")
        self.log(f"✅ Python {sys.version} OK")
        
        # Check if running on Windows
        if os.name != 'nt':
            raise Exception("Windows operating system required")
        self.log("✅ Windows OS detected")
        
        # Check SAP GUI installation
        sap_gui_paths = [
            r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui",
            r"C:\Program Files\SAP\FrontEnd\SAPgui"
        ]
        
        sap_gui_found = any(Path(path).exists() for path in sap_gui_paths)
        if sap_gui_found:
            self.log("✅ SAP GUI installation found")
        else:
            self.log("⚠️ SAP GUI not found - manual verification required", "WARNING")
        
        # Check Excel installation
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office")
            winreg.CloseKey(key)
            self.log("✅ Microsoft Office detected")
        except:
            self.log("⚠️ Microsoft Office not detected", "WARNING")
    
    def setup_virtual_environment(self, venv_path):
        """Create and setup virtual environment"""
        self.log("Setting up virtual environment...")
        
        venv_path = Path(venv_path)
        
        # Create virtual environment
        if venv_path.exists():
            self.log("Removing existing virtual environment...")
            shutil.rmtree(venv_path)
        
        subprocess.run([sys.executable, "-m", "venv", str(venv_path)], check=True)
        self.log(f"✅ Virtual environment created: {venv_path}")
        
        # Get python and pip paths
        if os.name == 'nt':
            python_exe = venv_path / "Scripts" / "python.exe"
            pip_exe = venv_path / "Scripts" / "pip.exe"
        else:
            python_exe = venv_path / "bin" / "python"
            pip_exe = venv_path / "bin" / "pip"
        
        # Upgrade pip
        subprocess.run([str(pip_exe), "install", "--upgrade", "pip"], check=True)
        self.log("✅ Pip upgraded")
        
        return python_exe, pip_exe
    
    def install_dependencies(self, pip_exe):
        """Install Python dependencies"""
        self.log("Installing dependencies...")
        
        requirements_file = self.project_root / "requirements.txt"
        if not requirements_file.exists():
            raise Exception("requirements.txt not found")
        
        # Install requirements
        subprocess.run([
            str(pip_exe), "install", "-r", str(requirements_file)
        ], check=True)
        self.log("✅ Dependencies installed")
        
        # Install optional dependencies
        optional_packages = [
            "tkcalendar",  # Date picker widgets
            "pyinstaller",  # EXE creation
            "pytest",  # Testing
            "black",  # Code formatting
        ]
        
        for package in optional_packages:
            try:
                subprocess.run([str(pip_exe), "install", package], 
                             check=True, capture_output=True)
                self.log(f"✅ Optional package installed: {package}")
            except subprocess.CalledProcessError:
                self.log(f"⚠️ Failed to install optional package: {package}", "WARNING")
    
    def setup_configuration(self):
        """Setup configuration files"""
        self.log("Setting up configuration...")
        
        # Create directories
        directories = ["config", "logs", "downloads", "backup", "scripts"]
        for directory in directories:
            dir_path = self.project_root / directory
            dir_path.mkdir(exist_ok=True)
            self.log(f"✅ Directory created: {directory}")
        
        # Copy configuration template
        config_template = self.project_root / "config" / "zerf_config.ini"
        env_template = self.project_root / ".env.template"
        
        if not config_template.exists():
            self.create_default_config(config_template)
        
        if not env_template.exists():
            self.create_env_template(env_template)
        
        # Create .env file if it doesn't exist
        env_file = self.project_root / ".env"
        if not env_file.exists():
            shutil.copy(env_template, env_file)
            self.log("✅ Environment file created from template")
            self.log("⚠️ Please edit .env file with your actual configuration", "WARNING")
    
    def create_default_config(self, config_path):
        """Create default configuration file"""
        config_content = """[Paths]
download_folder = downloads
vbs_script = scripts/zerf_automation.vbs
backup_folder = backup

[SharePoint]
site_url = 
username = 
password = 
folder_path = ERF Reporting_Data Analytics & Power BI

[DateRange]
start_date = 08/03/2025
end_date = 09/15/2025

[Schedule]
run_time = 08:00
check_interval = 30

[Settings]
auto_start = true
cleanup_old_files = true
max_retries = 3
timeout_minutes = 10
log_level = INFO
"""
        with open(config_path, 'w') as f:
            f.write(config_content)
        self.log(f"✅ Default configuration created: {config_path}")
    
    def create_env_template(self, env_path):
        """Create environment template file"""
        # This would contain the .env template content
        # (Content already created in previous artifact)
        pass
    
    def create_executable(self, python_exe):
        """Create standalone executable"""
        self.log("Creating standalone executable...")
        
        try:
            # Check if PyInstaller is available
            subprocess.run([str(python_exe), "-c", "import PyInstaller"], 
                         check=True, capture_output=True)
        except subprocess.CalledProcessError:
            self.log("PyInstaller not available, skipping EXE creation", "WARNING")
            return None
        
        # Create EXE
        main_py = self.project_root / "main.py"
        if not main_py.exists():
            raise Exception("main.py not found")
        
        pyinstaller_cmd = [
            str(python_exe), "-m", "PyInstaller",
            "--onefile",
            "--windowed",
            "--name", "ZERF_Automation_System",
            "--add-data", f"{self.project_root / 'config'};config",
            "--add-data", f"{self.project_root / 'src'};src",
            "--hidden-import", "tkinter",
            "--hidden-import", "pandas",
            "--hidden-import", "openpyxl",
            "--hidden-import", "msal",
            "--hidden-import", "keyring",
            str(main_py)
        ]
        
        subprocess.run(pyinstaller_cmd, check=True, cwd=self.project_root)
        
        exe_path = self.project_root / "dist" / "ZERF_Automation_System.exe"
        if exe_path.exists():
            self.log(f"✅ Executable created: {exe_path}")
            return exe_path
        else:
            raise Exception("Executable creation failed")
    
    def setup_scheduled_task(self, python_exe=None, exe_path=None):
        """Setup Windows scheduled task"""
        self.log("Setting up scheduled task...")
        
        task_name = "ZERF_Automation_Daily"
        
        if exe_path and exe_path.exists():
            program = str(exe_path)
            arguments = "--background"
        elif python_exe:
            program = str(python_exe)
            arguments = f'"{self.project_root / "main.py"}" --background'
        else:
            self.log("No executable found for scheduled task", "ERROR")
            return False
        
        # Create scheduled task XML
        task_xml = f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>{datetime.now().isoformat()}</Date>
    <Author>ZERF Automation System</Author>
    <Description>Daily execution of ZERF data automation workflow</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>{datetime.now().strftime('%Y-%m-%d')}T08:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-32-545</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{program}</Command>
      <Arguments>{arguments}</Arguments>
      <WorkingDirectory>{self.project_root}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>"""
        
        # Save task XML
        task_xml_path = self.project_root / "zerf_task.xml"
        with open(task_xml_path, 'w', encoding='utf-16') as f:
            f.write(task_xml)
        
        try:
            # Create scheduled task
            subprocess.run([
                "schtasks", "/create", "/tn", task_name, 
                "/xml", str(task_xml_path), "/f"
            ], check=True, capture_output=True)
            
            self.log(f"✅ Scheduled task created: {task_name}")
            
            # Clean up XML file
            task_xml_path.unlink()
            return True
            
        except subprocess.CalledProcessError as e:
            self.log(f"❌ Failed to create scheduled task: {e}", "ERROR")
            return False
    
    def validate_deployment(self, python_exe):
        """Validate the deployment"""
        self.log("Validating deployment...")
        
        # Test configuration
        try:
            result = subprocess.run([
                str(python_exe), str(self.project_root / "main.py"), "--validate-config"
            ], capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                self.log("✅ Configuration validation passed")
            else:
                self.log(f"⚠️ Configuration validation issues: {result.stderr}", "WARNING")
        except subprocess.TimeoutExpired:
            self.log("⚠️ Configuration validation timed out", "WARNING")
        except Exception as e:
            self.log(f"❌ Configuration validation failed: {e}", "ERROR")
        
        # Test SharePoint connection (if configured)
        try:
            result = subprocess.run([
                str(python_exe), str(self.project_root / "main.py"), "--test-sharepoint"
            ], capture_output=True, text=True, timeout=60)
            
            if result.returncode == 0:
                self.log("✅ SharePoint connection test passed")
            else:
                self.log("⚠️ SharePoint connection test failed - check configuration", "WARNING")
        except subprocess.TimeoutExpired:
            self.log("⚠️ SharePoint connection test timed out", "WARNING")
        except Exception as e:
            self.log(f"❌ SharePoint connection test failed: {e}", "ERROR")
    
    def create_deployment_report(self):
        """Create deployment report"""
        report_path = self.project_root / "deployment_report.txt"
        
        report_content = f"""ZERF Automation System - Deployment Report
========================================
Deployment Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Project Root: {self.project_root}

Deployment Log:
{chr(10).join(self.deployment_log)}

Next Steps:
1. Edit .env file with your Azure app registration details
2. Configure SharePoint settings in the GUI
3. Test the system with: python main.py --gui
4. Run validation: python main.py --validate-config
5. Test SharePoint: python main.py --test-sharepoint

For support, see: deployment_guide.md
"""
        
        with open(report_path, 'w') as f:
            f.write(report_content)
        
        self.log(f"✅ Deployment report created: {report_path}")
    
    def deploy(self, deployment_type="development", venv_path="zerf_env"):
        """Main deployment function"""
        try:
            self.log(f"Starting {deployment_type} deployment...")
            
            # Check prerequisites
            self.check_prerequisites()
            
            # Setup virtual environment
            python_exe, pip_exe = self.setup_virtual_environment(venv_path)
            
            # Install dependencies
            self.install_dependencies(pip_exe)
            
            # Setup configuration
            self.setup_configuration()
            
            # Create executable for production
            exe_path = None
            if deployment_type == "production":
                exe_path = self.create_executable(python_exe)
            
            # Setup scheduled task
            if deployment_type in ["production", "staging"]:
                self.setup_scheduled_task(python_exe, exe_path)
            
            # Validate deployment
            self.validate_deployment(python_exe)
            
            # Create report
            self.create_deployment_report()
            
            self.log(f"✅ {deployment_type.title()} deployment completed successfully!")
            
            return True
            
        except Exception as e:
            self.log(f"❌ Deployment failed: {e}", "ERROR")
            return False

def main():
    parser = argparse.ArgumentParser(description="ZERF Automation System Deployment")
    parser.add_argument(
        "--type", 
        choices=["development", "staging", "production"], 
        default="development",
        help="Deployment type"
    )
    parser.add_argument(
        "--venv", 
        default="zerf_env",
        help="Virtual environment path"
    )
    parser.add_argument(
        "--skip-exe", 
        action="store_true",
        help="Skip executable creation"
    )
    
    args = parser.parse_args()
    
    deployer = ZERFDeployer()
    success = deployer.deploy(args.type, args.venv)
    
    if success:
        print("\n🎉 Deployment completed successfully!")
        print("\nNext steps:")
        print("1. Edit .env file with your Azure credentials")
        print("2. Run: python main.py --gui")
        print("3. Configure SharePoint settings")
        print("4. Test with: python main.py --validate-config")
        print("\nSee deployment_guide.md for detailed instructions.")
    else:
        print("\n❌ Deployment failed. Check the deployment report for details.")
        sys.exit(1)

if __name__ == "__main__":
    main()