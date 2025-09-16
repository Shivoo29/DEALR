"""
SAP integration for ZERF Data Automation System
"""

import subprocess
import time
from pathlib import Path
from typing import Optional
from tenacity import retry, stop_after_attempt, wait_exponential

from ..utils.logger import get_logger
from ..utils.exceptions import SAPConnectionError, VBSScriptError, TimeoutError
from ..scripts.vbs_generator import VBSGenerator

logger = get_logger(__name__)

class SAPIntegration:
    """Handles all SAP-related operations"""
    
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.vbs_generator = VBSGenerator(config_manager)
        self.max_retries = config_manager.get_max_retries()
        self.timeout_minutes = config_manager.get_timeout_minutes()
    
    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=4, max=10),
        retry_error_callback=lambda retry_state: logger.warning(f"VBS script retry {retry_state.attempt_number}/3")
    )
    def extract_data(self, start_date: str, end_date: str) -> bool:
        """Extract data from SAP using VBS script"""
        try:
            logger.info("Starting SAP data extraction...")
            logger.info(f"Date range: {start_date} to {end_date}")
            
            # Generate VBS script
            vbs_script_path = self.vbs_generator.generate_script(start_date, end_date)
            
            # Validate SAP connection before running script
            if not self._check_sap_availability():
                raise SAPConnectionError("SAP GUI not available or not logged in")
            
            # Execute VBS script
            success = self._execute_vbs_script(vbs_script_path)
            
            if success:
                logger.info("✅ SAP data extraction completed successfully")
                return True
            else:
                raise VBSScriptError("VBS script execution failed")
                
        except Exception as e:
            logger.error(f"SAP data extraction failed: {e}")
            raise
    
    def _check_sap_availability(self) -> bool:
        """Check if SAP GUI is available and accessible"""
        try:
            # Create a simple VBS script to test SAP connection
            test_script = '''
            On Error Resume Next
            Dim SapGuiAuto, application, connection
            Set SapGuiAuto = GetObject("SAPGUI")
            If Err.Number <> 0 Then
                WScript.Echo "ERROR: SAP GUI not running"
                WScript.Quit 1
            End If
            
            Set application = SapGuiAuto.GetScriptingEngine
            If Err.Number <> 0 Then
                WScript.Echo "ERROR: SAP GUI scripting not enabled"
                WScript.Quit 1
            End If
            
            If application.Children.Count = 0 Then
                WScript.Echo "ERROR: No SAP connections available"
                WScript.Quit 1
            End If
            
            Set connection = application.Children(0)
            If connection.Children.Count = 0 Then
                WScript.Echo "ERROR: No active SAP sessions"
                WScript.Quit 1
            End If
            
            WScript.Echo "SUCCESS: SAP GUI is available"
            WScript.Quit 0
            '''
            
            test_script_path = Path("temp_sap_test.vbs")
            with open(test_script_path, 'w') as f:
                f.write(test_script)
            
            try:
                result = subprocess.run(
                    ['cscript', str(test_script_path), '//NoLogo'],
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                
                if result.returncode == 0:
                    logger.debug("SAP GUI availability check passed")
                    return True
                else:
                    logger.warning(f"SAP GUI check failed: {result.stdout.strip()}")
                    return False
                    
            finally:
                # Clean up test script
                if test_script_path.exists():
                    test_script_path.unlink()
                    
        except subprocess.TimeoutExpired:
            logger.error("SAP availability check timed out")
            return False
        except Exception as e:
            logger.error(f"SAP availability check error: {e}")
            return False
    
    def _execute_vbs_script(self, script_path: Path) -> bool:
        """Execute the VBS script for SAP data extraction"""
        try:
            logger.info(f"Executing VBS script: {script_path}")
            
            # Calculate timeout in seconds
            timeout_seconds = self.timeout_minutes * 60
            
            start_time = time.time()
            
            result = subprocess.run(
                ['cscript', str(script_path), '//NoLogo'],
                capture_output=True,
                text=True,
                timeout=timeout_seconds
            )
            
            execution_time = time.time() - start_time
            logger.info(f"VBS script execution time: {execution_time:.1f}s")
            
            if result.returncode == 0:
                logger.info("VBS script executed successfully")
                if result.stdout.strip():
                    logger.debug(f"VBS output: {result.stdout.strip()}")
                return True
            else:
                logger.error(f"VBS script failed with return code {result.returncode}")
                if result.stderr:
                    logger.error(f"VBS error: {result.stderr.strip()}")
                if result.stdout:
                    logger.error(f"VBS output: {result.stdout.strip()}")
                return False
                
        except subprocess.TimeoutExpired:
            logger.error(f"VBS script timed out after {self.timeout_minutes} minutes")
            raise TimeoutError(f"VBS script execution timed out after {self.timeout_minutes} minutes")
        except FileNotFoundError:
            logger.error("cscript.exe not found - Windows Scripting Host may not be installed")
            raise VBSScriptError("Windows Scripting Host not available")
        except Exception as e:
            logger.error(f"VBS script execution error: {e}")
            raise VBSScriptError(f"VBS script execution failed: {e}")
    
    def test_sap_connection(self) -> bool:
        """Test SAP connection and return detailed status"""
        try:
            logger.info("Testing SAP connection...")
            
            if not self._check_sap_availability():
                logger.error("❌ SAP connection test failed")
                return False
            
            # Test ZERF transaction access
            if not self._test_zerf_transaction():
                logger.error("❌ ZERF transaction access test failed")
                return False
            
            logger.info("✅ SAP connection test passed")
            return True
            
        except Exception as e:
            logger.error(f"SAP connection test error: {e}")
            return False
    
    def _test_zerf_transaction(self) -> bool:
        """Test access to ZERF transaction"""
        try:
            test_script = '''
            On Error Resume Next
            Dim SapGuiAuto, application, connection, session
            Set SapGuiAuto = GetObject("SAPGUI")
            Set application = SapGuiAuto.GetScriptingEngine
            Set connection = application.Children(0)
            Set session = connection.Children(0)
            
            ' Try to navigate to ZERF transaction
            session.findById("wnd[0]/tbar[0]/okcd").text = "zerf"
            session.findById("wnd[0]").sendVKey 0
            
            ' Wait a moment for the transaction to load
            WScript.Sleep 2000
            
            ' Check if we're in the ZERF transaction
            Dim title
            title = session.findById("wnd[0]").text
            
            If InStr(title, "ZERF") > 0 Or InStr(title, "Engineering Request") > 0 Then
                WScript.Echo "SUCCESS: ZERF transaction accessible"
                
                ' Navigate back to main menu
                session.findById("wnd[0]").sendVKey 15  ' F15 - Back
                WScript.Quit 0
            Else
                WScript.Echo "ERROR: ZERF transaction not accessible"
                WScript.Quit 1
            End If
            '''
            
            test_script_path = Path("temp_zerf_test.vbs")
            with open(test_script_path, 'w') as f:
                f.write(test_script)
            
            try:
                result = subprocess.run(
                    ['cscript', str(test_script_path), '//NoLogo'],
                    capture_output=True,
                    text=True,
                    timeout=60
                )
                
                return result.returncode == 0
                
            finally:
                if test_script_path.exists():
                    test_script_path.unlink()
                    
        except Exception as e:
            logger.error(f"ZERF transaction test error: {e}")
            return False
    
    def get_sap_system_info(self) -> dict:
        """Get SAP system information"""
        try:
            info_script = '''
            On Error Resume Next
            Dim SapGuiAuto, application, connection, session
            Set SapGuiAuto = GetObject("SAPGUI")
            Set application = SapGuiAuto.GetScriptingEngine
            Set connection = application.Children(0)
            Set session = connection.Children(0)
            
            WScript.Echo "System ID: " & connection.ConnectionString
            WScript.Echo "User: " & session.Info.User
            WScript.Echo "Client: " & session.Info.Client
            WScript.Echo "Language: " & session.Info.Language
            WScript.Echo "Session ID: " & session.Id
            '''
            
            info_script_path = Path("temp_sap_info.vbs")
            with open(info_script_path, 'w') as f:
                f.write(info_script)
            
            try:
                result = subprocess.run(
                    ['cscript', str(info_script_path), '//NoLogo'],
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                
                info = {}
                if result.returncode == 0 and result.stdout:
                    for line in result.stdout.strip().split('\n'):
                        if ':' in line:
                            key, value = line.split(':', 1)
                            info[key.strip()] = value.strip()
                
                return info
                
            finally:
                if info_script_path.exists():
                    info_script_path.unlink()
                    
        except Exception as e:
            logger.error(f"Failed to get SAP system info: {e}")
            return {}
    
    def cleanup_temp_files(self):
        """Clean up temporary VBS files"""
        try:
            temp_files = Path('.').glob('temp_*.vbs')
            for temp_file in temp_files:
                try:
                    temp_file.unlink()
                    logger.debug(f"Cleaned up temp file: {temp_file}")
                except Exception as e:
                    logger.warning(f"Failed to clean up {temp_file}: {e}")
        except Exception as e:
            logger.warning(f"Cleanup temp files error: {e}")
    
    def get_extraction_status(self) -> dict:
        """Get status of the last extraction operation"""
        return {
            'sap_available': self._check_sap_availability(),
            'zerf_accessible': self._test_zerf_transaction(),
            'system_info': self.get_sap_system_info(),
            'timeout_minutes': self.timeout_minutes,
            'max_retries': self.max_retries
        }