"""
VBS Script Generator for SAP ZERF automation
"""

from pathlib import Path
from datetime import datetime
from string import Template

from ..utils.logger import get_logger
from ..utils.exceptions import VBSScriptError

logger = get_logger(__name__)

class VBSGenerator:
    """Generates VBS scripts for SAP automation"""
    
    # VBS Script template with improved error handling
    VBS_TEMPLATE = '''
' ZERF Data Extraction VBS Script
' Generated on: $generation_time
' Date Range: $start_date to $end_date
' Target File: $target_file

Dim SapGuiAuto, application, connection, session, WshShell
Dim downloadPath, fileName, fullPath
Dim startTime, maxWaitTime, attempts

' Initialize timing
startTime = Timer
maxWaitTime = $max_wait_seconds ' Maximum wait time in seconds

Set WshShell = CreateObject("WScript.Shell")

' Enhanced error handling
On Error Resume Next

' Connect to SAP GUI
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot connect to SAP GUI. Please ensure SAP GUI is running and logged in."
    WScript.Quit 1
End If

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Then
    WScript.Echo "ERROR: SAP GUI scripting not enabled. Please enable scripting in SAP GUI."
    WScript.Quit 1
End If

If application.Children.Count = 0 Then
    WScript.Echo "ERROR: No SAP connections available."
    WScript.Quit 1
End If

Set connection = application.Children(0)
If connection.Children.Count = 0 Then
    WScript.Echo "ERROR: No active SAP sessions."
    WScript.Quit 1
End If

Set session = connection.Children(0)

' File path configuration
downloadPath = "$download_path"
fileName = "$filename"
fullPath = downloadPath & "\\" & fileName

WScript.Echo "Starting ZERF data extraction..."
WScript.Echo "Target file: " & fullPath
WScript.Echo "Date range: $start_date to $end_date"

' Clear any previous errors
Err.Clear

' SAP Navigation and Data Extraction
Try
    ' Maximize window and navigate to ZERF
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "zerf"
    session.findById("wnd[0]").sendVKey 0
    
    ' Wait for transaction to load
    WScript.Sleep 2000
    
    ' Check if we're in the right transaction
    If InStr(session.findById("wnd[0]").text, "ZERF") = 0 And InStr(session.findById("wnd[0]").text, "Engineering") = 0 Then
        WScript.Echo "ERROR: Failed to navigate to ZERF transaction"
        WScript.Quit 1
    End If
    
    ' Clear Ship-To-Plant field and open selection
    session.findById("wnd[0]/usr/ctxtSP$$00011-HIGH").text = ""
    session.findById("wnd[0]/usr/ctxtSP$$00011-HIGH").setFocus
    session.findById("wnd[0]/usr/ctxtSP$$00011-HIGH").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_SP$$00011_%_APP_%-VALU_PUSH").press
    
    ' Wait for selection dialog
    WScript.Sleep 1000
    
    ' Enter Ship-To-Plant values
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1010"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1020"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "1090"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "6100"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "6200"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6300"
    
    ' Confirm selection
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' Wait for dialog to close
    WScript.Sleep 1000
    
    ' Set date range
    session.findById("wnd[0]/usr/ctxtSP$$00018-LOW").text = "$start_date"
    session.findById("wnd[0]/usr/ctxtSP$$00018-HIGH").text = "$end_date"
    
    ' Execute the report
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Wait for report to load
    WScript.Sleep 5000
    
    ' Start export process
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
    
    ' Wait for export dialog
    WScript.Sleep 3000
    
    ' Handle SAP export dialog if it appears
    If Not (session.findById("wnd[1]") Is Nothing) Then
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
        WScript.Sleep 1000
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        WScript.Sleep 3000
    End If
    
Catch
    WScript.Echo "ERROR: SAP navigation failed - " & Err.Description
    WScript.Quit 1
End Try

' Handle Windows Save As dialog with improved reliability
WScript.Echo "Handling file save dialog..."

Dim dialogAttempts, maxDialogAttempts
maxDialogAttempts = 30
dialogAttempts = 0

Do While dialogAttempts < maxDialogAttempts
    If WshShell.AppActivate("Save As") Or WshShell.AppActivate("Export") Or WshShell.AppActivate("Save") Or WshShell.AppActivate("Speichern") Then
        WScript.Sleep 1000
        Exit Do
    End If
    WScript.Sleep 1000
    dialogAttempts = dialogAttempts + 1
Loop

If dialogAttempts < maxDialogAttempts Then
    WScript.Echo "Save dialog found, entering file path..."
    
    ' Clear any existing text and enter our path
    WshShell.SendKeys "^a"
    WScript.Sleep 500
    WshShell.SendKeys fullPath
    WScript.Sleep 1000
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep 2000
    
    ' Handle potential overwrite confirmation
    If WshShell.AppActivate("Confirm Save As") Or WshShell.AppActivate("Speichern bestätigen") Then
        WScript.Sleep 500
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 1000
    End If
    
    WScript.Echo "File save initiated"
Else
    WScript.Echo "WARNING: Save dialog not found within timeout period"
End If

' Wait for file to be created and verify it exists
WScript.Echo "Waiting for file to be saved..."
WScript.Sleep 5000

' Enhanced file verification with multiple attempts
Dim fileVerified, verifyAttempts, maxVerifyAttempts
fileVerified = False
maxVerifyAttempts = 10
verifyAttempts = 0

Do While Not fileVerified And verifyAttempts < maxVerifyAttempts
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(fullPath) Then
        ' File exists, now check if it's accessible (not locked)
        On Error Resume Next
        Dim testFile
        Set testFile = fso.OpenTextFile(fullPath, 1)
        If Err.Number = 0 Then
            testFile.Close
            fileVerified = True
            WScript.Echo "File successfully saved and verified: " & fullPath
        Else
            WScript.Echo "File exists but is still locked, waiting..."
            WScript.Sleep 2000
        End If
        On Error GoTo 0
    Else
        WScript.Echo "File not yet created, waiting..."
        WScript.Sleep 2000
    End If
    
    verifyAttempts = verifyAttempts + 1
    Set fso = Nothing
Loop

If Not fileVerified Then
    WScript.Echo "WARNING: Could not verify file creation within timeout period"
    WScript.Echo "Expected file: " & fullPath
End If

' Cleanup
Set WshShell = Nothing
Set session = Nothing
Set connection = Nothing
Set application = Nothing
Set SapGuiAuto = Nothing

WScript.Echo "VBS script execution completed"

' Subroutines
Sub Try
    ' Error handling marker
End Sub

Sub Catch
    If Err.Number <> 0 Then
        WScript.Echo "ERROR in SAP operation: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
End Sub
'''
    
    def __init__(self, config_manager):
        self.config_manager = config_manager
    
    def generate_script(self, start_date: str, end_date: str) -> Path:
        """Generate VBS script for SAP data extraction"""
        try:
            logger.info(f"Generating VBS script for date range: {start_date} to {end_date}")
            
            # Get configuration
            download_path = str(self.config_manager.get_download_folder().absolute()).replace("/", "\\")
            vbs_script_path = self.config_manager.get_vbs_script_path()
            
            # Ensure script directory exists
            vbs_script_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Generate filename
            today = datetime.now().strftime("%m-%d-%Y")
            filename = f"zerf_{today}.xlsx"
            
            # Calculate maximum wait time based on timeout setting
            max_wait_seconds = self.config_manager.get_timeout_minutes() * 60
            
            # Prepare template variables
            template_vars = {
                'generation_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'start_date': start_date,
                'end_date': end_date,
                'download_path': download_path,
                'filename': filename,
                'target_file': f"{download_path}\\{filename}",
                'max_wait_seconds': max_wait_seconds
            }
            
            # Generate script content
            template = Template(self.VBS_TEMPLATE)
            script_content = template.substitute(template_vars)
            
            # Write script to file
            with open(vbs_script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)
            
            logger.info(f"VBS script generated: {vbs_script_path}")
            logger.info(f"Target file: {download_path}\\{filename}")
            
            return vbs_script_path
            
        except Exception as e:
            logger.error(f"Failed to generate VBS script: {e}")
            raise VBSScriptError(f"VBS script generation failed: {e}")
    
    def validate_script(self, script_path: Path) -> bool:
        """Validate generated VBS script"""
        try:
            if not script_path.exists():
                logger.error(f"VBS script not found: {script_path}")
                return False
            
            # Read and basic validation
            with open(script_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Check for required elements
            required_elements = [
                'SAPGUI',
                'GetScriptingEngine',
                'zerf',
                'findById',
                'pressToolbarContextButton',
                'selectContextMenuItem'
            ]
            
            for element in required_elements:
                if element not in content:
                    logger.error(f"VBS script missing required element: {element}")
                    return False
            
            logger.debug("VBS script validation passed")
            return True
            
        except Exception as e:
            logger.error(f"VBS script validation error: {e}")
            return False
    
    def get_script_info(self, script_path: Path) -> dict:
        """Get information about a VBS script"""
        try:
            if not script_path.exists():
                return {}
            
            info = {
                'path': str(script_path),
                'size': script_path.stat().st_size,
                'modified': datetime.fromtimestamp(script_path.stat().st_mtime),
                'valid': self.validate_script(script_path)
            }
            
            # Extract date range from script if possible
            try:
                with open(script_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Look for date range in comments
                for line in content.split('\n'):
                    if 'Date Range:' in line:
                        info['date_range'] = line.split('Date Range:')[1].strip()
                        break
                        
            except Exception:
                pass
            
            return info
            
        except Exception as e:
            logger.error(f"Failed to get script info: {e}")
            return {}
    
    def cleanup_old_scripts(self, days_to_keep: int = 7):
        """Clean up old VBS scripts"""
        try:
            script_dir = self.config_manager.get_vbs_script_path().parent
            if not script_dir.exists():
                return
            
            cutoff_time = datetime.now().timestamp() - (days_to_keep * 24 * 60 * 60)
            cleaned_count = 0
            
            for script_file in script_dir.glob("*.vbs"):
                if script_file.stat().st_mtime < cutoff_time:
                    try:
                        script_file.unlink()
                        cleaned_count += 1
                        logger.debug(f"Cleaned up old script: {script_file}")
                    except Exception as e:
                        logger.warning(f"Failed to delete {script_file}: {e}")
            
            if cleaned_count > 0:
                logger.info(f"Cleaned up {cleaned_count} old VBS scripts")
                
        except Exception as e:
            logger.error(f"Script cleanup error: {e}")
    
    def create_test_script(self) -> Path:
        """Create a test VBS script for SAP connection testing"""
        test_script = '''
        ' SAP Connection Test Script
        On Error Resume Next
        
        Dim SapGuiAuto, application, connection, session
        Set SapGuiAuto = GetObject("SAPGUI")
        
        If Err.Number <> 0 Then
            WScript.Echo "FAIL: SAP GUI not running"
            WScript.Quit 1
        End If
        
        Set application = SapGuiAuto.GetScriptingEngine
        If Err.Number <> 0 Then
            WScript.Echo "FAIL: SAP GUI scripting not enabled"
            WScript.Quit 1
        End If
        
        If application.Children.Count = 0 Then
            WScript.Echo "FAIL: No SAP connections"
            WScript.Quit 1
        End If
        
        Set connection = application.Children(0)
        If connection.Children.Count = 0 Then
            WScript.Echo "FAIL: No active sessions"
            WScript.Quit 1
        End If
        
        Set session = connection.Children(0)
        WScript.Echo "SUCCESS: SAP connection OK"
        WScript.Echo "User: " & session.Info.User
        WScript.Echo "Client: " & session.Info.Client
        WScript.Echo "System: " & connection.ConnectionString
        '''
        
        test_script_path = Path("temp_sap_test.vbs")
        with open(test_script_path, 'w') as f:
            f.write(test_script)
        
        return test_script_path