Dim SapGuiAuto, application, connection, session, WshShell
Dim downloadPath, fileName, fullPath

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children(0)
Set session = connection.Children(0)
Set WshShell = CreateObject("WScript.Shell")

downloadPath = "C:\Users\jhash\Documents\DEALr\downloads"
fileName = "zerf_09-16-2025.xlsx"
fullPath = downloadPath & "\" & fileName

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
session.findById("wnd[0]/usr/ctxtSP$00018-LOW").text = "08/03/2025"
session.findById("wnd[0]/usr/ctxtSP$00018-HIGH").text = "09/15/2025"
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
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep 2000
    
    If WshShell.AppActivate("Confirm Save As") Then
        WScript.Sleep 500
        WshShell.SendKeys "{ENTER}"
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
