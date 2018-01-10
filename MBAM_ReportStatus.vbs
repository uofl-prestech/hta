On Error Resume Next
'**********************************************************************************************************************
'						        Create Log File
'**********************************************************************************************************************
CONST ForAppending = 8
Dim logShell, strLogDir, objFSO, htaLog
Set logShell = CreateObject("WScript.Shell")
strLogDir = logShell.currentDirectory
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set mbamLog = objFSO.OpenTextFile(strLogDir & "\MBAMLOG.log", ForAppending, True)
mbamLog.WriteLine(vbCrLf & "====================== Begin Logging at " & Now & " ======================")





'Stop MBAM Service
mbamLog.WriteLine(Now & " || Stopping MBAM Service")
strServiceName = "MBAMAgent"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
    objService.StopService()
Next
If Err = 0 Then
    mbamLog.WriteLine(Now & " || Successfully stopped MBAM Service")
Else
    mbamLog.WriteLine(Now & " || Error " & Err.Number & ": " & Err.Description)
End If
Err.Clear





'Setup for Registry modifications
mbamLog.WriteLine(Now & " || Setup for Registry modifications")
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement"
dwValue = 1





'Change Registry key for ClientWakeupFrequency from 90 to 1
mbamLog.WriteLine(Now & " || Change Registry key for ClientWakeupFrequency from 90 minutes to 1 minute")
strValueName = "ClientWakeupFrequency"
objRegistry.SetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, dwValue
If Err = 0 Then
   objRegistry.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
   mbamLog.WriteLine(Now & " || " & strKeyPath & "\" & strValueName & " contains: " & dwValue)
Else 
   mbamLog.WriteLine(Now & " || Error in creating key and DWORD value")
   mbamLog.WriteLine(Now & " || Error " & Err.Number & ": " & Err.Description)
End If
Err.Clear




'Change Registry key for StatusReportingFrequency from 120 to 1
mbamLog.WriteLine(Now & " || Change Registry key for StatusReportingFrequency from 120 minutes to 1 minute")
strValueName = "StatusReportingFrequency"
objRegistry.SetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, dwValue
If Err = 0 Then
   objRegistry.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
   mbamLog.WriteLine(Now & " || " & strKeyPath & "\" & strValueName & " contains: " & dwValue)
Else 
   mbamLog.WriteLine(Now & " || Error in creating key and DWORD value")
   mbamLog.WriteLine(Now & " || Error " & Err.Number & ": " & Err.Description)
End If
Err.Clear





'Start MBAM Service
mbamLog.WriteLine(Now & " || Starting MBAM service")
strServiceName = "MBAMAgent"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
    objService.StartService()
Next
If Err = 0 Then
    mbamLog.WriteLine(Now & " || Successfully started MBAM Service")
Else
    mbamLog.WriteLine(Now & " || Error " & Err.Number & ": " & Err.Description)
End If
Err.Clear





mbamLog.WriteLine(Now & " || Sleeping for 2 minutes")
wscript.sleep 120000
mbamLog.WriteLine(Now & " || Continuing script after 2 minute pause")


Set WshShell = CreateObject("Wscript.Shell")
mbamLog.WriteLine(Now & " || Running gpupdate /force")
Result = WshShell.Run ("cmd /c echo n | gpupdate /target:computer /force", 1, true)
mbamLog.WriteLine(Now & " || Rebooting computer")
WshShell.Run ("cmd /c shutdown /a", 1, true)
WshShell.Run ("cmd /c shutdown /r /t 0", 1, true)
mbamLog.Close
wscript.quit(0)