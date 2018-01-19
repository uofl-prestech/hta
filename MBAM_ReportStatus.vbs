If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute "cscript.exe", WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

On Error Resume Next
'**********************************************************************************************************************
'						        Create Log File
'**********************************************************************************************************************
CONST ForAppending = 8
Dim objFSO, mbamLog, Result, Return
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set mbamLog = objFSO.OpenTextFile("C:\Windows\TEMP\MBAMLOG.log", ForAppending, True)
mbamLog.WriteLine(vbCrLf & "====================== Begin Logging at " & Now & " ======================")
wscript.echo Now & " || Running GPupdate..."
Set WshShell = CreateObject("Wscript.Shell")
Result = WshShell.Run ("cmd /c echo n | gpupdate /target:computer /force", 1, true)
wscript.echo Now & " || GPupdate finished"

'Stop MBAM Service
mbamLog.WriteLine(Now & " || Stopping MBAM Service")
wscript.echo Now & " || Stopping MBAM Service..."
strServiceName = "MBAMAgent"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
    objService.StopService()
Next
If Err = 0 Then
    mbamLog.WriteLine(Now & " || Successfully stopped MBAM Service")
    wscript.echo Now & " || Successfully stopped MBAM Service"
Else
    mbamLog.WriteLine(Now & " || Error " & Err.Number & ": " & Err.Description)
End If
Err.Clear




'Setup for Registry modifications
mbamLog.WriteLine(Now & " || Setup for Registry modifications")
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
strValueName = "MBAM"




'Remove script from Run key so it doesn't run every time a user logs in
mbamLog.WriteLine(Now & " || Remove script from Run key")
Return = objRegistry.DeleteValue(HKEY_LOCAL_MACHINE, strKeyPath, strValueName)
If (Return = 0) And (Err.Number = 0) Then    
    wscript.echo Now & " || HKEY_LOCAL_MACHINE\" & strKeyPath & "\" & strValueName & " successfully deleted"
    mbamLog.WriteLine(Now & " || HKEY_LOCAL_MACHINE\" & strKeyPath & "\" & strValueName &  " successfully deleted")
Else
    wscript.echo Now & " || Delete key HKEY_LOCAL_MACHINE\" & strKeyPath & "\" & strValueName &  " failed. Error = " & Err.Number
    mbamLog.WriteLine(Now & " || Delete key HKEY_LOCAL_MACHINE\" & strKeyPath & "\" & strValueName &  " failed. Error = " & Err.Number)
End If
Err.Clear




'Change Registry key for ClientWakeupFrequency from 90 to 1
wscript.echo Now & " || Changing registry keys to force MBAM to check in every minute"
mbamLog.WriteLine(Now & " || Change Registry key for ClientWakeupFrequency from 90 minutes to 1 minute")
strKeyPath = "SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement"
dwValue = 1
strValueName = "ClientWakeupFrequency"
objRegistry.SetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, dwValue
If Err = 0 Then
   objRegistry.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
   mbamLog.WriteLine(Now & " || " & strKeyPath & "\" & strValueName & " contains: " & dwValue)
   wscript.echo Now & " || " & strKeyPath & "\" & strValueName & " contains: " & dwValue
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
   wscript.echo Now & " || " & strKeyPath & "\" & strValueName & " contains: " & dwValue
Else 
   mbamLog.WriteLine(Now & " || Error in creating key and DWORD value")
   mbamLog.WriteLine(Now & " || Error " & Err.Number & ": " & Err.Description)
End If
Err.Clear




'Start MBAM Service
mbamLog.WriteLine(Now & " || Starting MBAM service")
wscript.echo Now & " || Starting MBAM service..."
strServiceName = "MBAMAgent"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
    objService.StartService()
Next
If Err = 0 Then
    mbamLog.WriteLine(Now & " || Successfully started MBAM Service")
    wscript.echo Now & " || Successfully started MBAM Service"
Else
    mbamLog.WriteLine(Now & " || Error " & Err.Number & ": " & Err.Description)
End If
Err.Clear




mbamLog.WriteLine(Now & " || Sleeping for 3 minutes")
wscript.echo Now & " || Sleeping for 3 minutes to allow the MBAM client to communicate with the MBAM server..."
wscript.sleep 180000
mbamLog.WriteLine(Now & " || Continuing script after 3 minute pause")
wscript.echo Now & " || Continuing script after 3 minute pause"



mbamLog.WriteLine(Now & " || Running gpupdate /force")
wscript.echo Now & " || Running GPupdate..."
Result = WshShell.Run ("cmd /c echo n | gpupdate /target:computer /force", 1, true)
wscript.echo Now & " || GPupdate finished"
wscript.echo Now & " || Script complete!"
'mbamLog.WriteLine(Now & " || Rebooting computer")
'WshShell.Run "cmd /c shutdown /a", 1, true
'WshShell.Run "cmd /c shutdown /r /t 0", 1, true
mbamLog.WriteLine(Now & " || Finished!")
mbamLog.Close
wscript.echo "Press Enter key to continue"
wscript.StdIn.ReadLine
wscript.quit(0)