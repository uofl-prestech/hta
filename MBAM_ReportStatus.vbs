'Stop MBAM Service
strServiceName = "MBAMAgent"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
    objService.StopService()
Next




'Setup for Registry modifications
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement"
dwValue = 1



'Change Registry key for ClientWakeupFrequency from 90 to 1
strValueName = "ClientWakeupFrequency"

objRegistry.SetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, dwValue
If Err = 0 Then
   objRegistry.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
   WScript.Echo strKeyPath & "\" & strValueName & " contains: " & dwValue
Else 
   WScript.Echo "Error in creating key and DWORD value = " & Err.Number
End If


'Change Registry key for StatusReportingFrequency from 120 to 1
strValueName = "StatusReportingFrequency"

objRegistry.SetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, dwValue
If Err = 0 Then
   objRegistry.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
   WScript.Echo strKeyPath & "\" & strValueName & " contains: " & dwValue
Else 
   WScript.Echo "Error in creating key and DWORD value = " & Err.Number
End If




'Start MBAM Service
strServiceName = "MBAMAgent"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
    objService.StartService()
Next

wscript.echo "Sleeping for 2 minutes"
wscript.sleep 120000

wscript.quit(0)