'Define Variables and Objects.
Set WshShell = CreateObject("Wscript.Shell")

WshShell.Run ("cmd /c reg add ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\GPExtensions\{827D319E-6EAC-11D2-A4EA-00C04F79F83A}"" /v MaxNoGPOListChangesInterval /t REG_DWORD /d 960 /f")

WshShell.Run ("powershell.exe -executionpolicy bypass -command ""stop-service gpsvc -force""")

WshShell.Run ("powershell.exe -executionpolicy bypass -command ""start-service gpsvc -force""")

'Note: Gpupdate command has to be run twice as the ECHO command can't answer more than one question. 
'Refresh the USER policies and also answer no to logoff if asked.
'Result = WshShell.Run("cmd /c echo n | gpupdate /target:user /force",1,true)

'Refresh the Computer policies and answer no to reboot. 
Result = WshShell.Run ("cmd /c echo n | gpupdate /target:computer /force",1,true)

'Hand back the errorlevel
Wscript.Quit(Result)