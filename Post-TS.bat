type NUL > C:\Windows\TEMP\file.log
echo. TESTING > C:\Windows\TEMP\file.log
reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Runonce" /v MBAMUpdate /d "C:\windows\TEMP\MBAM_ReportStatus.vbs" /f