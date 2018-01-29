#**********************************************************************************************************************
#						        Create Log File
#**********************************************************************************************************************
[string]$currentDir = Get-Location

$logFile = "$currentDir\MBAMLOG.log"
filter timestamp {"$(Get-Date -Format G) $_"}

Function LogWrite
{
    Param ([string]$logString)
    $stamp = (Get-Date).toString("MM/dd/yyyy HH:mm:ss")
    $line = "$stamp $logString"
    Add-content $logFile -value $line
}

LogWrite "***** Begin Script MBAM_ReportStatus.ps1 *****"

#Run GPupdate
LogWrite " || Running GPupdate..."
Write-Output "|| Running GPupdate..." | timestamp
#Start-Process gpupdate.exe -NoNewWindow -Wait
LogWrite " || GPupdate Complete"
Write-Output "|| GPupdate Complete" | timestamp

#Stop MBAM Service
LogWrite " || Stopping MBAM Service..."
Write-Output "|| Stopping MBAM Service..." | timestamp
try {
    Stop-Service MBAMAgent -force -ErrorAction SilentlyContinue #Stop
    LogWrite " || Successfully stopped MBAM Service"
    Write-Output "|| Successfully stopped MBAM Service" | timestamp
}
catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    LogWrite " || Error - MBAMAgent Service not found"
    LogWrite " || $ErrorMessage"
    LogWrite " || Exiting script"
    Write-Output "|| Error - MBAMAgent Service not found" | timestamp
    Write-Output "|| Exiting script" | timestamp
}


#Modify Registry keys to force the MBAM client to report in every minute
try{
    #Change Registry key for ClientWakeupFrequency from 90 to 1
    LogWrite " || Change Registry key for ClientWakeupFrequency from 90 minutes to 1 minute"
    Write-Output "|| Change Registry key for ClientWakeupFrequency from 90 minutes to 1 minute" | timestamp
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\ClientWakeupFrequency to 1"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" ClientWakeupFrequency 1

    #Change Registry key for StatusReportingFrequency from 120 to 1
    LogWrite " || Change Registry key for StatusReportingFrequency from 120 minutes to 1 minute"
    Write-Output "|| Change Registry key for StatusReportingFrequency from 120 minutes to 1 minute" | timestamp
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\StatusReportingFrequency to 1"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" StatusReportingFrequency 1
}
catch{
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    LogWrite " || Error changing registry key $FailedItem"
    LogWrite " || $ErrorMessage"
    LogWrite " || Exiting script"
    Write-Output "|| Error changing registry key $FailedItem" | timestamp
    Write-Output "|| $ErrorMessage" | timestamp
    Write-Output "|| Exiting script" | timestamp
}


#Start MBAM Service
LogWrite " || Starting MBAM Service..."
Write-Output "|| Starting MBAM Service..." | timestamp
try {
    Start-Service MBAMAgent -force -ErrorAction SilentlyContinue #Stop
    LogWrite " || Successfully started MBAM Service"
    Write-Output "|| Successfully started MBAM Service" | timestamp
}
catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    LogWrite " || Error - MBAMAgent Service couldn't start"
    LogWrite " || $ErrorMessage"
    LogWrite " || Exiting script"
    Write-Output "|| Error - MBAMAgent Service couldn't start" | timestamp
    Write-Output "|| Exiting script" | timestamp
}




#Reset MBAM registry keys back to default values
try {
    #Change Registry key for ClientWakeupFrequency from 1 to 90
    LogWrite " || Change Registry key for ClientWakeupFrequency from 1 minute to 90 minutes"
    Write-Output "|| Change Registry key for ClientWakeupFrequency from 1 minute to 90 minutes" | timestamp
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\ClientWakeupFrequency to 90"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" ClientWakeupFrequency 90

    #Change Registry key for StatusReportingFrequency from 1 to 120
    LogWrite " || Change Registry key for StatusReportingFrequency from 1 minute to 120 minutes"
    Write-Output "|| Change Registry key for StatusReportingFrequency from 1 minute to 120 minutes" | timestamp
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\StatusReportingFrequency to 120"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" StatusReportingFrequency 120
}
catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    LogWrite " || Error changing registry key $FailedItem"
    LogWrite " || $ErrorMessage"
    LogWrite " || Exiting script"
    Write-Output "|| Error changing registry key $FailedItem" | timestamp
    Write-Output "|| $ErrorMessage" | timestamp
    Write-Output "|| Exiting script" | timestamp
}

<#


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
#>