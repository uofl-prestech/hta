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
    $svc = Get-Service MBAMAgent -ErrorAction SilentlyContinue
    $svc.WaitForStatus('Stopped', '00:00:15')
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
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" ClientWakeupFrequency 1 -ErrorAction SilentlyContinue

    #Change Registry key for StatusReportingFrequency from 120 to 1
    LogWrite " || Change Registry key for StatusReportingFrequency from 120 minutes to 1 minute"
    Write-Output "|| Change Registry key for StatusReportingFrequency from 120 minutes to 1 minute" | timestamp
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\StatusReportingFrequency to 1"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" StatusReportingFrequency 1 -ErrorAction SilentlyContinue
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
    Start-Service MBAMAgent -ErrorAction SilentlyContinue #Stop
    $svc = Get-Service MBAMAgent -ErrorAction SilentlyContinue
    $svc.WaitForStatus('Running', '00:00:30')
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
    Write-Output "|| $ErrorMessage" | timestamp
    Write-Output "|| Exiting script" | timestamp
}


#Wait for MBAM client to check in with server
$timer = [System.Diagnostics.Stopwatch]::StartNew()
$timer.start()
While($timer.Elapsed.Seconds -le 30){
    #Get MBAM log
    $MBAMLog = Get-WinEvent -LogName "Microsoft-Windows-MBAM/Operational" -ErrorAction SilentlyContinue
    Write-Output "$($timer.Elapsed.Seconds) Seconds elapsed - Event Log contains $($MBAMLog.Count) items"
    
    #If the MBAM log shows any records for Event ID 3, a response has been received from the server
    $MBAMResponse = $MBAMLog | Where-Object {$_.ID -eq 3}
    if($MBAMResponse.count -ge 1){
        Write-Output "A response has been received from the MBAM Server!"
        break
    }

    #Wait before checking the log again
    Start-Sleep -s 5
}


Read-Host -Prompt "Press any key to continue"


#Reset MBAM registry keys back to default values
try {
    #Change Registry key for ClientWakeupFrequency from 1 to 90
    LogWrite " || Change Registry key for ClientWakeupFrequency from 1 minute to 90 minutes"
    Write-Output "|| Change Registry key for ClientWakeupFrequency from 1 minute to 90 minutes" | timestamp
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\ClientWakeupFrequency to 90"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" ClientWakeupFrequency 90 -ErrorAction SilentlyContinue

    #Change Registry key for StatusReportingFrequency from 1 to 120
    LogWrite " || Change Registry key for StatusReportingFrequency from 1 minute to 120 minutes"
    Write-Output "|| Change Registry key for StatusReportingFrequency from 1 minute to 120 minutes" | timestamp
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\StatusReportingFrequency to 120"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" StatusReportingFrequency 120 -ErrorAction SilentlyContinue
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