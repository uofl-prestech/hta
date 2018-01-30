#***********************************************************************************************************************#
#						        Create Log File                                                                         #
#***********************************************************************************************************************#
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


#*********************************************************************************************************************#
#						        Run GPupdate                                                                          #
#*********************************************************************************************************************#
LogWrite " || Running GPupdate..."
Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Running GPupdate..."
#Start-Process gpupdate.exe -NoNewWindow -Wait
LogWrite " || GPupdate Complete"
Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| GPupdate Complete`n" -ForegroundColor Green


#*********************************************************************************************************************#
#						        Stop MBAM Service                                                                     #
#*********************************************************************************************************************#
LogWrite " || Stopping MBAM Service..."
Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Stopping MBAM Service..."
try {
    Stop-Service MBAMAgent -force -ErrorAction SilentlyContinue #Stop
    $svc = Get-Service MBAMAgent -ErrorAction SilentlyContinue
    $svc.WaitForStatus('Stopped', '00:00:15')
    LogWrite " || Successfully stopped MBAM Service"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Successfully stopped MBAM Service`n" -ForegroundColor Green
}
catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    LogWrite " || Error - MBAMAgent Service not found"
    LogWrite " || $ErrorMessage"
    LogWrite " || Exiting script"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Error - MBAMAgent Service not found" -ForegroundColor Red
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Exiting script"
}


#*********************************************************************************************************************#
#						        Modify Registry keys to force the MBAM client to report in every minute               #
#*********************************************************************************************************************#
try{
    #Change Registry key for ClientWakeupFrequency from 90 to 1
    LogWrite " || Change Registry key for ClientWakeupFrequency from 90 minutes to 1 minute"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Change Registry key for ClientWakeupFrequency from 90 minutes to 1 minute"
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\ClientWakeupFrequency to 1"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" ClientWakeupFrequency 1 -ErrorAction SilentlyContinue

    #Change Registry key for StatusReportingFrequency from 120 to 1
    LogWrite " || Change Registry key for StatusReportingFrequency from 120 minutes to 1 minute"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Change Registry key for StatusReportingFrequency from 120 minutes to 1 minute`n"
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\StatusReportingFrequency to 1"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" StatusReportingFrequency 1 -ErrorAction SilentlyContinue
}
catch{
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    LogWrite " || Error changing registry key $FailedItem"
    LogWrite " || $ErrorMessage"
    LogWrite " || Exiting script"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Error changing registry key $FailedItem" -ForegroundColor Red
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| $ErrorMessage"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Exiting script"
}


#*********************************************************************************************************************#
#						        Start MBAM Service                                                                    #
#*********************************************************************************************************************#
LogWrite " || Starting MBAM Service..."
Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Starting MBAM Service..."
try {
    Start-Service MBAMAgent -ErrorAction SilentlyContinue #Stop
    $svc = Get-Service MBAMAgent -ErrorAction SilentlyContinue
    $svc.WaitForStatus('Running', '00:00:30')
    LogWrite " || Successfully started MBAM Service"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Successfully started MBAM Service`n`n" -ForegroundColor Green
}
catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    LogWrite " || Error - MBAMAgent Service couldn't start"
    LogWrite " || $ErrorMessage"
    LogWrite " || Exiting script"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Error - MBAMAgent Service couldn't start" -ForegroundColor Red
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| $ErrorMessage"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Exiting script"
}


#*********************************************************************************************************************#
#						        Wait for MBAM client to check in with server                                          #
#*********************************************************************************************************************#
LogWrite " || Waiting for the MBAM client to communicate with the MBAM server"
Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "Waiting for the MBAM client to communicate with the MBAM server" -ForegroundColor Yellow
$timer = [System.Diagnostics.Stopwatch]::StartNew()
$timer.start()
While($timer.Elapsed.Seconds -le 180){
    #Get MBAM log
    $MBAMLog = Get-WinEvent -LogName "Microsoft-Windows-MBAM/Operational" -ErrorAction SilentlyContinue
    LogWrite " || $($timer.Elapsed.Seconds) Seconds elapsed - MBAM Event Log contains $($MBAMLog.Count) item(s)"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "$($timer.Elapsed.Seconds) Seconds elapsed - MBAM Event Log contains $($MBAMLog.Count) item(s)"
    
    #If the MBAM log shows any records for Event ID 3 and 29, a response has been received from the server
    $MBAMResponse = $MBAMLog | Where-Object {$_.ID -eq 3 -or $_.ID -eq 29}
    $entry3 = $MBAMResponse | Where-Object {$_.ID -eq 3} | Select-Object -Last 1
    $entry29 = $MBAMResponse | Where-Object {$_.ID -eq 29} | Select-Object -Last 1
    #If the log has any entries with ID 3 and hasn't reported it yet
    if($entry3 -and !$entry3Reported){
        LogWrite " || An Event with ID 3 has been logged"
        LogWrite " || $entry3.Message"
        Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "An Event with ID 3 has been logged" -ForegroundColor Green
        Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") $entry3.Message
        $entry3Reported = $TRUE
    }
    #If the log has any entries with ID 29 and hasn't reported it yet
    if($entry29 -and !$entry29Reported){
        LogWrite " || An Event with ID 29 has been logged"
        LogWrite " || $entry29.Message"
        Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "An Event with ID 29 has been logged" -ForegroundColor Green
        Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") $entry29.Message
        $entry29Reported = $TRUE
    }
    if($entry3Reported -and $entry29Reported){
        LogWrite " || \o/ MBAM communication successful \o/"
        Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "\o/ MBAM communication successful \o/`n" -ForegroundColor Green
        break
    }

    #Wait before checking the log again
    Start-Sleep -s 5
}


#*********************************************************************************************************************#
#						        Reset MBAM registry keys back to default values                                       #
#*********************************************************************************************************************#
try {
    #Change Registry key for ClientWakeupFrequency from 1 to 90
    LogWrite " || Change Registry key for ClientWakeupFrequency from 1 minute to 90 minutes"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Change Registry key for ClientWakeupFrequency from 1 minute to 90 minutes"
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\ClientWakeupFrequency to 90"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" ClientWakeupFrequency 90 -ErrorAction SilentlyContinue

    #Change Registry key for StatusReportingFrequency from 1 to 120
    LogWrite " || Change Registry key for StatusReportingFrequency from 1 minute to 120 minutes"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Change Registry key for StatusReportingFrequency from 1 minute to 120 minutes"
    LogWrite " || Setting key HKLM\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement\StatusReportingFrequency to 120"
    Set-ItemProperty -Path "hklm:\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement" StatusReportingFrequency 120 -ErrorAction SilentlyContinue
}
catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    LogWrite " || Error changing registry key $FailedItem"
    LogWrite " || $ErrorMessage"
    LogWrite " || Exiting script"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Error changing registry key $FailedItem" -ForegroundColor Red
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| $ErrorMessage"
    Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Exiting script"
}


Write-Host (Get-Date).toString("MM/dd/yyyy HH:mm:ss") "|| Script Finished!" -ForegroundColor Green