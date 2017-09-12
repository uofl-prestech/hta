#BitlockerUnlock.ps1
#Last Updated: 7/21/2017
#Decription:
#Unlock a Bitlocker encrypted drive
param([string]$blKey, [string]$drive)
$Error.clear()
Import-Module -Name "$($pwd)\Bitlocker" -WarningAction silentlyContinue

#$SecureString = ConvertTo-SecureString $blKey -AsPlainText -Force

Unlock-BitLocker -MountPoint "$($drive):" -RecoveryPassword $blKey

Foreach ($Errors in $Error) {
    $Errors.ToString()
}