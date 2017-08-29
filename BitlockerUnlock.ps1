#BitlockerUnlock.ps1
#Last Updated: 7/21/2017
#Decription:
#Unlock a Bitlocker encrypted drive
param([string]$blKey, [string]$drive)

Import-Module -Name "$($pwd)\Bitlocker"

#$SecureString = ConvertTo-SecureString $blKey -AsPlainText -Force

Unlock-BitLocker -MountPoint "$($drive):" -RecoveryPassword $blKey