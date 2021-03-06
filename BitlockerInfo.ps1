#BitlockerInfo.ps1
#Last Updated: 7/21/2017
#Decription:
#Collect encryption information and volume information for each volume on the computer
#and combine all of the information into on ArrayList of objects for output as HTML
#Get the ID and security principal of the current user account
param (
    [string]$directory = ""
 )
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
 
# Get the security principal for the Administrator role
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
 
# Check to see if we are currently running "as Administrator"
if ($myWindowsPrincipal.IsInRole($adminRole))
   {
		# We are running "as Administrator" - so change the title and background color to indicate this
		$Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
		$Host.UI.RawUI.BackgroundColor = "Red"
		Set-Location $directory
   }
else
   {
		# Start the new process
		$startInfo = New-Object System.Diagnostics.ProcessStartInfo
    	$startInfo.FileName = "powershell.exe"
		$startInfo.Arguments = "-executionpolicy bypass -file $($myInvocation.MyCommand.Definition) -directory $($PWD)"
    	$startInfo.Verb = "Runas"

		$process = New-Object System.Diagnostics.Process
		$process.StartInfo = $startInfo
		$process.Start()
		$process.WaitForExit()

    	#$User2 = Read-Host -Prompt "Input the user name $($PWD)"
		# Exit from the current, unelevated, process
		#$driveTable = Get-Content bitlockerinfo.html
		return
   }
# Run your code that needs to be elevated here
#Write-Host -NoNewLine "Running as Admin"
#$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

#*****************Bitlocker Info*****************
#Bitlocker Information
	$WMIVolumes = Get-WMIobject -namespace "Root\cimv2\security\MicrosoftVolumeEncryption" -ClassName "Win32_EncryptableVolume" -Credential $Credential

#Create an ArrayList variable named $encryptedVolumes to hold encryption information for each drive
	[System.Collections.ArrayList]$encryptedVolumes = @()

	foreach($volume in $WMIVolumes)
	{
	
#Create a PSObject named $driveObject to hold Drive Letter, Encryption Status, Key Type, and Key ID
		$driveObject = New-Object PSObject
		$driveObject | Add-Member -MemberType NoteProperty -Name DriveLetter -Value $volume.DriveLetter
		$Estatus = ""
		
		
		switch ($volume.ProtectionStatus)
		{
			"0" {$Estatus = "FullyDecrypted"}
			"1" {$Estatus = "FullyEncrypted"}
			"2" {$Estatus = "EncryptionInProgress"}
			"3" {$Estatus = "DecryptionInProgress"}
			"4" {$Estatus = "EncryptionPaused"}
			"5" {$Estatus = "DecryptionPaused"}
			default {$Estatus = "Error finding encryption state"}
		}
		
		
		$ProtectorIds = $volume.GetKeyProtectors("0").volumekeyprotectorID
		foreach ($ProtectorID in $ProtectorIds)
		{
		
			$KeyProtectorType = $volume.GetKeyProtectorType($ProtectorID).KeyProtectorType
			$KeyType = ""
			switch($KeyProtectorType)
			{
				"0"{$KeyType = "Unknown or other protector type";break}
				"1"{$KeyType = "Trusted Platform Module (TPM)";break}
				"2"{$KeyType = "External key";break}
				"3"{$KeyType = "Recovery Key ID";break}
				"4"{$KeyType = "TPM And PIN";break}
				"5"{$KeyType = "TPM And Startup Key";break}
				"6"{$KeyType = "TPM And PIN And Startup Key";break}
				"7"{$KeyType = "Public Key";break}
				"8"{$KeyType = "Passphrase";break}
				"9"{$KeyType = "TPM Certificate";break}
				"10"{$KeyType = "CryptoAPI Next Generation (CNG) Protector";break}
			}
			$driveObject | Add-Member -MemberType NoteProperty -Name $KeyType -Value $ProtectorID
			
		}
		$driveObject | Add-Member -MemberType NoteProperty -Name EncryptionState -Value $Estatus

#Add $driveObject to $encryptedVolumes
		[void]$encryptedVolumes.Add($driveObject)

	}




#*****************Drive Info*****************
#Hard Drive information
	[System.Collections.ArrayList]$driveResults = Get-WmiObject win32_Volume | select  Name, DriveType, Freespace, Capacity, DriveLetter

#Create an ArrayList named $driveList to hold drive and encryption information for each drive
	[System.Collections.ArrayList]$driveList = @()


	foreach ($drive in $driveResults)
	{
		#Skip volumes with no drive letter
		If($drive.DriveLetter -eq $null) {continue}
#Modify Drive Type
		If($drive.DriveType -like "3") {$drive.DriveType = "Local Disk"}
		If($drive.DriveType -like "5") {$drive.DriveType = "Compact Disk"}
		If($drive.DriveType -like "6") {$drive.DriveType = "RAM"}
		If($drive.DriveType -like "4") {$drive.DriveType = "Network Drive"}
		If($drive.DriveType -like "2") {$drive.DriveType = "Removable Disk"}
		If($drive.DriveType -like "1") {$drive.DriveType = "No Root Directory"}
		If($drive.DriveType -like "0") {$drive.DriveType = "Unknown"}

#Modify Drive Name by replacing ":\" with ":"
		$drive.Name = $drive.Name -replace ":\\", ":"

		foreach($volume in $encryptedVolumes)
		{
		
#$driveResults will contain every drive the computer has available, whether it is encrypted or not
#$encryptedVolumes will only contain drives that can be encrypted
#While looping through $driveResults, compare the current drive to drives in $encryptedVolumes
#If the drive exists in both arrays, add the encryption info to the current $driveResults object
			If($drive.Name -eq $volume.DriveLetter)
			{
				$volume.psobject.properties | % {$drive | Add-Member -MemberType $_.MemberType -Name $_.Name -Value $_.Value -force}
			}
		}
		[void]$driveList.Add($drive)

	}

#Add additional members to each object in the $driveList array
	$driveList | Add-Member -MemberType ScriptProperty -Name FreeSpaceinGB -Value {ForEach-Object{[math]::Round(($this.freespace / 1GB),2)}}
	$driveList | Add-Member -MemberType ScriptProperty -Name CapacityinGB -Value {ForEach-Object{[math]::Round(($this.capacity / 1GB),2)}}
	$driveList = $driveList | select * -ExcludeProperty Freespace, Capacity, DriveLetter, "Trusted Platform Module (TPM)" | sort-object Name

	#Write-Output $driveList | Format-List
	$driveList | ForEach-Object {$_ | convertTo-HTML -Fragment -As "List" -PostContent "<br>"} > .\bitlockerinfo.txt
	return