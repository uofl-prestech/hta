On Error Resume Next
Set ProgressUI = CreateObject("Microsoft.SMS.TsProgressUI") 
ProgressUI.CloseProgressDialog 

'Set objects and declare global variables
Set env = CreateObject("Microsoft.SMS.TSEnvironment")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'**********************************************************************************************************************
'						        Create Log File
'**********************************************************************************************************************
CONST ForAppending = 8
Dim logShell, strLogDir, objFSO, htaLog
Set logShell = CreateObject("WScript.Shell")
strLogDir = logShell.currentDirectory
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set htaLog = objFSO.OpenTextFile(strLogDir & "\HTALOG.log", ForAppending, True)

htaLog.WriteLine(vbCrLf & "====================== Begin Logging at " & Now & " ======================")
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: pingTest
'**********************************************************************************************************************
Dim iTimerID
Dim pingShell, pingShellExec
Sub pingTest
    htaLog.WriteLine(Now & " ***** Begin Sub pingTest *****")

    Dim comspec, strObj
    dim pingTestDiv: set pingTestDiv = document.getElementById("div-output-network")
    pingTestDiv.innerHTML = "<p class='cmdHeading'>Network connectivity test: </p>"
    Set pingShell = CreateObject("WScript.Shell")
    comspec = pingShell.ExpandEnvironmentStrings("%comspec%")

    htaLog.WriteLine(Now & " || Executing command: cmd /c ping.exe www.google.com")

    Set pingShellExec = pingShell.Exec(comspec & " /c ping.exe www.google.com")
    iTimerID = window.setInterval("vbscript:writePing()", 10)
    
    htaLog.WriteLine(Now & " ***** End Sub pingTest *****")
End Sub

Sub writePing
    dim pingTestDiv: set pingTestDiv = document.getElementById("div-output-network")
    pingOutput = pingShellExec.StdOut.ReadLine()
    pingTestDiv.innerHTML = pingTestDiv.innerHTML & pingOutput & "<br>"
    htaLog.WriteLine(Now & " || " & pingOutput)
	If pingShellExec.Status = 1 Then
        window.clearInterval(iTimerID)
        pingOutput = pingShellExec.StdOut.ReadAll()
        pingTestDiv.innerHTML = pingTestDiv.innerHTML & pingOutput & "<br>"
        htaLog.WriteLine(Now & " || " & pingOutput)
	End If

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: listDrives
'**********************************************************************************************************************
Function listDrives
    On Error Resume Next
    htaLog.WriteLine(Now & " ***** Begin Sub listDrives *****")
    Dim strComputer, objWMIService, colItems, admin, drivesObj
	Dim landingPageDiv: Set landingPageDiv = document.getElementById("bl-info-output")
    strComputer = "."
    Set drivesObj = CreateJsObj()

    'Quick check to see if we are running as admin
    CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    If Err.number = 0 Then 
        admin = true
        htaLog.WriteLine(Now & " || User is running script as admin")
    Else
        admin = false
        htaLog.WriteLine(Now & " || Error. Must run as Administrator to check encryption status")
    End If
    Err.Clear
    
    'Check drives for encryption if running as admin
    If admin = true Then
        Dim arEncryptionMethod
        arEncryptionMethod = Array("None", "AES 128 With Diffuser", "AES 256 With Diffuser", "AES 128", "AES 256", "Hardware Encryption", "XTS AES 128", "XTS AES 256", "Unknown")
        Dim arProtectionStatus
        arProtectionStatus = Array("Protection Off", "Protection On", "Protection Unknown")
        Dim arConversionStatus
        arConversionStatus = Array("Fully Decrypted", "Fully Encrypted", "Encryption In Progress", "Decryption In Progress", "Encryption Paused", "Decryption Paused")
        Dim arLockStatus
        arLockStatus = Array("Unlocked", "Locked")
        Dim arKeyType
        arKeyType = Array("Unknown or other protector type", "Trusted Platform Module (TPM)", "External key", "Numerical password", "TPM And PIN", "TPM And Startup Key",_
         "TPM And PIN And Startup Key", "Public Key", "Passphrase", "TPM Certificate", "CryptoAPI Next Generation (CNG) Protector")

        htaLog.WriteLine(Now & " || Executing command: GetObject(""winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2\Security\MicrosoftVolumeEncryption"")")
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2\Security\MicrosoftVolumeEncryption")

        htaLog.WriteLine(Now & " || Executing command: objWMIService.ExecQuery(""Select * from Win32_EncryptableVolume"",,48"")")
        Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_EncryptableVolume",,48)

        htaLog.WriteLine(Now & " || Win32_EncryptableVolume instance")
        htaLog.WriteLine(Now & " || 0=Protection OFF, 1= Protection ON, unlocked, 2=Protection ON, locked")
        For Each objItem in colItems
            'Only include drive if it has a drive letter
            If objItem.DriveLetter Then
            On Error Resume Next
                Dim EncryptionMethod, ProtectionStatus, ConversionStatus, EncryptionPercentage, VolumeKeyProtectorID, LockStatus, KeyType, driveInfo
                objItem.GetEncryptionMethod EncryptionMethod
                objItem.GetProtectionStatus ProtectionStatus
                objItem.GetConversionStatus ConversionStatus, EncryptionPercentage
                objItem.GetKeyProtectors 0, VolumeKeyProtectorID

                objItem.GetLockStatus LockStatus
                driveInfo = Array(ProtectionStatus, LockStatus, EncryptionMethod, ConversionStatus, EncryptionPercentage)

                objItem.DriveLetter = Replace(objItem.DriveLetter, ":", "")
                htaLog.Write(Now & " || Drive Letter: " & objItem.DriveLetter)
                htaLog.WriteLine(" || ProtectionStatus: " & objItem.ProtectionStatus)
                drivesObj.setProp objItem.DriveLetter, "Drive Letter", objItem.DriveLetter

                drivesObj.setProp objItem.DriveLetter, "Protection Status", arProtectionStatus(ProtectionStatus)
                drivesObj.setProp objItem.DriveLetter, "Encryption Method", arEncryptionMethod(EncryptionMethod)
                drivesObj.setProp objItem.DriveLetter, "Lock Status", arLockStatus(LockStatus)

                'Find Key Protector and Key ID
                For Each objId in VolumeKeyProtectorID
                    Dim VolumeKeyProtectorType
                    objItem.GetKeyProtectorType objId, VolumeKeyProtectorType
                    If VolumeKeyProtectorType <> "" Then
                        drivesObj.setProp objItem.DriveLetter, arKeyType(VolumeKeyProtectorType), objId
                    End If
                Next
            End If
        Next
    Else
        htaLog.WriteLine(Now & " || Skipping check for Encryption Status")
    End If

    htaLog.WriteLine(Now & " || Executing command: GetObject(""winmgmts:\\"" & strComputer & ""\root\cimv2"")")

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    
    htaLog.WriteLine(Now & " || Executing command: objWMIService.ExecQuery(""Select * from Win32_Volume"")")
    
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Volume")
	Dim arDriveTypes
    arDriveTypes = Array("Unknown", "No Root Directory", "Removable Disk", "Local Disk", "Network Drive", "Compact Disk", "RAM")
    For Each objItem In colItems
        'Only include drive if it has a drive letter

        If objItem.DriveLetter Then
            objItem.DriveLetter = Replace(objItem.DriveLetter, ":", "")
            Dim fso, checkForWindows
            checkForWindows = objItem.DriveLetter & ":\Windows"
            Set fso = CreateObject("Scripting.FileSystemObject")
            If(fso.FolderExists(checkForWindows) AND objItem.DriveLetter <> "X") Then
                htaLog.WriteLine(Now & " || Windows Drive found at: " & checkForWindows)
                document.getElementById("windows-drive-letter").Value = objItem.DriveLetter
                objItem.Label = "Windows Drive!"
            End If
            
            htaLog.Write(Now & " || ""Drive Letter: " & objItem.DriveLetter & " | ")
            drivesObj.setProp objItem.DriveLetter, "Label", objItem.Label
            htaLog.Write(Now & " || ""Label: " & objItem.Label & " | ")
            fsCapacity = ConvertSize(objItem.Freespace)
            htaLog.Write("Free Space: " & fsCapacity)
            drivesObj.setProp objItem.DriveLetter, "Free Space", fsCapacity
            capacity = ConvertSize(objItem.Capacity)
            htaLog.Write("Capacity: " & ConvertSize(objItem.Capacity))
            drivesObj.setProp objItem.DriveLetter, "Capacity", ConvertSize(objItem.Capacity)
            htaLog.Write("Drive Type: " & arDriveTypes(objItem.DriveType))
            drivesObj.setProp objItem.DriveLetter, "Drive Type", arDriveTypes(objItem.DriveType)
        End If
	Next

    htaLog.WriteLine(Now & " ***** End Sub listDrives *****")
    Set listDrives = drivesObj
End Function
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: mapNetDrive
'**********************************************************************************************************************
Sub mapNetDrive
	Dim strUser, strPass, strSharePath, strDriveLetter, objDrives, objNetwork, networkMapDiv
	strUser = document.getElementById("input-share-username").Value
    strPass = document.getElementById("input-share-password").Value
    strSharePath = document.getElementById("input-share-path").Value
	Set networkMapDiv = document.getElementById("div-output-network")
    networkMapDiv.innerHTML = ""

    If (strUser <> "") and (strPass <> "") and (strSharePath <> "") Then
		Set objNetwork = CreateObject("WScript.Network")
		On Error Resume Next
		objNetwork.RemoveNetworkDrive "N:"
		objNetwork.MapNetworkDrive "N:", strSharePath, "False", strUser, strPass
		Set objDrives = objNetwork.EnumNetworkDrives
	
		'Error traps
		If Err <> 0 Then
            Select Case Err.Number
                ' This case is if we try to remove a mapping that isn't there. Can be ignored
                Case -2147022646
                    
                ' Add case if error is permissions or no network connection etc
                Case Else
                    networkMapDiv.innerHTML = "Error " & Err.Number & ": " & Err.Description & "<br>"
                    Exit Sub
            End Select
		End If
		
		networkMapDiv.innerHTML = "<h2>Drives currently mapped:<h2>"
		For i = 0 To objDrives.Count-1
			networkMapDiv.innerHTML = networkMapDiv.innerHTML & objDrives.Item(i) & "<br>"
        Next
    Else
        networkMapDiv.innerHTML = "<h2>Username, Password, or Network Path not filled in<h2>"
	End If
	
End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: enumUsers
'**********************************************************************************************************************
Sub enumUsers
    htaLog.WriteLine(Now & " ***** Begin Sub enumUsers *****")

    Const HKLM = &H80000002
    Dim htmlString, strComputer, strHivePath, strKeyPath, strSubKeyPath, profilePath, userName, selectLength, strSourceDrive
    Dim userNameDiv: Set userNameDiv = document.getElementById("div-select-users")
    userNameDiv.innerHTML = ""
    set objshell = CreateObject("Wscript.shell")
    strComputer = "."
    strSourceDrive = document.getElementById("input-windows-drive").Value
    strHivePath = strSourceDrive & ":\Windows\System32\Config\SOFTWARE"
    strKeyPath = "TempSoftware\Microsoft\Windows NT\CurrentVersion\ProfileList\"

    htaLog.WriteLine(Now & " || strComputer = "".""")
    htaLog.WriteLine(Now & " || strHivePath = " & strHivePath)
    htaLog.WriteLine(Now & " || strKeyPath = ""TempSoftware\Microsoft\Windows NT\CurrentVersion\ProfileList""")
    
    htaLog.WriteLine(Now & " || Executing command: ""reg.exe load HKLM\TempSoftware " & strHivePath & ", 0, true")

    On Error Resume Next
    Err.Clear

    objshell.Run "reg.exe load HKLM\TempSoftware " & strHivePath, 0, true

    If Err <> 0 Then
        htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
        htaLog.WriteLine(Now & " || Error Description: " & Err.Description)
        Err.Clear
    End If
    
    htaLog.WriteLine(Now & " || Executing command: objshell.regRead(""HKLM\" & strKeyPath & ")")

    objshell.regRead("HKLM\" & strKeyPath)

    If Err <> 0 Then
        htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
        htaLog.WriteLine(Now & " || Error Description: " & Err.Description)
        htaLog.WriteLine(Now & " || Could not load HKLM\TempSoftware")
        userNameDiv.innerHTML = userNameDiv.innerHTML & "Could not load offline registry<br><br>"
        strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\ProfileList"

        htaLog.WriteLine(Now & " || strKeyPath = ""Software\Microsoft\Windows NT\CurrentVersion\ProfileList""")

        Err.Clear
    End If

    htaLog.WriteLine(Now & " || Executing command: GetObject(""winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv"")")
    Set objReg = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv")

    htaLog.WriteLine(Now & " || Executing command: objReg.EnumKey HKLM, "& strKeyPath & ", arrSubKeys")
    objReg.EnumKey HKLM, strKeyPath, arrSubKeys

    selectLength = 0
    htaLog.WriteLine(Now & " || Looping through registry subkeys to extract usernames that exist in this Windows installation")
    On Error Resume Next
    For Each subkey In arrSubKeys
        strSubKeyPath = "HKEY_LOCAL_MACHINE\" & strKeyPath & "\" & subkey & "\ProfileImagePath"
        htaLog.WriteLine(Now & " || strSubKeyPath = " & strSubKeyPath)
        profilePath = objshell.regRead(strSubKeyPath)
        htaLog.WriteLine(Now & " || profilePath = " & profilePath)
        profilePath = Split(profilePath, "\")
        userName = profilePath(Ubound(profilePath))
        htaLog.WriteLine(Now & " || userName = " & userName)

        If Err <> 0 Then
            htaLog.WriteLine(Now & " || Error - Cannot find user. Is the drive still encrypted?")
            Err.Clear
        End If

        If NOT (strcomp(userName,"systemprofile",0) = 0 OR strcomp(userName,"LocalService",0) = 0 OR strcomp(userName,"NetworkService",0) = 0 _
            OR strcomp(userName,"defaultuser0",0) = 0 OR strcomp(userName,"sccmpush",0) = 0) Then
            htmlString = htmlString & "<option value=" & subkey & ">" & userName & "</option>"
            env("env" & userName & "SID") = subkey
            selectLength = selectLength + 1
        End If
    Next
    userNameDiv.innerHTML = userNameDiv.innerHTML & "<span>Users: &nbsp&nbsp </span><br>"
    userNameDiv.innerHTML = userNameDiv.innerHTML & "<select id=""input-usmt-usernames"" name=""usmtUsernameList"" size="& selectLength &" multiple>" & htmlString & "</select>"
    
    htaLog.WriteLine(Now & " || Executing command: objShell.Run(""cmd /c reg.exe unload HKLM\TempSoftware"", 0, true)")
    objshell.Run "%comspec% /c reg.exe unload HKLM\TempSoftware", 0, true

    htaLog.WriteLine(Now & " ***** End Sub enumUsers *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: cmdPrompt
'**********************************************************************************************************************
Sub cmdPrompt
    htaLog.WriteLine(Now & " ***** Begin Sub cmdPrompt *****")
	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: cmd /k")
    cmdShell.Run "cmd /k"
    htaLog.WriteLine(Now & " ***** End Sub cmdPrompt *****")
End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: psPrompt
'**********************************************************************************************************************
Sub psPrompt
    htaLog.WriteLine(Now & " ***** Begin Sub psPrompt *****")
	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: Powershell.exe -noprofile -noexit -executionpolicy bypass")

    cmdShell.Run "Powershell.exe -noprofile -noexit -executionpolicy bypass"

    htaLog.WriteLine(Now & " ***** End Sub psPrompt *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: logViewer
'**********************************************************************************************************************
Sub logViewer
    htaLog.WriteLine(Now & " ***** Begin Sub logViewer *****")

	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: .\tools\cmtrace64.exe")
    cmdShell.Run ".\tools\cmtrace64.exe"
    
    htaLog.WriteLine(Now & " ***** End Sub logViewer *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: notepadPP
'**********************************************************************************************************************
Sub notepadPP
    htaLog.WriteLine(Now & " ***** Begin Sub notepadPP *****")

	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: .\tools\npp\notepadpp.exe")
    cmdShell.Run ".\tools\npp\notepadpp.exe"
    
    htaLog.WriteLine(Now & " ***** End Sub notepadPP *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: explorer
'**********************************************************************************************************************
Sub explorer
    htaLog.WriteLine(Now & " ***** Begin Sub explorer *****")

	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")
    cmdShell.Run ".\tools\explorerpp.exe"
    
    htaLog.WriteLine(Now & " ***** End Sub explorer *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: FTK
'**********************************************************************************************************************
Sub FTK
    htaLog.WriteLine(Now & " ***** Begin Sub FTK *****")
	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: "".\tools\FTK\FTKImager.exe""")
    cmdShell.Run """.\tools\FTK\FTKImager.exe"""
    htaLog.WriteLine(Now & " ***** End Sub FTK *****")
End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: WMIExplorer
'**********************************************************************************************************************
Sub WMIExplorer
    htaLog.WriteLine(Now & " ***** Begin Sub WMIExplorer *****")
	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: .\tools\WMIExplorer.exe")
    cmdShell.Run ".\tools\WMIExplorer.exe"
    htaLog.WriteLine(Now & " ***** End Sub WMIExplorer *****")
End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: ConvertSize(Size)
'**********************************************************************************************************************
Function ConvertSize(Size)
' Convert Bytes to KB, MB, GB, TB
	suffix = "0 Bytes" 
	If Size >= 1024 Then suffix = " KB" 
	If Size >= 1048576 Then suffix = " MB" 
	If Size >= 1073741824 Then suffix = " GB" 
	If Size >= 1099511627776 Then suffix = " TB" 
	 
	Select Case Suffix 
	    Case " KB" Size = Round(Size / 1024, 1) 
	    Case " MB" Size = Round(Size / 1048576, 1) 
	    Case " GB" Size = Round(Size / 1073741824, 1) 
	    Case " TB" Size = Round(Size / 1099511627776, 1) 
	End Select
    ConvertSize = Size & Suffix
End Function
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: ButtonFinishClick
'**********************************************************************************************************************
Sub ButtonFinishClick
' Execute OSD subroutine
    htaLog.WriteLine(Now & " ***** Begin Sub ButtonFinishClick *****")

    iAdobeReaderDC = document.getElementById("AdobeReaderDC").Checked
    iAdobeAcrobatProXI = document.getElementById("AdobeAcrobatProXI").Checked
    iChrome = document.getElementById("Chrome").Checked
    iFileMaker = document.getElementById("FileMaker").Checked
    iOffice2016 = document.getElementById("Office2016").Checked
    iOSD = document.getElementById("input-osd-checkbox").Checked
    iMBAM = document.getElementById("MBAM").Checked
    ifnfOsdCheckBox = document.getElementById("input-fnfosd-checkbox").Checked
    icompName = document.getElementById("input-comp-name").Value

    ' Set value of variable to true/false based on whether the checkbox is selected or not
    htaLog.Write(Now & " || Adobe Reader DC checked?")
    htaLog.WriteLine(" || AdobeReaderDC.Checked = " & iAdobeReaderDC)

    If iAdobeReaderDC Then
        strAdobeReaderDC = "true"
        else strAdobeReaderDC = "false"
    End If

    htaLog.Write(Now & " || Adobe Acrobat Pro XI checked?")
    htaLog.WriteLine(" || AdobeAcrobatProXI.Checked = " & iAdobeAcrobatProXI)

	If iAdobeAcrobatProXI Then
		strAdobeAcrobatProXI = "true"
		else strAdobeAcrobatProXI="false"
	End If		

    htaLog.Write(Now & " || Chrome checked?")
    htaLog.WriteLine(" || Chrome.Checked = " & iChrome)

    If iChrome Then
        strChrome = "true"
        else strChrome = "false"
    End If

    htaLog.Write(Now & " || FileMaker checked?")
    htaLog.WriteLine(" || FileMaker.Checked = " & iFileMaker)

	If iFileMaker Then
        strFileMaker = "true"
        else strFileMaker = "false"
    End If

    htaLog.Write(Now & " || Office2016 checked?")
    htaLog.WriteLine(" || Office2016.Checked = " & iOffice2016)

    If iOffice2016 Then
        strOffice2016 = "true"
        else strOffice2016 = "false"
    End If

    htaLog.Write(Now & " || OSD checked?")
    htaLog.WriteLine(" || OSD.Checked = " & iOSD)

    If iOSD OR ifnfOsdCheckBox Then
        strOSD = "true"
        iMBAM = "true"
        strCompName = icompName
        else strOSD = "false"
    End If

    htaLog.Write(Now & " || MBAM checked?")
    htaLog.WriteLine(" || MBAM.Checked = " & iMBAM)
    If iMBAM Then
        strMBAM = "true"
        else strMBAM = "false"
    End If
	
    ' Set value of variables that will be used by the task sequence, then close the window and allow the task sequence to continue.
    'MsgBox "Set Environment Variables" 
    
    htaLog.WriteLine(Now & " || Environment Variables:")

    env("envAdobeReaderDC") = strAdobeReaderDC
    env("envAdobeAcrobatProXI") = strAdobeAcrobatProXI
    env("envChrome") = strChrome
    env("envFileMaker") = strFileMaker
    env("envOffice2016") = strOffice2016
    env("envOSD") = strOSD
    env("OSDComputerName") = strCompName
    env("envMBAM") = strMBAM

    'For each v in env.GetVariables 
    '    htaLog.WriteLine(v & " = " & env(v)) 
    'Next 
    'MsgBox "End Finish Subroutine"

    window.close

    htaLog.WriteLine(Now & " ***** End Sub ButtonFinishClick *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: dismCapture
'**********************************************************************************************************************
Function dismCapture
' DISM Capture Image subroutine
    htaLog.WriteLine(Now & " ***** Begin Sub dismCapture *****")

	Dim dismShell, strName, destPath, sourcePath, returnCode, wimPath, fso
	Dim dismDiv: Set dismDiv = document.getElementById("dism-output")
    strSourcePath = document.getElementById("input-windows-drive").Value
    strDestPath = document.getElementById("input-external-drive").Value
    strName = document.getElementById("input-primary-username").Value
    Set dismShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    wimPath = strDestPath & ":\" & strName & ".wim"
    returnCode = -1

    'Check if file with this name already exists
    If (fso.FileExists(wimPath)) Then
        MsgBox wimPath & " already exists. Try changing the username or deleting the existing wim file."
        dismCapture = returnCode
        Exit Function
    End If

    htaLog.WriteLine(Now & " || strSourcePath = " & strSourcePath)
    htaLog.WriteLine(Now & " || strDestPath = " & strDestPath)
    htaLog.WriteLine(Now & " || strName = " & strName)

    dismDiv.innerHTML = "Running Command: X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /LogPath:X:\dism.log"
    
    htaLog.Writeline(Now & " || returnCode = dismShell.run (""cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /LogPath:X:\dism.log"", 1, True)")
    
    returnCode = dismShell.run("cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /LogPath:X:\dism.log", 1, True)

    ' /ScratchDir:"&strDestPath&":\

    dismDiv.innerHTML = "Capture Finished! <br><br> Return Code: " & returnCode

    htaLog.WriteLine(Now & " || Capture finished!")
    htaLog.WriteLine(Now & " || Return Code: " & returnCode)

    htaLog.WriteLine(Now & " ***** End Sub dismCapture *****")
    dismCapture = returnCode

End Function
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: runDISM_TS
'**********************************************************************************************************************
Sub runDISM_TS
' Set environment variables for running DISM during task sequence
    htaLog.WriteLine(Now & " ***** Begin Sub runDISM_TS *****")

    htaLog.WriteLine(Now & " || Setting Environment Variables")

    env("envDismSelected") = "true"
    env("envWindowsDrive") = document.getElementById("input-windows-drive").Value
    env("envBitlockerKey") = document.getElementById("input-bitlocker-key").Value
    env("envExternalDrive") = document.getElementById("input-external-drive").Value
    env("envPrimaryUsername") = document.getElementById("input-primary-username").Value

    htaLog.WriteLine(Now & " || env(""envDismSelected"") = " & env("envDismSelected"))
    htaLog.WriteLine(Now & " || env(""envWindowsDrive"") = " & env("envWindowsDrive"))
    htaLog.WriteLine(Now & " || env(""envBitlockerKey"") = " & env("envBitlockerKey"))
    htaLog.WriteLine(Now & " || env(""envExternalDrive"") = " & env("envExternalDrive"))
    htaLog.WriteLine(Now & " || env(""envPrimaryUsername"") = " & env("envPrimaryUsername"))

    htaLog.WriteLine(Now & " ***** End Sub runDISM_TS *****")

    window.close

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: copyWallpaper
'**********************************************************************************************************************
Sub copyWallpaper
    htaLog.WriteLine(Now & " ***** Begin Sub copyWallpaper *****")

    Dim strComputer, strHivePath, strKeyPath, userName, objshell, getUser, getUserSID, strWallpaperPath, strWallpaperRegPath,strSourceDrive, strDestDrive
    set objshell = CreateObject("Wscript.shell")
    getUser = document.getElementById("input-primary-username").Value
	htaLog.WriteLine("getUser = " & getUser) 
    getUserSID = env("env" & getUser & "SID")
    htaLog.WriteLine("getUserSID = " & getUserSID)

    On Error Resume Next
    Err.Clear
    strComputer = "."
    strSourceDrive = document.getElementById("input-windows-drive").Value
    strDestDrive = document.getElementById("input-external-drive").Value
    strHivePath = strSourceDrive & ":\Users\"& getUser & "\NTUSER.DAT"
    strWallpaperRegPath = "HKEY_USERS\TempUser\Control Panel\Desktop\WallPaper"

    htaLog.WriteLine(Now & " || strComputer = "".""")
    htaLog.WriteLine(Now & " || strHivePath = " & strHivePath)
    htaLog.WriteLine(Now & " || strWallpaperPath = ""TempUser\Microsoft\Windows NT\CurrentVersion\ProfileList""")
    htaLog.WriteLine(Now & " || Executing command: ""reg.exe load HKEY_USERS\TempUser " & strHivePath & ", 0, true")

    objshell.Run "reg.exe load HKEY_USERS\TempUser " & strHivePath, 0, true

    If Err <> 0 Then
        'Couldn't load NTUSER.DAT file
        htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
        htaLog.WriteLine(Now & " || Error Description: " & Err.Description)
        Exit Sub
    End If
    
    htaLog.WriteLine(Now & " || Executing command: strWallpaperPath = objshell.regRead(strWallpaperRegPath)")

    strWallpaperPath = objshell.regRead(strWallpaperRegPath)

    If Err <> 0 Then
        'Couldn't find Wallpaper registry key at the specified location
        htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
        htaLog.WriteLine(Now & " || Error Description: " & Err.Description)
        htaLog.WriteLine(Now & " || Could not load " & strWallpaperRegPath)
        Exit Sub
    Else
        Dim fsoWallPaper, fsoExtension, strNewPath
        Set fsoWallPaper = CreateObject("Scripting.FileSystemObject")
        htaLog.WriteLine(Now & " || Executing Command: fsoWallPaper.CopyFile " & strWallpaperPath & ", "& strDestDrive & ":\USMT\" & getUser & "\")
		fsoExtension = fsoWallPaper.GetExtensionName(strWallpaperPath)
		strNewPath = strDestDrive & ":\USMT\" & getUser & "\" & getUser & "." & fsoExtension
        fsoWallPaper.CopyFile strWallpaperPath, strNewPath
		env("envWallpaperFile") = strNewPath
    End If

    htaLog.WriteLine(Now & " ***** End Sub copyWallpaper *****")
End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: usmtScanstate(buttonClicked)
'**********************************************************************************************************************
Function usmtScanstate(buttonClicked)
    htaLog.WriteLine(Now & " ***** Begin Sub usmtScanstate(buttonClicked) *****")

    Dim getUser, WshShell, strCurrentDir, destDrive, scanStateDiv, returnCode, userArray, userArraySize, userIncludeString, windowsDrive
    userArray = Array()
    Set scanStateDiv = document.getElementById("scanstate-output")
    Set WshShell = CreateObject("WScript.Shell")
    strCurrentDir = WshShell.currentDirectory
    returnCode = 1
    Err.Clear

    htaLog.WriteLine(Now & " || strCurrentDir = " & strCurrentDir)
    htaLog.WriteLine(Now & " || usmtUsernameList.Options.Length = " & usmtUsernameList.Options.Length)

    For i = 0 to (usmtUsernameList.Options.Length - 1)
        htaLog.WriteLine(Now & " || usmtUsernameList.Options("&i&") = " & usmtUsernameList.Options(i).Value & ", Selected = " & usmtUsernameList.Options(i).Selected)
        If (usmtUsernameList.Options(i).Selected) Then
            ReDim Preserve userArray(UBound(userArray) + 1)
            userArray(UBound(userArray)) = usmtUsernameList.Options(i).Value
            userIncludeString = userIncludeString & "/ui:" & usmtUsernameList.Options(i).Value & " "
        End If
    Next
    htaLog.WriteLine(Now & " || userIncludeString = " & userIncludeString)
    userArraySize = Ubound(userArray)

	getUser = document.getElementById("input-primary-username").Value
	destDrive = document.getElementById("input-external-drive").Value
    strSourcePath = document.getElementById("input-windows-drive").Value

    htaLog.WriteLine(Now & " || usmtUsername.Value = " & getUser)
    htaLog.WriteLine(Now & " || usmtDrive.Value = " & destDrive)

    scanStateDiv.innerHTML = "USMT Command that will execute: <br><br>" & strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c <br> /offline:" & strCurrentDir & "\USMT\offline.xml <br> /i:" & strCurrentDir & "\USMT\migdocs.xml <br> /i:" & strCurrentDir & "\USMT\migapp.xml <br> /i:" & strCurrentDir & "\USMT\oopexcludes.xml <br> /progress:" & strCurrentDir & "\prog.log <br> /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log <br> /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 <br> /ue:* " & userIncludeString & ", 1, True"

    htaLog.WriteLine(Now & " || Execute Scanstate if buttonClicked = True, getUser is not blank, and destDrive is not blank")
    htaLog.WriteLine(Now & " || buttonClicked = " & buttonClicked & ", getUser = " & getUser & ", destDrive = " & destDrive)

    If buttonClicked = "True" AND getUser <> "" AND destDrive <> "" Then
        htaLog.WriteLine(Now & " || Executing command: WshShell.Run (""cmd /c "&strCurrentDir&"\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c /o /offline:USMT\offline.xml /i:USMT\migdocs.xml /i:USMT\migapp.xml /i:USMT\oopexcludes.xml /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 /ue:* "&userIncludeString&", 1, True)")

        returnCode = WshShell.Run ("cmd /c " & strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c /o /offline:USMT\offline.xml /i:USMT\migdocs.xml /i:USMT\migapp.xml /i:USMT\oopexcludes.xml /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 /ue:* " & userIncludeString, 1, True)

        If returnCode = 0 Then
            htaLog.WriteLine(Now & " || Scanstate Complete!")
            htaLog.WriteLine(Now & " || Copying WallPaper file from C:\Users\<getUser>\ to <windowsDrive>:\USMT\<getUser>\")
            copyWallpaper
            scanStateDiv.innerHTML = "Scanstate Complete! <br> Log files can be found in "&destDrive&":\USMT\"&getUser&"\"
        Else
            htaLog.WriteLine(Now & " || Scanstate Failed! Return Code: " & returnCode)
            htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
            htaLog.WriteLine(Now & " || Error Description: " & Err.Description)

            scanStateDiv.innerHTML = "Error Code: " & returnCode
            
            Set fso = CreateObject("Scripting.FileSystemObject")
            fileName = destDrive & ":\USMT\" & getUser & "\scanstate.log"
            'fileName = "A:\prestechHTA\prestech\USMT\scanstate.log"
            Set myFile = fso.OpenTextFile(fileName, 1)
            Do While myFile.AtEndOfStream <> True
                textLine = myFile.ReadLine
                strRead = strRead & textLine & "<br>"
            Loop
            myFile.Close
            htaLog.WriteLine(Now & " || Scanstate log can be found at " & fileName)
            scanStateDiv.innerHTML = scanStateDiv.innerHTML & "<br><br> Scanstate Failed! <br><br>" & strRead
        End If
    Else
        htaLog.WriteLine(Now & " || Scanstate Skipped!")
    End If

    htaLog.WriteLine(Now & " ***** End Sub usmtScanstate(buttonClicked) *****")
    usmtScanstate = returnCode
End Function
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: usmtLoadstate
'**********************************************************************************************************************
Sub usmtLoadstate
    htaLog.WriteLine(Now & " ***** Begin Sub usmtLoadstate *****")

    Dim ReturnCode, getUser, usmtDrive, strCurrentDir
    Dim objShell : Set objShell = CreateObject("WScript.Shell")
    strCurrentDir = objShell.currentDirectory
    getUser = document.getElementById("input-primary-username").Value
    usmtDrive = document.getElementById("input-external-drive").Value

    htaLog.WriteLine(Now & " || strCurrentDir = " & strCurrentDir)
    htaLog.WriteLine(Now & " || getUser = " & getuser)
    htaLog.WriteLine(Now & " || usmtDrive = " & usmtDrive)
    htaLog.WriteLine(Now & " || Executing command: objShell.Run (""cmd /k " & strCurrentDir & "\USMT\loadstate.exe /c "&usmtDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&usmtDrive&":\USMT\"&getUser&"\loadstate.log"", 1, True)")

    ReturnCode = objShell.Run ("cmd /c " & strCurrentDir & "\USMT\loadstate.exe /c "&usmtDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&usmtDrive&":\USMT\"&getUser&"\loadstate.log", 1, True)

    If Err <> 0 Then
        htaLog.WriteLine(Now & " || Loadstate Error: " & ReturnCode)
        Exit Sub
    End If

    'Load user settings into registry from NTUSER.DAT file originally located at C:\Users\<username>
    'This file should be copied to <usmtDrive>:\USMT\<getUser>\NTUSER.DAT during scanstate and loaded to the registry from there.
    'Registry Key: Computer\HKEY_USERS\<SID>\Control Panel\Desktop\WallPaper

    Const HKLM = &H80000002
    Dim objReg, strComputer, strKeyPath, strValueName, strUserSID, strWallpaperFile, profilePath, strSubKeyPath, userName, strHivePath
    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\"
    strComputer = "."
    strHivePath = "C:\Users\"& getUser & "\NTUSER.DAT"
    htaLog.WriteLine(Now & " || strHivePath = " & strHivePath)
    strWallpaperRegPath = "TempUser\Control Panel\Desktop"
    strValueName = "WallPaper"
    
    'Try multiple extensions in case running Loadstate outside of the Flush and Fill Task Sequence (No env variable for wallpaper path would be set)
    On Error Resume Next
    Dim fsoWP, arrExtensions
    Set fsoWP = CreateObject("Scripting.FileSystemObject")

    strWallpaperFile = env("envWallpaperFile")
    fsoWP.GetFile(env("envWallpaperFile"))
    If Err = 0 Then
        htaLog.WriteLine(Now & " || TS env check worked")
        'strWallpaperFile = null
    Else
        htaLog.WriteLine(Now & " || TS env check failed")
        htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
        htaLog.WriteLine(Now & " || Error Description: " & Err.Description & vbCrLf)
        Err.Clear

        strWpPath = usmtDrive & ":\USMT\" & getUser & "\" & getUser
        arrExtensions = Array(".jpg", ".jpeg", ".png", ".gif", ".bmp")
        For Each extension In arrExtensions
            strFullPathTest = strWpPath & extension
            fsoWP.GetFile(strFullPathTest)
            If Err <> 0 Then
                htaLog.WriteLine(Now & " || Error. " & strFullPathTest & " not found.")
                htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
                htaLog.WriteLine(Now & " || Error Description: " & Err.Description & vbCrLf)
                Err.Clear
            Else
                htaLog.WriteLine(Now & " || Yay! " & strFullPathTest & " found!")
                strWallpaperFile = strFullPathTest
                Exit For
            End IF
        Next
    End If
	
    htaLog.WriteLine(Now & " || strWallpaperFile = " & strWallpaperFile)
    
    htaLog.WriteLine(Now & " || Executing command: ""reg.exe load HKLM\TempUser " & strHivePath & ", 0, true")
    Err.Clear
    objshell.Run "reg.exe load HKLM\TempUser " & strHivePath, 0, true

    If Err <> 0 Then
        htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
        htaLog.WriteLine(Now & " || Error Description: " & Err.Description)
        Err.Clear
    End If

    htaLog.WriteLine(Now & " || Executing command: GetObject(""winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv"")")
    Set objReg = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv")

    htaLog.WriteLine(Now & " || Executing command: objReg.SetStringValue HKLM, " & strWallpaperRegPath & ", " & strValueName & ", " & strWallpaperFile)
    objReg.SetStringValue HKLM, strWallpaperRegPath, strValueName, strWallpaperFile

    If Err = 0 Then
        objReg.GetStringValue HKLM, strWallpaperRegPath, strValueName, strWallpaperFile
        htaLog.WriteLine(Now & " ||  " & strWallpaperRegPath & "\" & strValueName & " contains: " & strWallpaperFile)
    Else 
        htaLog.WriteLine(Now & " ||  Error in creating key and REG_SZ value = " & Err.Number)
    End If

    htaLog.WriteLine(Now & " || Executing command: objShell.Run(""reg unload HKLM\TempUser"", 0, true)")
    objshell.Run "reg unload HKLM\TempUser", 0, true

    htaLog.WriteLine(Now & " ***** End Sub usmtLoadstate *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: usmtLoadstate_TS
'**********************************************************************************************************************
Function usmtLoadstate_TS
    htaLog.WriteLine(Now & " ***** Begin Sub usmtLoadstate_TS *****")

    env("envUsmtLoadstate") = "True"
    env("envWindowsDrive") = document.getElementById("input-windows-drive").Value
    env("envExternalDrive") = document.getElementById("input-external-drive").Value
    env("envPrimaryUsername") = document.getElementById("input-primary-username").Value

    htaLog.WriteLine(Now & " || env(""envUsmtLoadstate"") = " & env("envUsmtLoadstate"))
    htaLog.WriteLine(Now & " || env(""envWindowsDrive"") = " & env("envWindowsDrive"))
    htaLog.WriteLine(Now & " || env(""envExternalDrive"") = " & env("envExternalDrive"))
    htaLog.WriteLine(Now & " || env(""envPrimaryUsername"") = " & env("envPrimaryUsername"))

    htaLog.WriteLine(Now & " ***** End Sub usmtLoadstate_TS *****")

End Function
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: runFlushFill
'**********************************************************************************************************************
Sub runFlushFill
    htaLog.WriteLine(Now & " ***** Begin Sub runFlushFill *****")

    Dim dismError, scanstateError
    iWindowsDrive = document.getElementById("input-windows-drive").Value
    iExternalDrive = document.getElementById("input-external-drive").Value
    iPrimaryUsername = document.getElementById("input-primary-username").Value
    icompName = document.getElementById("input-comp-name").Value
    iDismCheckBox = document.getElementById("input-dism-checkbox").Checked
    iScanStateCheckBox = document.getElementById("input-scanstate-checkbox").Checked
    iLoadStateCheckBox = document.getElementById("input-loadstate-checkbox").Checked

    htaLog.WriteLine(Now & " || windowsDrive.Value = " & iWindowsDrive)
    htaLog.WriteLine(Now & " || externalDrive.Value = " & iExternalDrive)
    htaLog.WriteLine(Now & " || primaryUsername.Value = " & iPrimaryUsername)
    htaLog.WriteLine(Now & " || compName.Value = " & icompName)
    htaLog.WriteLine(Now & " || dismCheckBox.Checked = " & iDismCheckBox)
    htaLog.WriteLine(Now & " || scanStateCheckBox.Checked = " & iScanStateCheckBox)
    htaLog.WriteLine(Now & " || loadStateCheckBox.Checked = " & iLoadStateCheckBox)

    If iDismCheckBox Then
        htaLog.WriteLine(Now & " || dismCheckBox is checked. Run dismCapture routine")
        dismError = dismCapture
        If dismError Then
            MsgBox "DISM error " & dismError
            Exit Sub
        End If
    End If

    htaLog.WriteLine(Now & " || scanStateCheckBox.Checked = " & iScanStateCheckBox)

    If iScanStateCheckBox Then
        htaLog.WriteLine(Now & " || scanStateCheckBox is checked. Run usmtScanstate routine")
        usmtError = usmtScanstate("True")
        If usmtError Then
            MsgBox "Scanstate Failed with error " & usmtError & ". Halting Flush and Fill sequence."
            htaLog.WriteLine(Now & " || Scanstate Failed with error " & usmtError & ". Halting Flush and Fill sequence.")
            Exit Sub
        End If
    End If

    htaLog.WriteLine(Now & " || loadStateCheckBox.Checked = " & iLoadStateCheckBox)

    If iLoadStateCheckBox Then
        htaLog.WriteLine(Now & " || loadStateCheckBox is checked. Run usmtLoadstate_TS routine")
        usmtLoadstate_TS
    End IF

    htaLog.WriteLine(Now & " || Run ButtonFinishClick routine")
    
    ButtonFinishClick

    htaLog.WriteLine(Now & " ***** End Sub runFlushFill *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: ButtonExitClick
'**********************************************************************************************************************
Sub ButtonExitClick
    htaLog.WriteLine(Now & " ***** Begin Sub ButtonExitClick *****")

    window.close
    
    htaLog.WriteLine(Now & " ***** End Sub ButtonExitClick *****")
End Sub
'**********************************************************************************************************************




Sub runFlushFill2(vars)

    'Remove beginning and trailing {} brackets
    strVars = Mid(vars, 2, Len(vars)-2)
    'Create dictionary object and remove al of the quotes from everything
    Set dict = CreateObject("Scripting.Dictionary")
    strVars = Replace(strVars, """", "")

    'Split the string into name: value pairs
    For Each pair In Split(strVars, ",")
        arr = Split(pair, ":", 2)
        If UBound(arr) = 1 Then dict(arr(0)) = arr(1)
    Next

    colKeys = dict.Keys
    For Each strKey in colKeys
    MsgBox strKey & " = " & dict.Item(strKey)
    Next



End Sub
