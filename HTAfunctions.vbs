On Error Resume Next
Set ProgressUI = CreateObject("Microsoft.SMS.TsProgressUI") 
ProgressUI.CloseProgressDialog 

'Set objects and declare global variables
Set env = CreateObject("Microsoft.SMS.TSEnvironment")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'************************************ Create Log File ************************************
CONST ForAppending = 8
Dim logShell, strLogDir, objFSO, htaLog
Set logShell = CreateObject("WScript.Shell")
strLogDir = logShell.currentDirectory
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set htaLog = objFSO.OpenTextFile(strLogDir & "\HTALOG.txt", ForAppending, True)

htaLog.WriteLine(vbCrLf & "====================== Begin Logging at " & Now & " ======================")

'************************************ Ping test subroutines ************************************
Dim iTimerID
Dim pingShell, pingShellExec
Sub pingTest
    htaLog.WriteLine(Now & " ***** Begin Sub pingTest *****")

    Dim comspec, strObj
    dim pingTestDiv: set pingTestDiv = document.getElementById("general-output")
    pingTestDiv.innerHTML = "<p class='cmdHeading'>Network connectivity test: </p>"
    Set pingShell = CreateObject("WScript.Shell")
    comspec = pingShell.ExpandEnvironmentStrings("%comspec%")

    htaLog.WriteLine(Now & " || Executing command: cmd /c ping.exe www.google.com")

    Set pingShellExec = pingShell.Exec(comspec & " /c ping.exe www.google.com")
    iTimerID = window.setInterval("vbscript:writePing()", 10)
    
    htaLog.WriteLine(Now & " ***** End Sub pingTest *****")
End Sub

Sub writePing
    dim pingTestDiv: set pingTestDiv = document.getElementById("general-output")
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

'************************************ List local drives ************************************
Sub listDrives
    htaLog.WriteLine(Now & " ***** Begin Sub listDrives *****")
    Dim strComputer, objWMIService, colItems, drivesHashTable, admin
	Dim landingPageDiv: Set landingPageDiv = document.getElementById("page-landing")
	landingPageDiv.innerHTML = "<h2 class='cmdHeading'>Drive List: </h2>"
    strComputer = "."
    
    On Error Resume Next
    CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    If Err.number = 0 Then 
        admin = true
        htaLog.WriteLine(Now & " || User is running script as admin")
    Else
        admin = false
        htaLog.WriteLine(Now & " || Error. Must run as Administrator to check encryption status")
    End If
    Err.Clear

    If admin = true Then
        htaLog.WriteLine(Now & " || Executing command: GetObject(""winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2\Security\MicrosoftVolumeEncryption"")")
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2\Security\MicrosoftVolumeEncryption")

        htaLog.WriteLine(Now & " || Executing command: objWMIService.ExecQuery(""Select * from Win32_EncryptableVolume"",,48"")")
        Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_EncryptableVolume",,48)
        Set drivesHashTable = CreateObject("scripting.dictionary")

        htaLog.WriteLine(Now & " || Win32_EncryptableVolume instance")
        htaLog.WriteLine(Now & " || 0=Protection OFF, 1= Protection ON, 2=Protection Unknown")
        For Each objItem in colItems 
            htaLog.Write(Now & " || Drive Letter: " & objItem.DriveLetter)
            htaLog.Write(" || ProtectionStatus: " & objItem.ProtectionStatus)
            If(objItem.ProtectionStatus = 1) OR (objItem.ProtectionStatus = 2) Then
                htaLog.Write(" || Encrypted" & vbCrLf)
                drivesHashTable.Add objItem.DriveLetter, "Encrypted"
            Else
                htaLog.Write(" || Not Encrypted" & vbCrLf)
                drivesHashTable.Add objItem.DriveLetter, "Not Encrypted"
            End If
        Next
    Else
        htaLog.WriteLine(Now & " || Skipping check for Encryption Status")
    End If

    htaLog.WriteLine(Now & " || Executing command: GetObject(""winmgmts:\\"" & strComputer & ""\root\cimv2"")")

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    
    htaLog.WriteLine(Now & " || Executing command: objWMIService.ExecQuery(""Select * from Win32_Volume"")")
    
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Volume")
	
	For Each objItem In colItems
	    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Caption: " & objItem.Caption & "<br>"
        landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Drive Letter: " & objItem.DriveLetter & " | "
        htaLog.Write(Now & " || ""Drive Letter: " & objItem.DriveLetter & " | ")
        landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Capacity: " & ConvertSize(objItem.Capacity)
        htaLog.Write("Capacity: " & ConvertSize(objItem.Capacity))
        If admin = true Then
            landingPageDiv.innerHTML = landingPageDiv.innerHTML & " | " & drivesHashTable(objItem.DriveLetter) & "<br>"
            htaLog.WriteLine("Encryption Status: " & drivesHashTable(objItem.DriveLetter))
        Else
            landingPageDiv.innerHTML = landingPageDiv.innerHTML & "<br>"
            htaLog.WriteLine("")
        End If
	    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Drive Type: " & objItem.DriveType & "<br>"
	    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "File System: " & objItem.FileSystem & " | "
	    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Label: " & objItem.Label & "<br><br>"
	Next

    htaLog.WriteLine(Now & " ***** End Sub listDrives *****")

End Sub

'************************************ Open new command prompt ************************************
Sub cmdPrompt
    htaLog.WriteLine(Now & " ***** Begin Sub cmdPrompt *****")
	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: cmd /k")
    cmdShell.Run "cmd /k"
    htaLog.WriteLine(Now & " ***** End Sub cmdPrompt *****")
End Sub

'************************************ Open powershell prompt ************************************
Sub psPrompt
    htaLog.WriteLine(Now & " ***** Begin Sub psPrompt *****")
	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: Powershell.exe -noprofile -noexit -executionpolicy bypass")

    cmdShell.Run "Powershell.exe -noprofile -noexit -executionpolicy bypass"

    htaLog.WriteLine(Now & " ***** End Sub psPrompt *****")

End Sub

'************************************ Open cmtrace64 log viewer ************************************
Sub logViewer
    htaLog.WriteLine(Now & " ***** Begin Sub logViewer *****")

	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: .\tools\cmtrace64.exe")
    cmdShell.Run ".\tools\cmtrace64.exe"
    
    htaLog.WriteLine(Now & " ***** End Sub logViewer *****")

End Sub

'************************************ Open Notepad ++ ************************************
Sub notepadPP
    htaLog.WriteLine(Now & " ***** Begin Sub notepadPP *****")

	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")

    htaLog.WriteLine(Now & " || Executing command: .\tools\npp\notepadpp.exe")
    cmdShell.Run ".\tools\npp\notepadpp.exe"
    
    htaLog.WriteLine(Now & " ***** End Sub notepadPP *****")

End Sub

'************************************ Open Explorer++ file manager ************************************
Sub explorer
    htaLog.WriteLine(Now & " ***** Begin Sub explorer *****")

	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")
    cmdShell.Run ".\tools\explorerpp.exe"
    
    htaLog.WriteLine(Now & " ***** End Sub explorer *****")

End Sub

'************************************ Convert Bytes to KB, MB, GB, TB subroutine************************************
Function ConvertSize(Size)
    'htaLog.WriteLine(Now & " ***** Begin Sub ConvertSize *****")

	suffix = " Bytes" 
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

    'htaLog.WriteLine(Now & " ***** End Sub ConvertSize *****")

End Function

'************************************ Execute OSD subroutine ************************************
Sub ButtonFinishClick
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

    For each v in env.GetVariables 
        htaLog.WriteLine(v & " = " & env(v)) 
    Next 
    'MsgBox "End Finish Subroutine"

    window.close

    htaLog.WriteLine(Now & " ***** End Sub ButtonFinishClick *****")

End Sub

'************************************ DISM Capture Image subroutine ************************************
Sub dismCapture
    htaLog.WriteLine(Now & " ***** Begin Sub dismCapture *****")

	Dim dismShell, strName, destPath, sourcePath, returnCode
	Dim dismDiv: Set dismDiv = document.getElementById("general-output")
    strSourcePath = document.getElementById("input-windows-drive").Value
    strDestPath = document.getElementById("input-external-drive").Value
    strName = document.getElementById("input-primary-username").Value
    Set dismShell = CreateObject("WScript.Shell")
    
    htaLog.WriteLine(Now & " || strSourcePath = " & strSourcePath)
    htaLog.WriteLine(Now & " || strDestPath = " & strDestPath)
    htaLog.WriteLine(Now & " || strName = " & strName)

    dismDiv.innerHTML = "Running Command: X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log"
    
    htaLog.Writeline(Now & " || returnCode = dismShell.run (""cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log"", 1, True)")

	returnCode = dismShell.run ("cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log", 1, True)
	
    dismDiv.innerHTML = "Capture Finished! <br><br> Return Code: " & returnCode

    htaLog.WriteLine(Now & " || Capture finished!")
    htaLog.WriteLine(Now & " || Return Code: " & returnCode)

    htaLog.WriteLine(Now & " ***** End Sub dismCapture *****")

End Sub

'************************************ Execute DISM script ************************************
Sub runDISM_TS
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

'************************************ Enumerate users ************************************
Sub enumUsers
    htaLog.WriteLine(Now & " ***** Begin Sub enumUsers *****")

    Const HKLM = &H80000002
    Dim htmlString, strComputer, strHivePath, strKeyPath, strSubKeyPath, profilePath, userName, selectLength, strSourcePath
    'Dim landingPageDiv: Set landingPageDiv = document.getElementById("page-landing")
    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "<h2 class='cmdHeading'>User List: </h2>"
    Dim userNameDiv: Set userNameDiv = document.getElementById("div-select-users")
    set objshell = CreateObject("Wscript.shell")
    strComputer = "."
    'strSourcePath = windowsDrive.value
    strHivePath = "C:\Windows\System32\Config\SOFTWARE"
    strKeyPath = "TempSoftware\Microsoft\Windows NT\CurrentVersion\ProfileList\"

    htaLog.WriteLine(Now & " || strComputer = "".""")
    htaLog.WriteLine(Now & " || strHivePath = " & strHivePath)
    htaLog.WriteLine(Now & " || strKeyPath = ""TempSoftware\Microsoft\Windows NT\CurrentVersion\ProfileList""")
    
    htaLog.WriteLine(Now & " || Executing command: ""cmd /c reg.exe load HKLM\TempSoftware " & strHivePath & ", 0, true")

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
        userNameDiv.innerHTML = userNameDiv.innerHTML & "Could not load HKLM\TempSoftware<br><br>"
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
            selectLength = selectLength + 1
        End If
    Next
    userNameDiv.innerHTML = userNameDiv.innerHTML & "<span>Users: &nbsp&nbsp </span><br>"
    userNameDiv.innerHTML = userNameDiv.innerHTML & "<select id=""input-usmt-usernames"" name=""usmtUsernameList"" size="& selectLength &" multiple>" & htmlString & "</select>"
    
    htaLog.WriteLine(Now & " || Executing command: objShell.Run(""cmd /c reg.exe unload HKLM\TempSoftware"", 0, true)")
    objshell.Run "%comspec% /c reg.exe unload HKLM\TempSoftware", 0, true

    htaLog.WriteLine(Now & " ***** End Sub enumUsers *****")

End Sub

'************************************ USMT Scanstate subroutine ************************************
Sub usmtScanstate(buttonClicked)
    htaLog.WriteLine(Now & " ***** Begin Sub usmtScanstate(buttonClicked) *****")

    Dim getUser, WshShell, strCurrentDir, destDrive, scanStateDiv, returnCode, userArray, userArraySize, userIncludeString
    userArray = Array()
    Set scanStateDiv = document.getElementById("general-output")
    Set WshShell = CreateObject("WScript.Shell")
    strCurrentDir = WshShell.currentDirectory

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

    htaLog.WriteLine(Now & " || usmtUsername.Value = " & getUser)
    htaLog.WriteLine(Now & " || usmtDrive.Value = " & destDrive)

    scanStateDiv.innerHTML = "USMT Command that will execute: <br><br>" & strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c <br> /offline:" & strCurrentDir & "\USMT\offline.xml <br> /i:" & strCurrentDir & "\USMT\migdocs.xml <br> /i:" & strCurrentDir & "\USMT\migapp.xml <br> /i:" & strCurrentDir & "\USMT\oopexcludes.xml <br> /progress:" & strCurrentDir & "\prog.log <br> /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log <br> /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 <br> /ue:* " & userIncludeString & ", 1, True"

    htaLog.WriteLine(Now & " || Execute Scanstate if buttonClicked = true, getUser is not blank, and destDrive is not blank")
    htaLog.WriteLine(Now & " || buttonClicked = " & buttonClicked & ", getUser = " & getUser & ", destDrive = " & destDrive)
    htaLog.WriteLine(Now & " || Executing command: WshShell.Run (""cmd /c "&strCurrentDir&"\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c /o /offline:USMT\offline.xml /i:USMT\migdocs.xml /i:USMT\migapp.xml /i:USMT\oopexcludes.xml /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 /ue:* "&userIncludeString&", 1, True)")

    If buttonClicked = "true" AND getUser <> "" AND destDrive <> "" Then
        returnCode = WshShell.Run ("cmd /c " & strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c /o /offline:USMT\offline.xml /i:USMT\migdocs.xml /i:USMT\migapp.xml /i:USMT\oopexcludes.xml /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 /ue:* " & userIncludeString, 1, True)

        If returnCode = 0 Then
            htaLog.WriteLine(Now & " || Scanstate Complete!")
            scanStateDiv.innerHTML = "Scanstate Complete! <br> Log files can be found in "&destDrive&":\USMT\"&getUser&"\"
        Else
            htaLog.WriteLine(Now & " || Scanstate Failed!")
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
    End If

    htaLog.WriteLine(Now & " ***** End Sub usmtScanstate(buttonClicked) *****")

End Sub

'************************************ Execute Loadstate ************************************
Sub usmtLoadstate
    htaLog.WriteLine(Now & " ***** Begin Sub usmtLoadstate *****")

'Dim GetScriptPath : GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)

    Dim ReturnCode, getUser, sourceDrive, strCurrentDir
    Dim objShell : Set objShell = CreateObject("WScript.Shell")
    strCurrentDir = objShell.currentDirectory
    getUser = document.getElementById("input-primary-username").Value
    sourceDrive = document.getElementById("input-external-drive").Value

    htaLog.WriteLine(Now & " || strCurrentDir = " & strCurrentDir)
    htaLog.WriteLine(Now & " || getUser = " & getuser)
    htaLog.WriteLine(Now & " || sourceDrive = " & sourceDrive)
    htaLog.WriteLine(Now & " || Executing command: objShell.Run (""cmd /k " & strCurrentDir & "\USMT\loadstate.exe /c "&sourceDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&sourceDrive&":\USMT\"&getUser&"\loadstate.log"", 1, True)")

    ReturnCode = objShell.Run ("cmd /k " & strCurrentDir & "\USMT\loadstate.exe /c "&sourceDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&sourceDrive&":\USMT\"&getUser&"\loadstate.log", 1, True)
    
    htaLog.WriteLine(Now & " || Return Code: " & ReturnCode)

    htaLog.WriteLine(Now & " ***** End Sub usmtLoadstate *****")

End Sub

'************************************ Set up Loadstate ************************************
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

'************************************ Execute Flush and Fill script ************************************
Sub runFlushFill
    htaLog.WriteLine(Now & " ***** Begin Sub runFlushFill *****")

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
        dismCapture
    End If

    htaLog.WriteLine(Now & " || scanStateCheckBox.Checked = " & iScanStateCheckBox)

    If iScanStateCheckBox Then
        htaLog.WriteLine(Now & " || scanStateCheckBox is checked. Run usmtScanstate routine")
        usmtScanstate "true"
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

'************************************ Exit HTA subroutine ************************************
Sub ButtonExitClick
    htaLog.WriteLine(Now & " ***** Begin Sub ButtonExitClick *****")

    window.close
    
    htaLog.WriteLine(Now & " ***** End Sub ButtonExitClick *****")
End Sub