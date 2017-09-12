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
	Dim landingPageDiv: Set landingPageDiv = document.getElementById("landing-page-input")
	landingPageDiv.innerHTML = "<h2 class='cmdHeading'>Drive List: </h2>"
    strComputer = "."
    
    On Error Resume Next
    htaLog.WriteLine(Now & " || Executing command: GetObject(""winmgmts:\\" & strComputer & "\root\CIMV2\Security\MicrosoftVolumeEncryption"")")
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2\Security\MicrosoftVolumeEncryption")
    If Err <> 0 Then
        htaLog.WriteLine(Now & " || Error. Must run as Administrator to check encryption status")
        admin = false
        Err.Clear
    Else
        htaLog.WriteLine(Now & " || Executing command: objWMIService.ExecQuery(""Select * from Win32_EncryptableVolume"",,48"")")
        Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_EncryptableVolume",,48)
        Set drivesHashTable = CreateObject("scripting.dictionary")

        htaLog.WriteLine(Now & " || Win32_EncryptableVolume instance")
        htaLog.WriteLine(Now & " || 0=Protection OFF, 1= Protection ON, 2=Protection Unknown")
        admin = true
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

    htaLog.WriteLine(Now & " || Executing command: .\cmtrace64.exe")
    cmdShell.Run ".\cmtrace64.exe"
    
    htaLog.WriteLine(Now & " ***** End Sub logViewer *****")

End Sub

'************************************ Open Explorer++ file manager ************************************
Sub explorer
    htaLog.WriteLine(Now & " ***** Begin Sub explorer *****")

	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")
    cmdShell.Run ".\explorerpp.exe"
    
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

    ' Set value of variable to true/false based on whether the checkbox is selected or not
    'Set env = CreateObject("Microsoft.SMS.TSEnvironment")
    'MsgBox "Begin Finish Subroutine"
    htaLog.WriteLine(Now & " || Adobe Reader DC checked?")
    htaLog.Write(Now & " || AdobeReaderDC.Checked = " & AdobeReaderDC.Checked)
    htaLog.Write(" || OSDAdobeReaderDC.Checked = " & OSDAdobeReaderDC.Checked)
    htaLog.Write(" || fnfAdobeReaderDC.Checked = " & fnfAdobeReaderDC.Checked & vbCrLf)

    If AdobeReaderDC.Checked OR OSDAdobeReaderDC.Checked OR fnfAdobeReaderDC.Checked Then
        strAdobeReaderDC = "true"
        else strAdobeReaderDC = "false"
    End If
    'MsgBox "AdobeReaderDC = " & strAdobeReaderDC

    htaLog.WriteLine(Now & " || Adobe Acrobat Pro XI checked?")
    htaLog.Write(Now & " || AdobeAcrobatProXI.Checked = " & AdobeAcrobatProXI.Checked)
    htaLog.Write(" || OSDAdobeAcrobatProXI.Checked = " & OSDAdobeAcrobatProXI.Checked)
    htaLog.Write(" || fnfAdobeAcrobatProXI.Checked = " & fnfAdobeAcrobatProXI.Checked & vbCrLf)
	If AdobeAcrobatProXI.Checked OR OSDAdobeAcrobatProXI.Checked OR fnfAdobeAcrobatProXI.Checked Then
		strAdobeAcrobatProXI = "true"
		else strAdobeAcrobatProXI="false"
	End If		
   'MsgBox "AdobeAcrobatProXI = " & strAdobeAcrobatProXI

    htaLog.WriteLine(Now & " || Chrome checked?")
    htaLog.Write(Now & " || Chrome.Checked = " & Chrome.Checked)
    htaLog.Write(" || OSDChrome.Checked = " & OSDChrome.Checked)
    htaLog.Write(" || fnfChrome.Checked = " & fnfChrome.Checked & vbCrLf)
    If Chrome.Checked OR OSDChrome.Checked OR fnfChrome.Checked Then
        strChrome = "true"
        else strChrome = "false"
    End If
    'MsgBox "Chrome = " & strChrome

    htaLog.WriteLine(Now & " || FileMaker checked?")
    htaLog.Write(Now & " || FileMaker.Checked = " & FileMaker.Checked)
    htaLog.Write(" || OSDFileMaker.Checked = " & OSDFileMaker.Checked)
    htaLog.Write(" || fnfFileMaker.Checked = " & fnfFileMaker.Checked & vbCrLf)
	If FileMaker.Checked OR OSDFileMaker.Checked OR fnfFileMaker.Checked Then
        strFileMaker = "true"
        else strFileMaker = "false"
    End If
    'MsgBox "FileMaker = " & strFileMaker

    htaLog.WriteLine(Now & " || Office2016 checked?")
    htaLog.Write(Now & " || Office2016.Checked = " & Office2016.Checked)
    htaLog.Write(" || OSDOffice2016.Checked = " & OSDOffice2016.Checked)
    htaLog.Write(" || fnfOffice2016.Checked = " & fnfOffice2016.Checked & vbCrLf)
    If Office2016.Checked OR OSDOffice2016.Checked OR fnfOffice2016.Checked Then
        strOffice2016 = "true"
        else strOffice2016 = "false"
    End If
    'MsgBox "Office2016 = " & strOffice2016

    htaLog.WriteLine(Now & " || OSD checked?")
    htaLog.Write(Now & " || OSD.Checked = " & OSD.Checked)
    htaLog.Write(" || fnfOSD.Checked = " & fnfOSD.Checked & vbCrLf)
    If OSD.Checked OR fnfOSD.Checked Then
        strOSD = "true"
        MBAM.Checked = "true"
        strCompName = compName.value
        else strOSD = "false"
    End If
    'MsgBox "OSD = " & strOSD & vbCrLf & "MBAM = " & strMBAM & vbCrLf & "CompName = " & strCompName

    htaLog.WriteLine(Now & " || MBAM checked?")
    htaLog.WriteLine(Now & " || MBAM.Checked = " & MBAM.Checked)
    If MBAM.Checked Then
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
    strSourcePath = windowsDrive.value
    strDestPath = dismDrive.value
    strName = dismUsername.value
    Set dismShell = CreateObject("WScript.Shell")
    
    htaLog.WriteLine(Now & " || strSourcePath = " & strSourcePath)
    htaLog.WriteLine(Now & " || strDestPath = " & strDestPath)
    htaLog.WriteLine(Now & " || strName = " & strName)

    dismDiv.innerHTML = "Running Command: X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log"
    
    htaLog.Writeline(Now & " || returnCode = dismShell.run (""cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:""&strDestPath&"":\""&strName&"".wim /CaptureDir:""&strSourcePath&"":\ /Name:""&CHR(34) & strName &CHR(34) &"" /ScratchDir:""&strDestPath&"":\ /LogPath:X:\dism.log"", 1, True)")

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
    env("dismSourceDrive") = windowsDrive.value
    env("bitlockerKey") = blKey.value
    env("dismDestDrive") = dismDrive.value
    env("dismUsername") = dismUsername.value

    htaLog.WriteLine(Now & " || env(""envDismSelected"") = " & env("envDismSelected"))
    htaLog.WriteLine(Now & " || env(""dismSourceDrive"") = " & env("dismSourceDrive"))
    htaLog.WriteLine(Now & " || env(""bitlockerKey"") = " & env("bitlockerKey"))
    htaLog.WriteLine(Now & " || env(""dismDestDrive"") = " & env("dismDestDrive"))
    htaLog.WriteLine(Now & " || env(""dismUserName"") = " & env("dismUserName"))

    htaLog.WriteLine(Now & " ***** End Sub runDISM_TS *****")

    window.close

End Sub

'************************************ Enumerate users ************************************
Sub enumUsers
    htaLog.WriteLine(Now & " ***** Begin Sub enumUsers *****")

    Const HKLM = &H80000002
    Dim htmlString, strComputer, strHivePath, strKeyPath, strSubKeyPath, profilePath, userName, selectLength
    'Dim landingPageDiv: Set landingPageDiv = document.getElementById("landing-page-input")
    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "<h2 class='cmdHeading'>User List: </h2>"
    Dim usmtUserNameDiv: Set usmtUserNameDiv = document.getElementById("usmt-step4")
    Dim fnfUserNameDiv: Set fnfUserNameDiv = document.getElementById("fnf-step5")
    set objshell = CreateObject("Wscript.shell")
    strComputer = "."
    strHivePath = "C:\Windows\System32\Config\SOFTWARE"
    strKeyPath = "TempSoftware\Microsoft\Windows NT\CurrentVersion\ProfileList"

    htaLog.WriteLine(Now & " || strComputer = "".""")
    htaLog.WriteLine(Now & " || strHivePath = " & strHivePath)
    htaLog.WriteLine(Now & " || strKeyPath = ""TempSoftware\Microsoft\Windows NT\CurrentVersion\ProfileList""")
    
    htaLog.WriteLine(Now & " || Executing command: ""cmd /c reg.exe load HKLM\TempSoftware " & strHivePath & ", 0, true")

    On Error Resume Next
    Err.Clear

    objshell.Run "%comspec% /c reg.exe load HKLM\TempSoftware " & strHivePath, 0, true
    objshell.regRead("HKEY_LOCAL_MACHINE\" & strKeyPath)

    If Err <> 0 Then
        htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
        htaLog.WriteLine(Now & " || Error Description: " & Err.Description)
        htaLog.WriteLine(Now & " || Could not load HKLM\TempSoftware")
        usmtUserNameDiv.innerHTML = usmtUserNameDiv.innerHTML & "Could not load HKLM\TempSoftware<br><br>"
        fnfUserNameDiv.innerHTML = fnfUserNameDiv.innerHTML & "Could not load HKLM\TempSoftware<br><br>"
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
            'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "User: <b>" & userName & "</b> ==> SID: <b id=" & userName & ">" & subkey & "</b><br>"
            'usmtUserNameDiv.innerHTML = usmtUserNameDiv.innerHTML & "<option value=""all"">I love all sports!</option>"
            htmlString = htmlString & "<option value=" & subkey & ">" & userName & "</option>"
            selectLength = selectLength + 1
        End If
    Next
    usmtUserNameDiv.innerHTML = usmtUserNameDiv.innerHTML & "<span>Users: &nbsp&nbsp </span><br>"
    usmtUserNameDiv.innerHTML = usmtUserNameDiv.innerHTML & "<select id=""usmt-username-input"" name=""usmtUsernameList"" size="& selectLength &" multiple>" & htmlString & "</select>"
    fnfUserNameDiv.innerHTML = fnfUserNameDiv.innerHTML & "<span>Users: &nbsp&nbsp </span><br>"
    fnfUserNameDiv.innerHTML = fnfUserNameDiv.innerHTML & "<select id=""fnf-username-input"" name=""fnfUsernameList"" size="& selectLength &" multiple>" & htmlString & "</select>"

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

	getUser = usmtUsername.Value
	destDrive = usmtDrive.Value

    htaLog.WriteLine(Now & " || usmtUsername.Value = " & getUser)
    htaLog.WriteLine(Now & " || usmtDrive.Value = " & destDrive)

    scanStateDiv.innerHTML = "USMT Command that will execute: <br><br>" & strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c <br> /offline:" & strCurrentDir & "\USMT\offline.xml <br> /i:" & strCurrentDir & "\USMT\migdocs.xml <br> /i:" & strCurrentDir & "\USMT\migapp.xml <br> /i:" & strCurrentDir & "\USMT\oopexcludes.xml <br> /progress:" & strCurrentDir & "\prog.log <br> /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log <br> /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 <br> /ue:* " & userIncludeString & ", 1, True"

    htaLog.WriteLine(Now & " || Execute Scanstate if buttonClicked = true, getUser is not blank, and destDrive is not blank")
    htaLog.WriteLine(Now & " || buttonClicked = " & buttonClicked & ", getUser = " & getUser & ", destDrive = " & destDrive)
    htaLog.WriteLine(Now & " || Executing command: WshShell.Run (""cmd /c "&strCurrentDir&"\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c /o /offline:USMT\offline.xml /i:USMT\migdocs.xml /i:USMT\migapp.xml /i:USMT\oopexcludes.xml /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 /ue:* "&userIncludeString&", 1, True)")

    If buttonClicked = "true" AND getUser <> "" AND destDrive <> "" Then
        'returnCode = WshShell.run(strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c /offline:" & strCurrentDir & "\USMT\offline.xml /i:" & strCurrentDir & "\USMT\migdocs.xml /i:" & strCurrentDir & "\USMT\migapp.xml /i:" & strCurrentDir & "\USMT\oopexcludes.xml /progress:" & strCurrentDir & "\prog.log /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 /ue:* " & userIncludeString, 1, True)
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

    Dim ReturnCode, getUser, sourceDrive, strCurrentDir
    Dim objShell : Set objShell = CreateObject("WScript.Shell")
    strCurrentDir = objShell.currentDirectory
    getUser = usmtUsername.Value
    sourceDrive = usmtDrive.Value

    htaLog.WriteLine(Now & " || strCurrentDir = " & strCurrentDir)
    htaLog.WriteLine(Now & " || getUser = " & getuser)
    htaLog.WriteLine(Now & " || sourceDrive = " & sourceDrive)
    htaLog.WriteLine(Now & " || Executing command: objShell.Run (""cmd /k " & strCurrentDir & "\USMT\loadstate.exe /c "&sourceDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&sourceDrive&":\USMT\"&getUser&"\loadstate.log"", 1, True)")

    ReturnCode = objShell.Run ("cmd /k " & strCurrentDir & "\USMT\loadstate.exe /c "&sourceDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&sourceDrive&":\USMT\"&getUser&"\loadstate.log", 1, True)
    
    htaLog.WriteLine(Now & " || Return Code: " & ReturnCode)

    Wscript.Quit(ReturnCode)

    htaLog.WriteLine(Now & " ***** End Sub usmtLoadstate *****")

End Sub

'************************************ Set up Loadstate ************************************
Sub usmtLoadstate_TS
    htaLog.WriteLine(Now & " ***** Begin Sub usmtLoadstate_TS *****")

    env("envUsmtLoadstate") = "True"
    env("envUsmtSourceDrive") = windowsDrive.Value
    env("envUsmtDestDrive") = usmtDrive.Value
    env("envUsmtUsername") = usmtUsername.Value

    htaLog.WriteLine(Now & " || env(""envUsmtLoadstate"") = " & env("envUsmtLoadstate"))
    htaLog.WriteLine(Now & " || env(""envUsmtSourceDrive"") = " & env("envUsmtSourceDrive"))
    htaLog.WriteLine(Now & " || env(""envUsmtDestDrive"") = " & env("envUsmtDestDrive"))
    htaLog.WriteLine(Now & " || env(""envUsmtUsername"") = " & env("envUsmtUsername"))

    htaLog.WriteLine(Now & " ***** End Sub usmtLoadstate_TS *****")

End Sub

'************************************ Execute Flush and Fill script ************************************
Sub runFlushFill
    htaLog.WriteLine(Now & " ***** Begin Sub runFlushFill *****")

    windowsDrive.Value = fnfWindowsDrive.Value
    dismDrive.Value = fnfDrive.Value
    dismUsername.Value = fnfUsername.Value
    usmtUsername.Value = fnfUsername.Value
    usmtDrive.Value = fnfDrive.Value
    compName.Value = fnfCompName.Value

    htaLog.WriteLine(Now & " || windowsDrive.Value = " & windowsDrive.Value)
    htaLog.WriteLine(Now & " || dismDrive.Value = " & dismDrive.Value)
    htaLog.WriteLine(Now & " || dismUsername.Value = " & dismUsername.Value)
    htaLog.WriteLine(Now & " || usmtUsername.Value = " & usmtUsername.Value)
    htaLog.WriteLine(Now & " || usmtDrive.Value = " & usmtDrive.Value)
    htaLog.WriteLine(Now & " || compName.Value = " & compName.Value)
    htaLog.WriteLine(Now & " || dismCheckBox.Checked = " & dismCheckBox.Checked)

    If dismCheckBox.Checked Then
        htaLog.WriteLine(Now & " || dismCheckBox is checked. Run dismCapture routine")
        dismCapture
    End If

    htaLog.WriteLine(Now & " || scanStateCheckBox.Checked = " & scanStateCheckBox.Checked)

    If scanStateCheckBox.Checked Then
        htaLog.WriteLine(Now & " || scanStateCheckBox is checked. Run usmtScanstate routine")
        For i = 0 to (fnfUsernameList.Options.Length - 1)
            If (fnfUsernameList.Options(i).Selected) Then
                usmtUsernameList.Options(i).Selected = True
            End If
        Next
        usmtScanstate "true"
    End If

    htaLog.WriteLine(Now & " || loadStateCheckBox.Checked = " & loadStateCheckBox.Checked)

    If loadStateCheckBox.Checked Then
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