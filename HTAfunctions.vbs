On Error Resume Next
Set ProgressUI = CreateObject("Microsoft.SMS.TsProgressUI") 
ProgressUI.CloseProgressDialog 

'Set objects and declare global variables
Set env = CreateObject("Microsoft.SMS.TSEnvironment")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'************************************ Ping test subroutines ************************************
Dim iTimerID
Dim pingShell, pingShellExec
Sub pingTest
    Dim comspec, strObj
    dim pingTestDiv: set pingTestDiv = document.getElementById("general-output")
    pingTestDiv.innerHTML = "<p class='cmdHeading'>Network connectivity test: </p>"
    Set pingShell = CreateObject("WScript.Shell")
    comspec = pingShell.ExpandEnvironmentStrings("%comspec%")
    Set pingShellExec = pingShell.Exec(comspec & " /c ping.exe www.google.com")
	iTimerID = window.setInterval("vbscript:writePing()", 10)
End Sub

Sub writePing
	dim pingTestDiv: set pingTestDiv = document.getElementById("general-output")
	pingTestDiv.innerHTML = pingTestDiv.innerHTML & pingShellExec.StdOut.ReadLine() & "<br>"
	If pingShellExec.Status = 1 Then
		window.clearInterval(iTimerID)
		pingTestDiv.innerHTML = pingTestDiv.innerHTML & pingShellExec.StdOut.ReadAll() & "<br>"
	End If

End Sub

'************************************ List local drives ************************************
Sub listDrives
	Dim landingPageDiv: Set landingPageDiv = document.getElementById("landing-page-input")
	landingPageDiv.innerHTML = "<h2 class='cmdHeading'>Drive List: </h2>"
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Volume")
	
	For Each objItem In colItems
	    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Caption: " & objItem.Caption & "<br>"
	    landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Drive Letter: " & objItem.DriveLetter & " | "
	    landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Capacity: " & ConvertSize(objItem.Capacity) & "<br>"
	    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Drive Type: " & objItem.DriveType & "<br>"
	    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "File System: " & objItem.FileSystem & " | "
	    'landingPageDiv.innerHTML = landingPageDiv.innerHTML & "Label: " & objItem.Label & "<br><br>"
	Next

End Sub

'************************************ Open new command prompt ************************************
Sub cmdPrompt
	Dim cmdShell, cmdShellExec, comspec, strObj
    Set cmdShell = CreateObject("WScript.Shell")
	cmdShell.Run "cmd /k"
End Sub

'************************************ Open cmtrace64 log viewer ************************************
Sub logViewer
	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")
	cmdShell.Run ".\cmtrace64.exe"
End Sub

'************************************ Open Explorer++ file manager ************************************
Sub explorer
	Dim cmdShell
    Set cmdShell = CreateObject("WScript.Shell")
	cmdShell.Run ".\explorerpp.exe"
End Sub

'************************************ Convert Bytes to KB, MB, GB, TB subroutine************************************
Function ConvertSize(Size)
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
End Function

'************************************ Execute OSD subroutine ************************************
Sub ButtonFinishClick
    ' Set value of variable to true/false based on whether the checkbox is selected or not
    Set env = CreateObject("Microsoft.SMS.TSEnvironment")
    'MsgBox "Begin Finish Subroutine"
    If AdobeReaderDC.Checked OR OSDAdobeReaderDC.Checked OR fnfAdobeReaderDC.Checked Then
        strAdobeReaderDC = "true"
        else strAdobeReaderDC = "false"
    End If
    'MsgBox "AdobeReaderDC = " & strAdobeReaderDC
	If AdobeAcrobatProXI.Checked OR OSDAdobeAcrobatProXI.Checked OR fnfAdobeAcrobatProXI.Checked Then
		strAdobeAcrobatProXI = "true"
		else strAdobeAcrobatProXI="false"
	End If		
   'MsgBox "AdobeAcrobatProXI = " & strAdobeAcrobatProXI
    If Chrome.Checked OR OSDChrome.Checked OR fnfChrome.Checked Then
        strChrome = "true"
        else strChrome = "false"
    End If
    'MsgBox "Chrome = " & strChrome
	If FileMaker.Checked OR OSDFileMaker.Checked OR fnfFileMaker.Checked Then
        strFileMaker = "true"
        else strFileMaker = "false"
    End If
    'MsgBox "FileMaker = " & strFileMaker
    If Office2016.Checked OR OSDOffice2016.Checked OR fnfOffice2016.Checked Then
        strOffice2016 = "true"
        else strOffice2016 = "false"
    End If
    'MsgBox "Office2016 = " & strOffice2016
    If OSD.Checked OR fnfOSD.Checked Then
        strOSD = "true"
        MBAM.Checked = "true"
        strCompName = compName.value
        else strOSD = "false"
    End If
    'MsgBox "OSD = " & strOSD & vbCrLf & "MBAM = " & strMBAM & vbCrLf & "CompName = " & strCompName
    If MBAM.Checked Then
        strMBAM = "true"
        else strMBAM = "false"
    End If
	
    ' Set value of variables that will be used by the task sequence, then close the window and allow the task sequence to continue.
    'MsgBox "Set Environment Variables"   
    env("envAdobeReaderDC") = strAdobeReaderDC
    env("envAdobeAcrobatProXI") = strAdobeAcrobatProXI
    env("envChrome") = strChrome
    env("envFileMaker") = strFileMaker
    env("envOffice2016") = strOffice2016
    env("envOSD") = strOSD
    env("OSDComputerName") = strCompName
    env("envMBAM") = strMBAM
    'MsgBox "End Finish Subroutine"
    window.close
    'MsgBox "Window Closed"
End Sub

'************************************ DISM Capture Image subroutine ************************************
Sub dismCapture
	Dim dismShell, strName, destPath, sourcePath, returnCode
	Dim dismDiv: Set dismDiv = document.getElementById("general-output")
    strSourcePath = windowsDrive.value
    strDestPath = dismDrive.value
    strName = dismUsername.value
	Set dismShell = CreateObject("WScript.Shell")

	dismDiv.innerHTML = "Running Command: X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log"
	returnCode = dismShell.run ("cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log", 1, True)
	
    dismDiv.innerHTML = "Capture Finished! <br><br> Return Code: " & returnCode

End Sub

'************************************ Execute DISM script ************************************
Sub runDISM_TS
    env("envDismSelected") = "true"
    env("dismSourceDrive") = windowsDrive.value
    env("bitlockerKey") = blKey.value
    env("dismDestDrive") = dismDrive.value
    env("dismUsername") = dismUsername.value

    window.close

End Sub

'************************************ Enumerate users ************************************
Sub enumUsers
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
    
    objshell.Run "%comspec% /c reg.exe load HKLM\TempSoftware " & strHivePath, 0, true
    
    If Err <> 0 Then
        usmtUserNameDiv.innerHTML = usmtUserNameDiv.innerHTML & "Could not load HKLM\TempSoftware<br><br>"
        fnfUserNameDiv.innerHTML = fnfUserNameDiv.innerHTML & "Could not load HKLM\TempSoftware<br><br>"
        strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\ProfileList"
        Err.Clear
    End If

    Set objReg = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv")

    objReg.EnumKey HKLM, strKeyPath, arrSubKeys

    selectLength = 0
    For Each subkey In arrSubKeys
        strSubKeyPath = "HKEY_LOCAL_MACHINE\" & strKeyPath & "\" & subkey & "\ProfileImagePath"
        profilePath = objshell.regRead(strSubKeyPath)
        profilePath = Split(profilePath, "\")
        userName = profilePath(Ubound(profilePath))

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

    objshell.Run "%comspec% /c reg.exe unload HKLM\TempSoftware", 0, true

End Sub

'************************************ USMT Scanstate subroutine ************************************
Sub usmtScanstate(buttonClicked)
    Dim getUser, WshShell, strCurrentDir, destDrive, scanStateDiv, returnCode, userArray, userArraySize, userIncludeString
    userArray = Array()
    Set scanStateDiv = document.getElementById("general-output")
    Set WshShell = CreateObject("WScript.Shell")
    strCurrentDir = WshShell.currentDirectory

    For i = 0 to (usmtUsernameList.Options.Length - 1)
        If (usmtUsernameList.Options(i).Selected) Then
            ReDim Preserve userArray(UBound(userArray) + 1)
            userArray(UBound(userArray)) = usmtUsernameList.Options(i).Value
            userIncludeString = userIncludeString & "/ui:" & usmtUsernameList.Options(i).Value & " "
        End If
    Next
    userArraySize = Ubound(userArray)

	getUser = usmtUsername.Value
	destDrive = usmtDrive.Value

    scanStateDiv.innerHTML = "USMT Command that will execute: <br><br>" & strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c <br> /offline:" & strCurrentDir & "\USMT\offline.xml <br> /i:" & strCurrentDir & "\USMT\migdocs.xml <br> /i:" & strCurrentDir & "\USMT\migapp.xml <br> /i:" & strCurrentDir & "\USMT\oopexcludes.xml <br> /progress:" & strCurrentDir & "\prog.log <br> /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log <br> /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 <br> /ue:* " & userIncludeString & ", 1, True"

    If buttonClicked = "true" AND getUser <> "" AND destDrive <> "" Then
        'returnCode = WshShell.run(strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c /offline:" & strCurrentDir & "\USMT\offline.xml /i:" & strCurrentDir & "\USMT\migdocs.xml /i:" & strCurrentDir & "\USMT\migapp.xml /i:" & strCurrentDir & "\USMT\oopexcludes.xml /progress:" & strCurrentDir & "\prog.log /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 /ue:* " & userIncludeString, 1, True)
        returnCode = WshShell.Run ("cmd /c " & strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c /o /offline:USMT\offline.xml /i:USMT\migdocs.xml /i:USMT\migapp.xml /i:USMT\oopexcludes.xml /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5 /ue:* " & userIncludeString, 1, True)

        If returnCode = 0 Then
            'WshShell.run "%comspec% /c cmtrace.exe " & strCurrentDir & "\prog.log"
            scanStateDiv.innerHTML = "Scanstate Complete! <br> Log files can be found in "&destDrive&":\USMT\"&getUser&"\"
        Else
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

            scanStateDiv.innerHTML = scanStateDiv.innerHTML & "<br><br> Scanstate Failed! <br><br>" & strRead
        End If
    End If

End Sub

'************************************ Execute Loadstate ************************************
Sub usmtLoadstate
    Dim ReturnCode, getUser, sourceDrive, strCurrentDir
    Dim objShell : Set objShell = CreateObject("WScript.Shell")
    strCurrentDir = objShell.currentDirectory
    getUser = usmtUsername.Value
    sourceDrive = usmtDrive.Value

    ReturnCode = objShell.Run ("cmd /k " & strCurrentDir & "\USMT\loadstate.exe /c "&sourceDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&sourceDrive&":\USMT\"&getUser&"\loadstate.log", 1, True)	
    Wscript.Quit(ReturnCode)

End Sub

'************************************ Set up Loadstate ************************************
Sub usmtLoadstate_TS
    env("envUsmtLoadstate") = "True"
    env("envUsmtSourceDrive") = windowsDrive.Value
    env("envUsmtDestDrive") = usmtDrive.Value
    env("envUsmtUsername") = usmtUsername.Value

    window.close

End Sub

'************************************ Execute Flush and Fill script ************************************
Sub runFlushFill
    windowsDrive.Value = fnfWindowsDrive.Value
    dismDrive.Value = fnfDrive.Value
    dismUsername.Value = fnfUsername.Value
    usmtUsername.Value = fnfUsername.Value
    usmtDrive.Value = fnfDrive.Value
    compName.Value = fnfCompName.Value
    'dismMsgResult = MsgBox ("Run DISM?",vbYesNo+vbInformation, "")
    'If dismMsgResult = 6 Then
    If dismCheckBox.Checked Then
        dismCapture
    End If
    'usmtMsgResult = MsgBox ("Run USMT?",vbYesNo+vbInformation, "")
    'If usmtMsgResult = 6 Then
    If scanStateCheckBox.Checked Then
        For i = 0 to (fnfUsernameList.Options.Length - 1)
            If (fnfUsernameList.Options(i).Selected) Then
                usmtUsernameList.Options(i).Selected = True
            End If
        Next
        usmtScanstate "true"
    End If

    If loadStateCheckBox.Checked Then
        MsgBox("Set Environment variables for USMT Loadstate")
        usmtLoadstate_TS
    End IF
    'MsgBox "Finish"
    ButtonFinishClick
End Sub

'************************************ Exit HTA subroutine ************************************
Sub ButtonExitClick
	window.close
End Sub