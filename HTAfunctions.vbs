On Error Resume Next
Set ProgressUI = CreateObject("Microsoft.SMS.TsProgressUI") 
ProgressUI.CloseProgressDialog 

'Set objects and declare global variables
'Set env = CreateObject("Microsoft.SMS.TSEnvironment")
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
        MBAM.Checked = "True"
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

'************************************ USMT Scanstate subroutine ************************************
Sub usmtScanstate(buttonClicked)
	Dim getUser, WshShell, strCurrentDir, destDrive, scanStateDiv, returnCode
    Set scanStateDiv = document.getElementById("general-output")
    Set WshShell = CreateObject("WScript.Shell")
    strCurrentDir = WshShell.currentDirectory
	getUser = usmtUsername.Value
	destDrive = usmtDrive.Value

    scanStateDiv.innerHTML = "USMT Command that will execute: <br><br>" & strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c <br> /offline:" & strCurrentDir & "\USMT\offline.xml <br> /i:" & strCurrentDir & "\USMT\migdocs.xml <br> /i:" & strCurrentDir & "\USMT\migapp.xml <br> /i:" & strCurrentDir & "\USMT\oopexcludes.xml <br> /progress:" & strCurrentDir & "\prog.log <br> /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log <br> /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5, 1, True"

    If buttonClicked = "true" AND getUser <> "" AND destDrive <> "" Then
        returnCode = WshShell.run(strCurrentDir & "\USMT\scanstate.exe "&destDrive&":\USMT\"&getUser&" /c /offline:" & strCurrentDir & "\USMT\offline.xml /i:" & strCurrentDir & "\USMT\migdocs.xml /i:" & strCurrentDir & "\USMT\migapp.xml /i:" & strCurrentDir & "\USMT\oopexcludes.xml /progress:" & strCurrentDir & "\prog.log /L:"&destDrive&":\USMT\"&getUser&"\scanstate.log /listfiles:"&destDrive&":\USMT\"&getUser&"\filesCopied.log /V:5", 1, True)

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

'************************************ Execute DISM script ************************************
Sub usmtLoadstate
    env("envUsmtLoadstate") = "true"
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
        usmtScanstate "true"
    End If

    If loadStateCheckBox.Checked Then
        usmtLoadstate
    End IF
    'MsgBox "Finish"
    ButtonFinishClick
End Sub

'************************************ Exit HTA subroutine ************************************
Sub ButtonExitClick
	window.close
End Sub