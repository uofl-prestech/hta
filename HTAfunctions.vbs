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
    dim pingTestDiv: set pingTestDiv = document.getElementById("toolsOutput")
    pingTestDiv.innerHTML = "<p class='cmdHeading'>Network connectivity test: </p>"
    Set pingShell = CreateObject("WScript.Shell")
    comspec = pingShell.ExpandEnvironmentStrings("%comspec%")
    Set pingShellExec = pingShell.Exec(comspec & " /c ping.exe www.google.com")
	iTimerID = window.setInterval("vbscript:writePing()", 10)
End Sub

Sub writePing
	dim pingTestDiv: set pingTestDiv = document.getElementById("toolsOutput")
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
	strCurDir = cmdShell.CurrentDirectory
	cmdShell.Run strCurDir & "\cmtrace64.exe"
End Sub

'************************************ DISM Capture Image subroutine ************************************
Sub dismCapture
	Dim dismShell, strName, destPath, sourcePath, returnCode
	dim dismDiv: set dismDiv = document.getElementById("dismOutput")
    strSourcePath = windowsDrive.value
    strDestPath = dismDrive.value
    strName = dismUsername.value
	Set dismShell = CreateObject("WScript.Shell")

	dismDiv.innerHTML = "Running Command: X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log"
	returnCode = dismShell.run ("cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log", 1, True)
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	fileName = "X:\dism.log"
	' Set myFile = fso.OpenTextFile(fileName, 1)
	' Do While myFile.AtEndOfStream <> True
	'    textLine = myFile.ReadLine
	'    strRead = strRead & textLine & "<br>"
	' Loop
	' myFile.Close
	
     dismDiv.innerHTML = "Capture Finished! <br><br> Return Code: " & returnCode

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

    If AdobeReaderDC.Checked OR OSDAdobeReaderDC.Checked Then
        strAdobeReaderDC = "true"
        else strAdobeReaderDC = "false"
    End If

	If AdobeAcrobatProXI.Checked OR OSDAdobeAcrobatProXI.Checked Then
		strAdobeAcrobatProXI = "true"
		else strAdobeAcrobatProXI="false"
	End If		
   
    If Chrome.Checked OR OSDChrome.Checked Then
        strChrome = "true"
        else strChrome = "false"
    End If

	If FileMaker.Checked OR OSDFileMaker.Checked Then
        strFileMaker = "true"
        else strFileMaker = "false"
    End If
      
    If Office2016.Checked OR OSDOffice2016.Checked Then
        strOffice2016 = "true"
        else strOffice2016 = "false"
    End If

    If OSD.Checked Then
        strOSD = "true"
        strMBAM = "true"
        strCompName = compName.value
        else strOSD = "false"
    End If

    If MBAM.Checked Then
        strMBAM = "true"
        else strMBAM = "false"
    End If
	
    ' Set value of variables that will be used by the task sequence, then close the window and allow the task sequence to continue.
        
    env("envAdobeReaderDC") = strAdobeReaderDC
    env("envAdobeAcrobatProXI") = strAdobeAcrobatProXI
    env("envChrome") = strChrome
    env("envFileMaker") = strFileMaker
    env("envOffice2016") = strOffice2016
    env("envOSD") = strOSD
    env("OSDComputerName") = strCompName
    env("envMBAM") = strMBAM
   
    window.close
End Sub

'************************************ Execute DISM script ************************************
Sub runDISM_TS
    Dim oShell
    Set oShell = CreateObject("WScript.Shell")

    env("envDismSelected") = "true"
    env("dismSourceDrive") = windowsDrive.value
    env("bitlockerKey") = blKey.value
    env("dismDestDrive") = dismDrive.value
    env("dismUsername") = dismUsername.value

    Dim dismShell, strName, strDestPath, strSourcePath
    Set dismShell = CreateObject("WScript.Shell")
    strSourcePath = windowsDrive.value
    strDestPath = dismDrive.value
    strName = dismUsername.value

    dismShell.run "cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log",1 ,True

    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = "X:\dism.log"
    Set myFile = fso.OpenTextFile(fileName,1)

    Do While myFile.AtEndOfStream <> True
        textLine = myFile.ReadLine
        strRead = strRead & textLine & "<br>"
    Loop
    myFile.Close

    wscript.echo "Capture Finished!      " & strRead

End Sub

'************************************ Exit HTA subroutine ************************************
Sub ButtonExitClick
	window.close
End Sub