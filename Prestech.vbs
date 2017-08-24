Option Explicit
Dim ReturnCode
Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim GetScriptPath : GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)
 '**********
'* MAIN *
'**********
 ReturnCode = objShell.Run ("cmd /c " & GetScriptPath & "\ServiceUI.exe -process:TsProgressUI.exe %SYSTEMROOT%\SYSTEM32\mshta.EXE " & GetScriptPath & "\Prestech.hta", 1, True)	
 Wscript.Quit(ReturnCode)
 '**************
'* END MAIN *
'**************
' Hides current Progress UI to bring the HTA to the front
'Set ProgressUI = CreateObject("Microsoft.SMS.TsProgressUI")
'ProgressUI.CloseProgressDialog
' Create a WshShell object
'set sh = CreateObject("Wscript.Shell")
' Call the Run method, and pass your command to it (eg. "mshta.exe MyHTA.hta").
' The last parameter ensures that the VBscript does not proceed / terminate until the mshta process is closed.
'call sh.Run("LandingPage.hta", 1, True)
'call sh.Run("ServiceUI.exe -session:1 %WINDIR%\system32\mshta.exe LandingPage.hta", 1, True)
'call sh.Run("LandingPage.hta", 1, True)

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
	strCurDir = cmdShell.CurrentDirectory
	cmdShell.Run strCurDir & "\cmtrace64.exe"
End Sub