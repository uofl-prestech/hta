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
