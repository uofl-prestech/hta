Option Explicit
Dim ReturnCode
Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim GetScriptPath : GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)

 ReturnCode = objShell.Run ("cmd /c " & GetScriptPath & "\ServiceUI.exe -process:TsProgressUI.exe %SYSTEMROOT%\SYSTEM32\mshta.EXE " & GetScriptPath & "\Prestech.hta", 1, True)	
 Wscript.Quit(ReturnCode)