Option Explicit
Dim ReturnCode, getUser, usmtDrive
Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim GetScriptPath : GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)
getUser = envUsmtUsername
usmtDrive = envUsmtDestDrive

'**********
'* MAIN *
'**********
 ReturnCode = objShell.Run ("cmd /c " & GetScriptPath & "\USMT\loadstate.exe D:\USMT\" & getUser & " /i:D:\USMT\migapp.xml /i:D:\USMT\migdocs.xml", 1, True)	
 Wscript.Quit(ReturnCode)
