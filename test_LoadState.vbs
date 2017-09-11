Option Explicit
Dim ReturnCode, getUser
Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim GetScriptPath : GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)
getUser = "prestech"


'**********
'* MAIN *
'**********
 ReturnCode = objShell.Run ("cmd /k " & GetScriptPath & "\USMT\loadstate.exe /c D:\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:USMT\loadstate.log", 1, True)	
 Wscript.Quit(ReturnCode)
