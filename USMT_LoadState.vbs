Option Explicit
Dim ReturnCode, getUser, usmtDrive
Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim strCurrentDir : strCurrentDir = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)
getUser = envPrimaryUsername
usmtDrive = envExternalDrive

 ReturnCode = objShell.Run ("cmd /k " & strCurrentDir & "\USMT\loadstate.exe /c "&sourceDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&sourceDrive&":\USMT\"&getUser&"\loadstate.log", 1, True)
 Wscript.Quit(ReturnCode)
