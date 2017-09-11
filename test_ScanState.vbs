Option Explicit
Dim ReturnCode, getUser
Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim GetScriptPath : GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)
getUser = "prestech"


'**********
'* MAIN *
'**********
 ReturnCode = objShell.Run ("cmd /k " & GetScriptPath & "\USMT\scanstate.exe D:\USMT\"&getUser&" /c /offline:USMT\offline.xml /i:USMT\migdocs.xml /i:USMT\migapp.xml /i:USMT\oopexcludes.xml /progress:USMT\prog.log /L:D:\USMT\"&getUser&"\scanstate.log /listfiles:D:\USMT\"&getUser&"\filesCopied.log /V:5 /ue:sccmpush /ue:defaultuser0 /ue:AD\sccmpush /ue:Administrator /ue:grawem01", 1, True)
 Wscript.Quit(ReturnCode)
