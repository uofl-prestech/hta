If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute "cscript.exe", WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

Set WshShell = CreateObject("Wscript.Shell")
Dim GetScriptPath : GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)
ReturnCode = WshShell.Run (GetScriptPath & "\mshta.EXE " & GetScriptPath & "\Prestech.hta", 1, True)	