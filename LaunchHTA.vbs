If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute "cscript.exe", WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

Set WshShell = CreateObject("Wscript.Shell")
Dim GetScriptPath : GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)
Dim GetScriptDrive : GetScriptDrive = Left(WScript.ScriptFullName, 1)

WshShell.CurrentDirectory = GetScriptDrive & ":"
ReturnCode = WshShell.Run (GetScriptDrive & ":\LaunchMedia.cmd")

'Set fso = CreateObject("Scripting.FileSystemObject")
'If fso.FileExists(GetScriptPath & "\Prestech.hta") Then
  'If running from the DP200077 directory
'  ReturnCode = WshShell.Run (GetScriptPath & "\mshta.EXE " & GetScriptPath & "\Prestech.hta", 1, True)
'Else
  'If running from the root of the flash drive
  'GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1) & "\SMS\PKG\DP200077"
  'ReturnCode = WshShell.Run (GetScriptPath & "\mshta.EXE " & GetScriptPath & "\Prestech.hta", 1, True)
'End If

