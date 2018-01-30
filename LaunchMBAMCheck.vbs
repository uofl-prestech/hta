If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute "cscript.exe", WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

Dim GetScriptPath : GetScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)

Dim cmdShell
Set cmdShell = CreateObject("WScript.Shell")
cmdShell.CurrentDirectory = GetScriptPath
cmdShell.Run "powershell.exe -noprofile -noexit -executionpolicy bypass -file ./MBAM_ReportStatus.ps1", 1, true