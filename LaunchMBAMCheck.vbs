If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute "cscript.exe", WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

Dim cmdShell
Set cmdShell = CreateObject("WScript.Shell")
cmdShell.Run "powershell.exe -noprofile -executionpolicy bypass -file ./MBAM_ReportStatus.ps1", 1, true