'************************************ Create Log File ************************************
CONST ForAppending = 8
Dim logShell, strLogDir, objFSO, htaLog
Set logShell = CreateObject("WScript.Shell")
strLogDir = logShell.currentDirectory
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set htaLog = objFSO.OpenTextFile(strLogDir & "\HTALOG.log", ForAppending, True)	
Set env = CreateObject("Microsoft.SMS.TSEnvironment")
On Error Resume Next

Dim ReturnCode, getUser, usmtDrive, strCurrentDir
Dim objShell : Set objShell = CreateObject("WScript.Shell")
strCurrentDir = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\", -1, 1) - 1)
getUser = env("envPrimaryUsername")
usmtDrive = env("envExternalDrive")

htaLog.WriteLine(Now & " || strCurrentDir = " & strCurrentDir)
htaLog.WriteLine(Now & " || getUser = " & getuser)
htaLog.WriteLine(Now & " || usmtDrive = " & usmtDrive)
htaLog.WriteLine(Now & " || Executing command: objShell.Run (""cmd /k " & strCurrentDir & "\USMT\loadstate.exe /c "&usmtDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&usmtDrive&":\USMT\"&getUser&"\loadstate.log"", 1, True)")

ReturnCode = objShell.Run("cmd /c " & strCurrentDir & "\USMT\loadstate.exe /c "&usmtDrive&":\USMT\" & getUser & " /i:USMT\migapp.xml /i:USMT\migdocs.xml /v:13 /l:"&usmtDrive&":\USMT\"&getUser&"\loadstate.log", 1, True)

If Err <> 0 Then
	htaLog.WriteLine(Now & " || Loadstate Error: " & ReturnCode)
	Wscript.Quit(ReturnCode)
End If

'Load user settings into registry from NTUSER.DAT file originally located at C:\Users\<username>
'This file should be copied to <usmtDrive>:\USMT\<getUser>\NTUSER.DAT during scanstate and loaded to the registry from there.
'Registry Key: Computer\HKEY_USERS\<SID>\Control Panel\Desktop\WallPaper

Const HKLM = &H80000002
Dim objReg, strComputer, strKeyPath, strValueName, strUserSID, strWallpaperFile, profilePath, strSubKeyPath, userName, strHivePath
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\"
strComputer = "."
strHivePath = "C:\Users\"& getUser & "\NTUSER.DAT"
htaLog.WriteLine(Now & " || strHivePath = " & strHivePath)
strWallpaperRegPath = "TempUser\Control Panel\Desktop"
strValueName = "WallPaper"
strWallpaperFile = env("envWallpaperFile")
htaLog.WriteLine(Now & " || strWallpaperFile = " & strWallpaperFile)

htaLog.WriteLine(Now & " || Executing command: ""reg.exe load HKLM\TempUser " & strHivePath & ", 0, true")
Err.Clear
objshell.Run "reg.exe load HKLM\TempUser " & strHivePath, 0, true

If Err <> 0 Then
	htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
	htaLog.WriteLine(Now & " || Error Description: " & Err.Description)
	Err.Clear
End If

htaLog.WriteLine(Now & " || Executing command: GetObject(""winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv"")")
Set objReg = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv")

htaLog.WriteLine(Now & " || Executing command: objReg.SetStringValue HKLM, " & strWallpaperRegPath & ", " & strValueName & ", " & strWallpaperFile)
objReg.SetStringValue HKLM, strWallpaperRegPath, strValueName, strWallpaperFile

If Err = 0 Then
	objReg.GetStringValue HKLM, strWallpaperRegPath, strValueName, strWallpaperFile
	htaLog.WriteLine(Now & " ||  " & strWallpaperRegPath & "\" & strValueName & " contains: " & strWallpaperFile)
Else 
	htaLog.WriteLine(Now & " ||  Error in creating key and REG_SZ value = " & Err.Number)
End If

htaLog.WriteLine(Now & " || Executing command: objShell.Run(""reg unload HKLM\TempUser"", 0, true)")
objshell.Run "reg unload HKLM\TempUser", 0, true

htaLog.WriteLine(Now & " ***** End USMT_LoadState.vbs *****")
htaLog.Close