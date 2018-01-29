#************************************ Create Log File ************************************
[string]$currentDir = Get-Location

$logFile = "$currentDir\HTALOG.log"
$stamp = (Get-Date).toString("MM/dd/yyyy HH:mm:ss")

Function LogWrite
{
    Param ([string]$logString)
    Add-content $logFile -value $logString
}


LogWrite "***** Begin Script USMT_Loadstate.ps *****"

$getUser = "prestech" #env("envPrimaryUsername")
$usmtDrive = "J:" #FindUSMTDrive(getUser)

LogWrite "$($stamp) || currentDir = $currentDir"
LogWrite "$($stamp) || getUser = $getuser"
LogWrite "$($stamp) || usmtDrive = $usmtDrive"
LogWrite "Executing Loadstate command"

$storePath = "$usmtDrive\USMT\$getUser"
$loadStateArgs = "/c /i:USMT\migapp.xml /i:USMT\migdocs.xml" #/l:$usmtDrive\USMT\getUser\loadstate.log /v:13

LogWrite "$($stamp) || Before Invoke-Command"

Invoke-Command -ScriptBlock {
    $currentDir = $args[0]
    $getUser = $args[1]
    $storePath = $args[2]
    $loadStateArgs = $args[3]

    LogWrite "$($stamp) || currentDir = $currentDir"
    LogWrite "$($stamp) || getUser = $getUser"
    LogWrite "$($stamp) || storePath = $storePath"
    LogWrite "$($stamp) || loadStateArgs = $loadStateArgs"

    #Check if Loadstate.exe exists
    if (Test-Path "$currentDir\USMT\loadstate.exe") {$loadState = "$currentDir\USMT\loadstate.exe"}
    else {LogWrite "Unable to locate LoadState in $currentDir\USMT\ . Cannot restore user state." -ForegroundColor Red -BackgroundColor Black; return}

    LogWrite "$($stamp) || $loadState `"$storePath\$getUser`" $($loadStateArgs.Split(" ")) /l:$usmtDrive\USMT\getUser\loadstate.log /v:13"
    &$loadState "$storePath\$getUser" $loadStateArgs.Split(" ") /l:$usmtDrive\USMT\getUser\loadstate.log /v:13
} -ArgumentList $currentDir, $getUser, $storePath, $loadStateArgs

LogWrite "$($stamp) || After Invoke-Command"

<#Set env = CreateObject("Microsoft.SMS.TSEnvironment")

'Load user settings into registry from NTUSER.DAT file originally located at C:\Users\<username>
'This file should be copied to <usmtDrive>:\USMT\<getUser>\NTUSER.DAT during scanstate and loaded to the registry from there.
'Registry Key: Computer\HKEY_USERS\<SID>\Control Panel\Desktop\WallPaper

Const HKLM = &H80000002
Dim objReg, strComputer, strKeyPath, strValueName, strUserSID, strWallpaperFile, profilePath, strSubKeyPath, strHivePath, strWpPath, strFullPathTest
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\"
strComputer = "."
strHivePath = "C:\Users\"& getUser & "\NTUSER.DAT"
htaLog.WriteLine(Now & " || strHivePath = " & strHivePath)
strWallpaperRegPath = "TempUser\Control Panel\Desktop"
strValueName = "WallPaper"

strWallpaperFile = env("envWallpaperFile")
'env("envWallpaperFile") should be set during FnF Scanstate, but not if we are running OSD + Loadstate
Dim fsoWP, arrExtensions
Set fsoWP = CreateObject("Scripting.FileSystemObject")
fsoWP.GetFile(env("envWallpaperFile"))
If Err = 0 Then		'env was set during FnF Scanstate
	htaLog.WriteLine(Now & " || TS env check worked")
Else				'env wasn't set or this is an OSD + Loadstate operation
	htaLog.WriteLine(Now & " || TS env check failed")
	htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
	htaLog.WriteLine(Now & " || Error Description: " & Err.Description & vbCrLf)
	Err.Clear

	strWpPath = usmtDrive & ":\USMT\" & getUser & "\" & getUser
	'Check various image file types to see if the wallpaper file exists in the USMT folder on the external USB drive
	arrExtensions = Array(".jpg", ".jpeg", ".png", ".gif", ".bmp")
	For Each extension In arrExtensions
		strFullPathTest = strWpPath & extension
		fsoWP.GetFile(strFullPathTest)
		If Err <> 0 Then
			htaLog.WriteLine(Now & " || Error. " & strFullPathTest & " not found.")
			htaLog.WriteLine(Now & " || Error Number: " & Err.Number)
			htaLog.WriteLine(Now & " || Error Description: " & Err.Description & vbCrLf)
			Err.Clear
		Else
			htaLog.WriteLine(Now & " || Yay! " & strFullPathTest & " found!")
			strWallpaperFile = strFullPathTest
			Exit For
		End IF
	Next
End If

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
	objReg.GetStringValue HKLM, strWallpaperRegPath, strValueName, wpVerified
	htaLog.WriteLine(Now & " || HKLM\" & strWallpaperRegPath & "\" & strValueName & " contains: " & wpVerified)
Else 
	htaLog.WriteLine(Now & " || Error in creating key and REG_SZ value = " & Err.Number)
End If

htaLog.WriteLine(Now & " || Executing command: objShell.Run(""reg.exe unload HKLM\TempUser"", 0, true)")
objshell.Run "reg.exe unload HKLM\TempUser", 0, true

htaLog.WriteLine(Now & " ***** End USMT_LoadState.vbs *****")
htaLog.Close



'**********************************************************************************************************************
'						        Function: FindUSMTDrive()
'**********************************************************************************************************************
Function FindUSMTDrive(getUser)
	On Error Resume Next
	Dim username : username = getUser
	Dim alphabet : alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	htaLog.WriteLine(Now & " || Attempting to find USMT Drive")
	For i=1 To Len(alphabet)
    	Dim letter : letter = Mid(alphabet,i,1)
		htaLog.WriteLine(Now & " || Checking for folder - " & letter & ":\USMT\" & username & "\")
		If fso.FolderExists(letter & ":\USMT\" & username & "\") Then
			FindUSMTDrive = letter
			Exit Function
		End If
	Next
End Function
'**********************************************************************************************************************
#>