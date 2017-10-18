'**********************************************************************************************************************
'						        Function: TPMCheck
'**********************************************************************************************************************
Sub TPMCheck
	htaLog.WriteLine(Now & " ***** Begin Sub TPMCheck *****")
	
	Dim bIsEnabled, bIsActivated, bIsOwned, objTPM, strStatusState, nRC, objWMITPM, strConnectionStr1, strStatusMessage, strTPMWarning, TPMBox, CompBox
	bIsEnabled = "False"
	bIsActivated = "False"
	bIsOwned = "False"
	strStatusMessage = "<p class=""tpm-error"">ERROR - TPM Not Found</p>"
	strTPMWarning = "<p class=""tpm-error""><span>!</span><span>!</span><span>!</span> CHECK TPM SETTINGS <span>!</span><span>!</span><span>!</span></p>"
	Dim outputDiv: set outputDiv = document.getElementById("tpm-check-output")
	outputDiv.innerHTML = ""
	document.getElementById("input-tpm-checkbox").Value = false

	On Error Resume Next

    CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    If Err.number = 0 Then 
        admin = true
        htaLog.WriteLine(Now & " || User is running script as admin")
    Else
		admin = false
		strStatusMessage = "<p class=""tpm-error"">Error: Must run as Administrator to check TPM status!</p>"
		htaLog.WriteLine(Now & " || Error: Must run as Administrator to check TPM status!")
		outputDiv.innerHTML = outputDiv.innerHTML & strStatusMessage
		Exit Sub
    End If
    Err.Clear
	'---------------------------------------------------------------------------------------- 
	'Connect to TPM WMI provider 
	'---------------------------------------------------------------------------------------- 
	strConnectionStr1 = "winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!root\cimv2\Security\MicrosoftTpm" 
	Err.clear

	htaLog.WriteLine(Now & " || Executing command: GetObject(""winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!root\cimv2\Security\MicrosoftTpm)""")

	Set objWMITPM = GetObject(strConnectionStr1) 
	If Err.Number <> 0 Then
		strStatusState = "Not Found"
	End If 
	Err.clear

	Set objTpm = objWMITPM.Get("Win32_Tpm=@") 
	If Err.Number <> 0 Then 
		strStatusState = "Not Found"
	End If 

	If strStatusState = "Not Found" Then
		htaLog.WriteLine(Now & " || Error: " & Err.Number & ". " & Err.Description & ". TPM Not Found")
		outputDiv.innerHTML = outputDiv.innerHTML & strStatusMessage
		outputDiv.innerHTML = outputDiv.innerHTML & strTPMWarning
		Err.Clear
	Else

		'----------------------------------------------------------------------------------------- 
		'Get TPM status data to determine if TPM is enabled, activated, and owned 
		'----------------------------------------------------------------------------------------- 
		nRC = objTpm.IsEnabled(bIsEnabled)

		If nRC <> 0 OR bIsEnabled = "False" Then
			strStatusState = "ERROR"
			htaLog.WriteLine(Now & " || TPM not enabled")
			bIsEnabled = "<span class=""tpmError"">False</span>"
		End If 

		nRC = objTpm.IsActivated(bIsActivated)
		If nRC <> 0 OR bIsActivated = "False" Then
			strStatusState = "ERROR"
			htaLog.WriteLine(Now & " || TPM not activated")
			bIsActivated = "<span class=""tpmError"">False</span>"
		End If 

		nRC = objTpm.IsOwned(bIsOwned)
		If nRC <> 0 OR bIsOwned = "False" Then
			strStatusState = "ERROR"
			htaLog.WriteLine(Now & " || TPM not owned")
			bIsOwned = "<span class=""tpmError"">False</span>"
		End If

		outputDiv.innerHTML = outputDiv.innerHTML & "TPM found in the following state: <br>"
		outputDiv.innerHTML = outputDiv.innerHTML & "Enabled - " & bIsEnabled & "<br>"
		outputDiv.innerHTML = outputDiv.innerHTML & "Activated - " & bIsActivated & "<br>"
		outputDiv.innerHTML = outputDiv.innerHTML & "Owned - " & bIsOwned & "<br>"

		If strStatusState = "ERROR" Then
			outputDiv.innerHTML = outputDiv.innerHTML & strTPMWarning
		Else
			htaLog.WriteLine(Now & " || TPM is enabled and activated")
			outputDiv.innerHTML = outputDiv.innerHTML & "<p id=""tpmGood"">TPM Enabled and Activated</p>"
			document.getElementById("input-tpm-checkbox").Value = true
		End If
	End If

	htaLog.WriteLine(Now & " ***** End Sub TPMCheck *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: BitlockerInfo
'**********************************************************************************************************************
Sub BitlockerInfo
	htaLog.WriteLine(Now & " ***** Begin Sub BitlockerInfo *****")

	Const ForReading = 1
	Const TriStateTrue = -1	'Open file as Unicode
	Dim cmdShell
	Dim driveDiv:set driveDiv = document.getElementById("bl-info-output")

	Set cmdShell = CreateObject("WScript.Shell")
	
	htaLog.WriteLine(Now & " || Executing command: cmdShell.Run ""powershell.exe -noprofile -windowstyle hidden -noninteractive -executionpolicy bypass -file ./BitlockerInfo.ps1"", 1, true")

	cmdShell.Run "powershell.exe -noprofile -windowstyle hidden -noninteractive -executionpolicy bypass -file ./BitlockerInfo.ps1", 1, true

	htaLog.WriteLine(Now & " || Creating bitlockerinfo.txt file at " & strLogDir)

	Set fso = CreateObject("Scripting.FileSystemObject")
	fileName = "bitlockerinfo.txt"
	Set myFile = fso.OpenTextFile(fileName, ForReading, false, TriStateTrue)
	Do While myFile.AtEndOfStream <> True
	   textLine = myFile.ReadLine
	   strRead = strRead & textLine & vbCrLf
	Loop
	myFile.Close
	htaLog.WriteLine(Now & " || Bitlocker info:")
	htaLog.WriteLine(strRead)
	driveDiv.innerHTML = "<H2>Bitlocker Info</H2><div>" & strRead & "</div><br>"

	htaLog.WriteLine(Now & " ***** End Sub BitlockerInfo *****")

End Sub
'**********************************************************************************************************************

'**********************************************************************************************************************
'						        Function: BitlockerUnlock
'**********************************************************************************************************************
Sub BitlockerUnlock
	htaLog.WriteLine(Now & " ***** Begin Sub BitlockerUnlock *****")
	Dim cmdShell, blOutput, drive, key
	drive = document.getElementById("input-bitlocker-drive").Value
	key = document.getElementById("input-bitlocker-key").Value

	htaLog.WriteLine(Now & " || document.getElementById(""input-bitlocker-drive"").Value = " & drive)
	htaLog.WriteLine(Now & " || document.getElementById(""input-bitlocker-key"").Value = " & key)

	Dim driveDiv:set driveDiv = document.getElementById("bl-unlock-output")
	Set cmdShell = CreateObject("Wscript.Shell")

	htaLog.WriteLine(Now & " || Executing command: cmdShell.Exec(""Powershell.exe -noprofile -windowstyle hidden -noninteractive -executionpolicy bypass -File ./BitlockerUnlock.ps1 -blKey " & Chr(34) & key & Chr(34) & " -drive " & Chr(34) & drive & Chr(34) & ")")

	Set outData = cmdShell.Exec("Powershell.exe -noprofile -windowstyle hidden -noninteractive -executionpolicy bypass -File ./BitlockerUnlock.ps1 -blKey " & Chr(34) & key & Chr(34) & " -drive " & Chr(34) & drive & Chr(34))
	outData.StdIn.Close

	blOutput = outData.StdOut.ReadAll
	htaLog.WriteLine(Now & " || " & blOutput)
	'driveDiv.innerHTML = "<div>" & blOutput & "</div><br>"

	htaLog.WriteLine(Now & " ***** End Sub BitlockerUnlock *****")

End Sub
'**********************************************************************************************************************