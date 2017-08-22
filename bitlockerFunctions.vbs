Sub TPMCheck
	Dim bIsEnabled, bIsActivated, bIsOwned, objTPM, strStatusState, nRC, objWMITPM, strConnectionStr1, strStatusMessage, strTPMWarning
	bIsEnabled = "False"
	bIsActivated = "False"
	bIsOwned = "False"
	strStatusMessage = "<h2 class=""tpmStatus"">ERROR - TPM Not Found</h2>"
	strTPMWarning = "<h2 class=""tpmStatus""><span>!</span><span>!</span><span>!</span> CHECK TPM SETTINGS <span>!</span><span>!</span><span>!</span></h2>"
	Dim outputDiv: set outputDiv = document.getElementById("osdOutput")
	outputDiv.innerHTML = ""
	On Error Resume Next

	'---------------------------------------------------------------------------------------- 
	'Connect to TPM WMI provider 
	'---------------------------------------------------------------------------------------- 
	strConnectionStr1 = "winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!root\cimv2\Security\MicrosoftTpm" 
	Err.clear

	Set objWMITPM = GetObject(strConnectionStr1) 
	If Err.Number <> 0 Then
		strStatusState = "Not Found"
	End If 
	Err.clear

	Set objTpm = objWMITPM.Get("Win32_Tpm=@") 
	If Err.Number <> 0 Then 
		strStatusState = "Not Found"
	End If 
	Err.Clear

	If strStatusState = "Not Found" Then
		outputDiv.innerHTML = strStatusMessage
		outputDiv.innerHTML = outputDiv.innerHTML & strTPMWarning
	else
		'----------------------------------------------------------------------------------------- 
		'Get TPM status data to determine if TPM is enabled, activated, and owned 
		'----------------------------------------------------------------------------------------- 
		nRC = objTpm.IsEnabled(bIsEnabled)

		If nRC <> 0 OR bIsEnabled = "False" Then
			strStatusState = "ERROR"
			bIsEnabled = "<span class=""tpmError"">False</span>"
		End If 

		nRC = objTpm.IsActivated(bIsActivated)
		If nRC <> 0 OR bIsActivated = "False" Then
			strStatusState = "ERROR"
			bIsActivated = "<span class=""tpmError"">False</span>"
		End If 

		nRC = objTpm.IsOwned(bIsOwned)
		If nRC <> 0 OR bIsOwned = "False" Then
			strStatusState = "ERROR"
			bIsOwned = "<span class=""tpmError"">False</span>"
		End If

		outputDiv.innerHTML = "TPM found in the following state: <br>"
		outputDiv.innerHTML = outputDiv.innerHTML & "Enabled - " & bIsEnabled & "<br>"
		outputDiv.innerHTML = outputDiv.innerHTML & "Activated - " & bIsActivated & "<br>"
		outputDiv.innerHTML = outputDiv.innerHTML & "Owned - " & bIsOwned & "<br>"

		If strStatusState = "ERROR" Then
			outputDiv.innerHTML = outputDiv.innerHTML & strTPMWarning
		else
			outputDiv.innerHTML = outputDiv.innerHTML & "<h2>TPM Enabled and Activated</h2>"
		End If
	End If

End Sub

'************************************ Bitlocker Info subroutine ************************************
Sub BitlockerInfo
	Dim cmdShell, outData
	Dim driveDiv:set driveDiv = document.getElementById("dismOutput")
	
    Set cmdShell = CreateObject("WScript.Shell")
	Set outData = cmdShell.Exec("powershell.exe -noprofile -windowstyle hidden -noninteractive -executionpolicy bypass -file BitlockerInfo.ps1")
	outData.StdIn.Close
	driveDiv.innerHTML = "<div>" & outData.StdOut.ReadAll & "</div><br>"
End Sub

'************************************ Unlock Bitlocker Drive subroutine ************************************
Sub BitlockerUnlock
	Dim cmdShell, key, drive
	key = blKey.Value
	drive = windowsDrive.Value
	Dim driveDiv:set driveDiv = document.getElementById("dismOutput")
	
	Set cmdShell = CreateObject("Wscript.Shell")
	Set outData = cmdShell.Exec("Powershell.exe -noprofile -windowstyle hidden -noninteractive -executionpolicy bypass -File BitlockerUnlock.ps1 -blKey " & Chr(34) & key & Chr(34) & " -drive " & Chr(34) & drive & Chr(34))
	outData.StdIn.Close
	driveDiv.innerHTML = "<div>" & outData.StdOut.ReadAll & "</div><br>" 
End Sub