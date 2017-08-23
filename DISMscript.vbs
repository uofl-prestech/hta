'************************************ DISM Capture Image  ************************************
Dim dismShell, strName, strDestPath, strSourcePath, blShell, key
strSourcePath = env("dismSourceDrive")
strDestPath = env("dismDestDrive")
strName = env("dismUsername")
key = env("bitlockerKey")

'********** Unlock Bitlocker Drive  **********
Set blShell = CreateObject("Wscript.Shell")
Set outData = blShell.Exec("Powershell.exe -noprofile -windowstyle hidden -noninteractive -executionpolicy bypass -File BitlockerUnlock.ps1 -blKey " & Chr(34) & key & Chr(34) & " -drive " & Chr(34) & strSourcePath & Chr(34)) 


'********** Capture Image **********
Set dismShell = CreateObject("WScript.Shell")
returnCode = dismShell.run ("cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\ /LogPath:X:\dism.log", 1, True)
