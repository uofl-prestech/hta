'************************************ DISM Capture Image  ************************************
Dim dismShell, strName, strDestPath, strSourcePath
Set dismShell = CreateObject("WScript.Shell")
strSourcePath = env("dismSourceDrive")
strDestPath = env("dismDestDrive")
strName = env("dismUsername")

dismShell.run "cmd.exe /c X:\windows\system32\DISM.exe /Capture-Image /ImageFile:"&strDestPath&":\"&strName&".wim /CaptureDir:"&strSourcePath&":\ /Name:"&CHR(34) & strName &CHR(34) &" /ScratchDir:"&strDestPath&":\",1 ,True

Set fso = CreateObject("Scripting.FileSystemObject")
fileName = "X:\windows\logs\dism\dism.log"
Set myFile = fso.OpenTextFile(fileName,1)

Do While myFile.AtEndOfStream <> True
    textLine = myFile.ReadLine
    strRead = strRead & textLine & "<br>"
Loop
myFile.Close

'wscript.echo "Capture Finished!      " & strRead