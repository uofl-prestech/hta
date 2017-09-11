Const HKLM = &H80000002
set objshell = CreateObject("Wscript.shell")
strComputer = "."
strHivePath = "C:\Windows\System32\Config\SOFTWARE"
'strHivePath = "E:\SOFTWARE"
'strKeyPath = "TempSoftware\Microsoft\Windows NT\CurrentVersion\ProfileList"
strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\ProfileList"

'objshell.Run "%comspec% /k reg.exe load HKLM\TempSoftware " & strHivePath, 1, true
Set objReg = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv")

objReg.EnumKey HKLM, strKeyPath, arrSubKeys

For Each subkey In arrSubKeys
    strSubKeyPath = "HKEY_LOCAL_MACHINE\" & strKeyPath & "\" & subkey & "\ProfileImagePath"
    profilePath = objshell.regRead(strSubKeyPath)
    profilePath = Split(profilePath, "\")
    userName = profilePath(Ubound(profilePath))

    wscript.echo "User: " & userName & " ==> SID: " & subkey
Next
'objReg.SetStringValue HKEY_USERS,strKeyPath1,"Wallpaper","C:\WINDOWS\Web\Wallpaper\kiosk.bmp"


'objshell.Exec ("%comspec% /k reg.exe unload HKLM\TempSoftware")


