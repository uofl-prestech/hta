'Define Variables and Objects.
Set WshShell = CreateObject("Wscript.Shell")

'Note: Gpupdate command has to be run twice as the ECHO command can't answer more than one question. 

'Refresh the USER policies and also answer no to logoff if asked.
'wscript.echo "Updating user policy"
'Result = WshShell.Run("cmd /c echo n | gpupdate /target:user /force",1,true)

'Refresh the Computer policies and answer no to reboot. 
wscript.echo "Updating computer policy"
Result = WshShell.Run("cmd /c echo n | gpupdate /target:computer /force",1,true)

'Hand back the errorlevel
Wscript.Quit(Result)