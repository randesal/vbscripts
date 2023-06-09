Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "\\Backend_NAS\Nasserver_PE\Adrian\SCRIPTS\vbox\virtualbox.bat" & Chr(34), 0
Set WshShell = Nothing