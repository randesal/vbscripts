Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "virtualbox.bat" & Chr(34), 0
Set WshShell = Nothing