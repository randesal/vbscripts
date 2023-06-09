Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "del_mapmerger.bat" & Chr(34), 0
Set WshShell = Nothing