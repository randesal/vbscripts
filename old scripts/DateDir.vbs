sMon =Month(date)
sDay =Day(date)
sYear =Year(date)
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
if len(Month(date)) = 1 then sMon = "0" & Month(date)
if len(Day(date)) = 1 then sDay = "0" & Day(date)
strFolder = scriptdir & "\" & sYear & "\" &  sMon '& "\" & sDay
Set objFSO = CreateObject("wscript.shell")
objFSO.run "cmd /c MkDir " & strFolder 