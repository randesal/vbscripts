do
on error resume next
dim drv(11)
drv(1) = "\\192.168.1.11\Hscope\log\r_st\"
drv(2) = "\\192.168.1.21\Hscope\log\r_st\"
drv(3) = "\\192.168.1.31\Hscope\log\r_st\"
drv(4) = "\\192.168.1.41\Hscope\log\r_st\"
drv(5) = "\\192.168.1.51\Hscope\log\r_st\"
drv(6) = "\\192.168.1.61\Hscope\log\r_st\"
drv(7) = "\\192.168.1.71\Hscope\log\r_st\"
drv(8) = "\\192.168.1.81\Hscope\log\r_st\"
drv(9) = "\\192.168.1.91\Hscope\log\r_st\"
drv(10) = "\\192.168.1.141\Hscope\log\r_st\"
drv(11) = "\\192.168.1.151\Hscope\log\r_st\"

for x = 1 to 11
for y = 1 to 2
select case y
case 1
drvPath = drv(x) & "Tonghsing-TOP"
case 2
drvPath = drv(x) & "Tonghsing-BOT"
end select
Set fs = CreateObject("Scripting.FileSystemObject")
Set drve = fs.GetDrive(fs.GetDriveName(drvPath))
sP = FormatNumber(drve.FreeSpace/1024/1024/1024, 0)
if sP < 2000 then 
call delfiles(drvPath)
'msgbox drvPath
end if
next
next
loop

sub delfiles(spath)

  Set fs = CreateObject("Scripting.FileSystemObject")
  sLog = "C:\Autodelete\delete.log"
  If fs.fileexists(sLog) = False Then fs.CreateTextFile sLog
  Set h = fs.GetFile(sLog)
  Set f = fs.GetFolder(spath)
  Set fc = f.subfolders
  
    For Each f1 In fc
    Set g = fs.GetFolder(f1.Path)
    Set gc = g.subfolders
        For Each g1 In gc
        d = g1.Name
        e = g1.Path
        slen = Len(g1.Name) <> 8
        sdate = g1.datelastmodified
        s = sdate < Now - 5
            If s And slen Then
            fs.deletefolder e
	'msgbox e
            Set ts = h.OpenAsTextStream(8, -2)
            ts.Write e & "," & sdate & "," & Now & vbCrLf
            ts.Close
            End If
        Next
Next
end sub