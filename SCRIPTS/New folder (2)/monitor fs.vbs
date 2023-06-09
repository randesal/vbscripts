ON ERROR RESUME NEXT
sLog = "\\192.168.1.7\Nasserver_PE\Monitoring\Fileserver_Space.csv"
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.fileexists(sLog) = False Then fs.CreateTextFile sLog
Set f = fs.GetFile(sLog)
sDrive = "\\192.168.1.75/file server"
Set d = fs.GetDrive(sDrive)
sDRV =d.FreeSpace / 1024 / 1024/1024
sDRVtext = NOW & "," & sDRV & vbCrLf
Set ts = f.OpenAsTextStream(8, -2)
ts.Write sDRVtext
ts.CLOSE
