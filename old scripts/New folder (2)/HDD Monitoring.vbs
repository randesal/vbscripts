Dim sDrive(100), sDRV(11)
ON ERROR RESUME NEXT
sLog = "\\192.168.1.7\Nasserver_PE\Monitoring\AOI_HDD_Space.csv"
sLog2 = "\\192.168.1.76\Adrian\AOI_HDD_Space.csv"
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.fileexists(sLog) = False Then fs.CreateTextFile sLog
If fs.fileexists(sLog2) = False Then fs.CreateTextFile sLog2
Set f = fs.GetFile(sLog)
Set f2 = fs.GetFile(sLog2)
sDrive(1) = "\\192.168.1.11\Hscope"
sDrive(2) = "\\192.168.1.21\Hscope"
sDrive(3) = "\\192.168.1.31\Hscope"
sDrive(4) = "\\192.168.1.41\Hscope"
sDrive(5) = "\\192.168.1.51\Hscope"
sDrive(6) = "\\192.168.1.61\Hscope"
sDrive(7) = "\\192.168.1.71\Hscope"
sDrive(8) = "\\192.168.1.81\Hscope"
sDrive(9) = "\\192.168.1.91\Hscope"
sDrive(10) = "\\192.168.1.121\Hscope"
For x = 1 To 10
Set d = fs.GetDrive(sDrive(x))
    sDRV(x) =d.FreeSpace / 1024 / 1024/1024
Next
sDRVtext = NOW & Join(sDRV, ",") & vbCrLf
Set ts = f.OpenAsTextStream(8, -2)
Set ts2 = f2.OpenAsTextStream(8, -2)
ts.Write sDRVtext
ts.CLOSE
ts2.Write sDRVtext
ts2.CLOSE