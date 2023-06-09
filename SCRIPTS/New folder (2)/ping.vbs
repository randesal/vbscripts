Dim sDrive(100),sDRV(11)
on error resume next
sLog = "\\192.168.1.7\Nasserver_PE\Monitoring\AOI_connection.csv"
sLog2 = "\\192.168.1.76\Adrian\AOI_connection.csv"
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.fileexists(sLog) = False Then fs.CreateTextFile sLog
If fs.fileexists(sLog2) = False Then fs.CreateTextFile sLog2
Set f = fs.GetFile(sLog)
Set f2 = fs.GetFile(sLog2)
set a = CreateObject("Wscript.Shell")
sDrive(1) = "192.168.1.11"
sDrive(2) = "192.168.1.21"
sDrive(3) = "192.168.1.31"
sDrive(4) = "192.168.1.41"
sDrive(5) = "192.168.1.51"
sDrive(6) = "192.168.1.61"
sDrive(7) = "192.168.1.71"
sDrive(8) = "192.168.1.81"
sDrive(9) = "192.168.1.91"
sDrive(10) = "192.168.1.121"
for x = 1 to 10
pingIP = sDrive(X)
boolcode = a.Run("ping -n 1 -w 100 " & pingIP, 0, True)
select case boolcode
case 0
sDrv(X) = "G"
case 1
sDrv(X) = "NG"
end select
next
sDRVtext = NOW & Join(sDRV, ",") & vbCrLf
Set ts = f.OpenAsTextStream(8, -2)
Set ts2 = f2.OpenAsTextStream(8, -2)
ts.Write sDRVtext
ts.CLOSE
ts2.Write sDRVtext
ts2.CLOSE