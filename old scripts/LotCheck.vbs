Set WShell = CreateObject("WScript.Shell")
'sLog = "\\192.168.1.75\file server\LotCheck.log"
Set fs = CreateObject("Scripting.FileSystemObject")
'if fs.fileexists(slog)=false then  fs.CreateTextFile slog
'Set f = fs.GetFile(sLog)
' Set ts = f.OpenAsTextStream(1, -2)
'Do While ts.AtEndOfStream <> True
'    s = ts.ReadLine
'Loop
'    ts.Close
'if s = "" then s = ","
's = split(s,",")
'LastLot = s(1)

'Ln = ucase(InputBox("Input lot number","Lot Number",lastLot))
Ln = ucase(InputBox("Input lot number","Lot Number"))
if ln <> "" then

TEST = "\\192.168.1.75\file server\Backend-test\test e-map\_EMap Data\" & Ln
AOITOP = "\\192.168.1.75\file server\BACKEND AOI DATA\csv4\TongHsing-TOP\" & Ln
CDCTOP = "\\192.168.1.75\file server\BACKEND CDC-II RAW DATA\TongHsing-TOP\" & Ln
AOIBOT = "\\192.168.1.75\file server\BACKEND AOI DATA\csv4\TongHsing-BOT\" & Ln
CDCBOT = "\\192.168.1.75\file server\BACKEND CDC-II RAW DATA\TongHsing-BOT\" & Ln
Merge = "\\192.168.1.75\file server\MERGE EMAP\" & Ln
FTPTest = "\\192.168.1.75\file server\FTP Test DATA\" & LN
EMapFinal = "\\192.168.1.75\file server\EMAP FINAL\" & Ln

If fs.folderexists(TEST) = True Then
a = MsgBox("Open TEST emap", vbyesno+vbsystemmodal)
If a <> vbNo Then
WShell.Run "explorer.exe " & TEST
End If
End If


If fs.folderexists(AOITOP) = True Then
a = MsgBox("Open AOI top emap", vbyesno+vbsystemmodal)
If a <> vbNo Then
WShell.Run "explorer.exe " & AOITOP
End If
End If


If fs.folderexists(AOIBOT) = True Then
a = MsgBox("Open AOI bot emap", vbyesno+vbsystemmodal)
If a <> vbNo Then
WShell.Run "explorer.exe " & AOIBOT
End If
End If


If fs.folderexists(CDCTOP) = True Then
a = MsgBox("Open CDC top emap", vbyesno+vbsystemmodal)
If a <> vbNo Then
WShell.Run "explorer.exe " & CDCTOP
End If
End If

If fs.folderexists(CDCBOT) = True Then
a = MsgBox("Open CDC bot emap", vbyesno+vbsystemmodal)
If a <> vbNo Then
WShell.Run "explorer.exe " & CDCBOT
End If
End If

If fs.folderexists(Merge) = True Then
a = MsgBox("Open Merge emap", vbyesno+vbsystemmodal)
If a <> vbNo Then
WShell.Run "explorer.exe " & Merge
End If
End If

If fs.folderexists(FTPTEST) = True Then
a = MsgBox("Open FTP TEST emap", vbyesno+vbsystemmodal)
If a <> vbNo Then
WShell.Run "explorer.exe " & FTPTEST
End If
End If

If fs.folderexists(EmapFinal) = True Then
a = MsgBox("Open EMap Final emap", vbyesno+vbsystemmodal)
If a <> vbNo Then
WShell.Run "explorer.exe " & EmapFinal
End If
End If

If fs.folderexists(TEST) = False Then t = "No TEST data" & vbcrlf
If fs.folderexists(Merge) = False Then u = "No MERGE data" & vbcrlf
If fs.folderexists(AOIBOT) = False Then v = "No AOI BOT data" & vbcrlf
If fs.folderexists(AOITOP) = False Then w = "No AOI TOP data" & vbcrlf
If fs.folderexists(FTPTEST) = False Then x = "No FTP TEST data" & vbcrlf
If fs.folderexists(CDCTOP) = False Then y = "No CDC TOP data" & vbcrlf
If fs.folderexists(CDCBOT) = False Then z = "No CDC BOT data" & vbcrlf
If fs.folderexists(EmapFinal) = False Then s = "No Emap Final BOT data" & vbcrlf
msg = t & u & v & w & x & y & z & s
if msg <> "" then
msgbox  t & u & v & w & x & y & z & s
end if
end if
