Set fs = CreateObject("Scripting.FileSystemObject")

Ln = ucase(InputBox("Input lot number","Lot Number"))
if Ln <> "" then end sub 

TEST = "\\192.168.1.75\file server\Backend-test\test e-map\_EMap Data\" & Ln
AOITOP = "\\192.168.1.75\file server\BACKEND AOI DATA\csv4\TongHsing-TOP\" & Ln

If fs.folderexists(AOITOP) = True Then
a = MsgBox("Open AOI top emap", vbyesno+vbsystemmodal)
If a <> vbNo Then
WShell.Run "explorer.exe " & AOITOP
sLot = now & "," & Ln & ",AOI TOP"& vbcrlf
    Set ts = f.OpenAsTextStream(8, -2)
    ts.Write slot
    ts.Close
End If
End If


If fs.folderexists(AOITOP) = False Then w = "No AOI TOP data" & vbcrlf

msg = "No AOI TOP data"
if msg <> "" then
msgbox "No AOI TOP data"
end if
end if
