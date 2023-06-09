Set WShell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")

Ln = UCase(InputBox("Input lot number", "DKM Search", lastLot))
If Ln <> "" Then
sValid = "\\192.168.1.75\file server\Yield Validation"

Dim f, f1, fc, s
    Set f = fs.GetFolder(sValid)
    Set fc = f.SubFolders
    For Each f1 In fc
        If fs.fileexists(f1.Path & "\" & Ln & ".xls") = True Then
            a = MsgBox("Open TEST emap", vbYesNo + vbSystemModal)
            If a <> vbNo Then
            WShell.Run "explorer.exe " & f1.Path & "\" & Ln & ".xls"
            End If
        End If
    Next
End If