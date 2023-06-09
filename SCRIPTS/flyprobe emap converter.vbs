 Dim fs, f, f1, fc, s
spath = inputbox("Enter File path")
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(spath)
    Set fc = f.Files
    For Each f1 In fc
        fname1 = f1.Name
        Set f = fs.OpenTextFile(f1.Path, 1, 0)
        stext = f.readall
        For x = 1 To 4
        stext = Replace(stext, "   ", "  ")
        Next
	stext = Replace(stext, "DIE",  vbcrlf & "DIE")
        stext = Replace(stext, "  ", vbTab)
        fpath = f1.Path
        Set a = fs.CreateTextFile(fpath, True)
        a.Write stext
        a.Close
    Next