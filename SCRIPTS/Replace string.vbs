 Dim fs, f, f1, fc, s
spath = inputbox("Enter Folder path")
sOldStr = inputbox("Old text")
sNewStr = inputbox("New text")
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(spath)
    Set fc = f.Files
    For Each f1 In fc
        fname1 = f1.Name
        Set f = fs.OpenTextFile(f1.Path, 1, 0)
        stext = f.readall
        stext = Replace(stext, sOldStr, sNewStr)
        fpath = f1.Path
        Set a = fs.CreateTextFile(fpath, True)
        a.Write stext
        a.Close
    Next
Msgbox "Done!"