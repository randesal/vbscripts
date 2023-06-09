Dim fs, f, f1, fc, s
spath = InputBox("Enter Folder path")
sOldStr = InputBox("Old text")
sNewStr = InputBox("New text")
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(spath)
    Set fc = f.Files
    For Each f1 In fc
        stext = Replace(f1.Path, sOldStr, sNewStr)
        fpath = f1.Path
        fs.movefile fpath,stext
    Next
MsgBox "Done!"