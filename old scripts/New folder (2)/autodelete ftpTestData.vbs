spath = "\\192.168.1.75\file server\FTP Test DATA"
  Set fs = CreateObject("Scripting.FileSystemObject")
  sLog = "C:\Autodelete\FTPTESTdelete.log"
  If fs.fileexists(sLog) = False Then fs.CreateTextFile sLog
  Set h = fs.GetFile(sLog)
  Set f = fs.GetFolder(spath)
    Set gc = f.subfolders
        For Each g1 In gc
        d = g1.Name
        e = g1.Path
        slen = Len(g1.Name) <> 8
        sdate = g1.datelastmodified
        s = sdate < Now - 100
            If s And slen Then
            fs.deletefolder e
	'msgbox e
            Set ts = h.OpenAsTextStream(8, -2)
            ts.Write e & "," & sdate & "," & Now & vbCrLf
            ts.Close
            End If
        Next
