dim sPath
dim sFolder
dim sNewFolder

sPath = "C:\Users\BackendPE\Desktop\New folder\"

set fs = createobject("scripting.filesystemobject")
set f = fs.GetFolder(sPath)

Do 
  WScript.Sleep 3000
  seeChanges
Loop

Sub seeChanges()
  sFolder = right("0" & Month(now),2) & right("0" & Day(now),2)
  sNewFolder = sPath & sFolder
  
  set fc = f.Files
  if fc.Count = 0 then Exit Sub

  if fs.FolderExists(sNewFolder)=False then fs.CreateFolder(sNewFolder)

  For Each f1 In fc
    if fs.FileExists(sNewFolder & "\" & f1.Name) then fs.DeleteFile(sNewFolder & "\" & f1.Name)
    fs.MoveFile f1.Path,sNewFolder & "\" & f1.Name
  Next
End Sub
