On Error Resume Next

''''''''''''settings'''''''''''''
sFolder1 = "C:\Documents and Settings\Backend PE\Desktop\1"
sFolder2 = "C:\Documents and Settings\Backend PE\Desktop\2"
sFolder3 = "C:\Documents and Settings\Backend PE\Desktop\3"
sTime = "08:00:00"        '8hours

''''''''''''change settings upto here only'''''''''''''


t8hrs = TimeValue(Now()) + TimeValue(sT2)

Do
WSCRIPT.SLEEP(300000) 					'DELAY EXECUTION FOR 300000 MILLISECONDS = 5mins
Call s5Mins(sFolder1, sFolder2)

If TimeValue(Now()) > TimeValue(t8hrs) Then
t8hrs = TimeValue(Now()) + TimeValue(sTime)
Call s8hrs(sFolder2, sFolder3)
End If

Loop

Sub s5Mins(sFolder1, sFolder2)
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.getfolder(sFolder1)
If f.Files.Count = 0 Then Exit Sub
fs.movefile sFolder1 & "\*", sFolder2
End Sub

Sub s8hrs(sFolder2, sFolder3)
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.getfolder(sFolder2)
If f.Files.Count = 0 Then Exit Sub
fs.movefile sFolder2 & "\*", sFolder3
End Sub

