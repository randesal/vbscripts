Set objShell = CreateObject ("WScript.Shell")
objshell.run "ssh.exe -D 1080 conference@192.168.45.8"
WScript.Sleep 2000
objshell.sendkeys "conference"
WScript.Sleep 500
objshell.sendkeys "{enter}"