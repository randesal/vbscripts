Set objShell = CreateObject ("WScript.Shell")
objshell.run "ssh.exe -D 1080 -L 995:202.124.131.39:995 -L 25:202.124.131.39:25 prod150_2@192.168.1.253"
WScript.Sleep 2000
objshell.sendkeys "ul251"
WScript.Sleep 500
objshell.sendkeys "{enter}"