On Error Resume Next

Dim objShell, strTemp
Set objShell = WScript.CreateObject("WScript.Shell")

strTemp = "HKEY_CURRENT_USER\Volatile Environment\USERNAME"
WScript.Echo "Logged in User: " & objShell.RegRead(strTemp) 

strTemp = "HKEY_CURRENT_USER\Volatile Environment\USERPROFILE"
WScript.Echo "User Home: " & objShell.RegRead(strTemp) 