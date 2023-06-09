color 0A
cls
ForFiles /P "U:\BACKEND AOI DATA\csv4\TongHsing-TOP" /M "*MAPMERGER" /D -5 /C "Cmd /C If @isdir==TRUE del /s /q /f @path && rd /s /q @path"
ForFiles /P "U:\BACKEND AOI DATA\csv4\TongHsing-BOT" /M "*MAPMERGER" /D -5 /C "Cmd /C If @isdir==TRUE del /s /q /f @path && rd /s /q @path"
ForFiles /P "U:\Backend-test\test e-map\_Emap Data" /M "*MAPMERGER" /D -5 /C "Cmd /C If @isdir==TRUE del /s /q /f @path && rd /s /q @path"
PAUSE