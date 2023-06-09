net use z: /delete
net use z: "\\192.168.16.144\output"  /persistent:YES
xcopy  "z:\*.xls" "\\192.168.1.7\Nasserver_TEST\ATM Data Backup" /s /y /d
xcopy "z:\*.csv" "\\192.168.1.7\Nasserver_TEST\ATM Data Backup" /s /y /d
PAUSE