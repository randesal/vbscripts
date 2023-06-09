del "\\192.168.16.144\output\*.tmp" /s /q
xcopy "\\192.168.16.144\output\*.xls" "\\192.168.1.7\Nasserver_TEST\ATM Data Backup" /s /y /d
xcopy "\\192.168.16.144\output\*.csv" "\\192.168.1.7\Nasserver_TEST\ATM Data Backup" /s /y /d