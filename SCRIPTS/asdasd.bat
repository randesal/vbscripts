set sd=%date:~7,2%
set sm=%date:~4,2%
set sy=%date:~10,4%
set file1=%sy%-%sm%-*
set file2=%sy%%sm%%sd%
set sPath=\\192.168.1.75\file server\Backend Machine Info Log\
set test=\Status\Test Statistics\%file1%.csv
set dest=C:\Adrian\Testdata\

echo f|copy "%sPath%SS 1%test%" "%dest%S1.tmp"
echo f|copy "%sPath%SS 2%test%" "%dest%S2.tmp"
echo f|copy "%sPath%SS 3%test%" "%dest%S3.tmp"
echo f|copy "%sPath%SS 4%test%" "%dest%S4.tmp"
echo f|copy "%sPath%SS 5%test%" "%dest%S5.tmp"
echo f|copy "%sPath%SS 6%test%" "%dest%S6.tmp"
echo f|copy "%sPath%SS 7%test%" "%dest%S7.tmp"
echo f|copy "%sPath%SS 8%test%" "%dest%S8.tmp"
echo f|copy "%sPath%SS 9%test%" "%dest%S9.tmp"
echo f|copy "%sPath%PT1%test%" "%dest%P1.tmp"
echo f|copy "%sPath%PT2%test%" "%dest%P2.tmp"
echo f|copy "%sPath%PT3%test%" "%dest%P3.tmp"
echo f|copy "%sPath%PT4%test%" "%dest%P4.tmp"
echo f|copy "%sPath%PT5%test%" "%dest%P5.tmp"
echo f|copy "%sPath%PT6%test%" "%dest%P6.tmp"

copy *.tmp "%dest%\Old\Compiled_%file2%.dat"
echo d|xcopy "%dest%\Old\*.dat" "%dest%" /y /d:%sm%-%sd%-%sy%
del *.csv
copy *.dat Compiled.csv
del "%dest%*.tmp"
del "%dest%*.dat"


