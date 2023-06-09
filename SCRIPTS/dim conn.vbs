'on error resume next
dim conn
dim insestSQL
dim conString

dim sFile
dim Currtime
dim sYear
dim sMonth
dim sDay
dim Mach(18)

conString = "Driver={MySQL ODBC 8.0 UNICODE Driver};Server=192.168.1.7;Port=3306;Database=Ceramic_Data;User=Backend_PE;Password=Backdoor0122!;Option=3;"

for i = 0 to 7
    strdate = formatdatetime(now - i,2)
    'msgbox strDate 

    '<Date>
        sYear  = year(strDate)
        sMonth = month(strDate)
        if len(month(strDate))=1 then sMonth = "0" & month(strDate)
        sDay   = day(strDate)
        if len(day(strDate))=1   then sDay   = "0" & day(strDate)
        sFile  = sYear & "-" & sMonth & "-" & sDay & ".csv"
    '</Date>

    '<Machine Names>
        Mname = split("PT1,PT2,PT3,PT4,PT5,PT6,PT7,PT8,PT9,SS1,SS2,SS3,SS4,SS5,SS6,SS7,SS8,SS9,SS10",",")
    '</Machine Names>

    '<Machine Path>s
        Mach(0) = "\\192.168.1.75\file server\Backend Machine Info Log\PT1\Status\Test Statistics\" & sFile
        Mach(1) = "\\192.168.1.75\file server\Backend Machine Info Log\PT2\Status\Test Statistics\" & sFile
        Mach(2) = "\\192.168.1.75\file server\Backend Machine Info Log\PT3\Status\Test Statistics\" & sFile
        Mach(3) = "\\192.168.1.75\file server\Backend Machine Info Log\PT4\Status\Test Statistics\" & sFile
        Mach(4) = "\\192.168.1.75\file server\Backend Machine Info Log\PT5\Status\Test Statistics\" & sFile
        Mach(5) = "\\192.168.1.75\file server\Backend Machine Info Log\PT6\Status\Test Statistics\" & sFile
        Mach(6) = "\\192.168.1.75\file server\Backend Machine Info Log\PT7\Status\Test Statistics\" & sFile
        Mach(7) = "\\192.168.1.75\file server\Backend Machine Info Log\PT8\Status\Test Statistics\" & sFile
        Mach(8) = "\\192.168.1.75\file server\Backend Machine Info Log\PT9\Status\Test Statistics\" & sFile
        Mach(9) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 1\Status\Test Statistics\" & sFile
        Mach(10) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 2\Status\Test Statistics\" & sFile
        Mach(11) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 3\Status\Test Statistics\" & sFile
        Mach(12) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 4\Status\Test Statistics\" & sFile
        Mach(13) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 5\Status\Test Statistics\" & sFile
        Mach(14) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 6\Status\Test Statistics\" & sFile
        Mach(15) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 7\Status\Test Statistics\" & sFile
        Mach(16) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 8\Status\Test Statistics\" & sFile
        Mach(17) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 9\Status\Test Statistics\" & sFile
        Mach(18) = "\\192.168.1.75\file server\Backend Machine Info Log\SS 10\Status\Test Statistics\" & sFile
    '</Machine Path>

    ' <createString>
        createString = "CREATE TABLE IF NOT EXISTS `Ceramic_Data`.`OSTest` (`unixtime` int(10), `Date` varchar(10), `Time` varchar(8), `LotNo` varchar(24), `Product` varchar(31), `Tileid` varchar(32), `Pass` int(3), `Open` int(2), `Short` int(1), `OpenShort` int(1), `4Line Error` int(1), `NonTest` int(1), `High Volt(V)` int(3), `Continuity(Ohm)` int(3), `Isolation(Ohm)` varchar(3), `Four Wire Base Value(mOhm)` decimal(3,2), `Test Time` int(3), `Software Revision` varchar(23), `Low Resistance Value` varchar(5), `Low Resistance Ratio` int(2), `Low Resistance Ratio value` decimal(4,2), `Standard Resistance(1ohm)` varchar(6), `Standard Resistance(10ohm)` varchar(6), `Test Mode` varchar(5), `Total stroke` int(2)) ENGINE = InnoDB DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci;"
    '</createString>

    for x = lbound(Mach) to ubound(Mach)  'insert to db
        set fs = createobject("Scripting.Filesystemobject")
        set conn = createobject("ADODB.Connection")
        if fs.fileexists(Mach(x)) then 
            set f  = fs.OpenTextFile(Mach(x),1,1)
            stext = f.readline
            Do Until f.AtEndOfStream
                stext = f.readline
                strtext = split(stext,",")
                    for y = lbound(strtext) to ubound(strtext)
                        strtext(y) = "'" & strtext(y) & "'"
                        'if isnumeric(strtext(y)) = false then strtext(y) = "'" & strtext(y) & "'"
                    next 
                    sline = join(strtext,",") & ",'" & mname(x) & "'"
                    if ubound(strtext) = 24 then sline = join(strtext,",") & ",0,'" & mname(x) & "'"
                'msgbox lbound(strtext) & ", " &  ubound(strtext)
                insertString = "INSERT INTO `OSTest` (`unixtime`, `Date`, `Time`, `LotNo`, `Product`, `Tileid`, `Pass`, `Open`, `Short`, `OpenShort`, `4Line Error`, `NonTest`, `High Volt(V)`, `Continuity(Ohm)`, `Isolation(Ohm)`, `Four Wire Base Value(mOhm)`, `Test Time`, `Software Revision`, `Low Resistance Value`, `Low Resistance Ratio`, `Low Resistance Ratio value`, `Standard Resistance(1ohm)`, `Standard Resistance(10ohm)`, `Test Mode`, `Total stroke`, `Step No`, `Machine`) VALUES (" & sline & ") ON DUPLICATE KEY UPDATE `unixtime`= value(`unixtime`), `Date`= value(`Date`), `Time`= value(`Time`), `LotNo`= value(`LotNo`), `Product`= value(`Product`), `Pass`= value(`Pass`), `Open`= value(`Open`), `Short`= value(`Short`), `OpenShort`= value(`OpenShort`), `4Line Error`= value(`4Line Error`), `NonTest`= value(`NonTest`), `High Volt(V)`= value(`High Volt(V)`), `Continuity(Ohm)`= value(`Continuity(Ohm)`), `Total stroke`= value(`Total stroke`), `Step No`= value(`Step No`), `Machine`= value(`Machine`);"
                conn.open conString
                'conn.execute(createString)
                conn.execute(insertString)
                conn.close
            Loop
        end if
    next

Next
    
    set conn = nothing
    set fs = nothing

    msgbox "Done Update!"
