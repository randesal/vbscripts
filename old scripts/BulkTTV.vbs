Option Explicit
'<Declaration of Variables>
    '<DB Var>
        dim conn
        dim conStringDB
        dim insertString
        dim createString
        dim sTable
    '<File Var>
        dim sFile
        dim Currtime
        dim sYear
        dim sMonth
        dim sDay
        dim sPath
    '<Data Var>
        dim sLot
        dim sPN
        dim sTileID
        dim sOperator
        dim sMach
        dim nVal
        dim sVal
        dim sValInsert
        dim sDate
        dim val(29)
'<Declaration of variables>
'<initialize>
    conStringDB = "Driver={MySQL ODBC 8.0 UNICODE Driver};" & _
            "Server=192.168.1.7;" & _
            "Port=3306;" & _
            "Database=Ceramic_Data;" & _
            "User=Backend_PE;" & _
            "Password=Backdoor0122!;" & _
            "Option=3;"
    
    sMach = "TTV"
    sTable = "Statistics"
    sVal = ""
    sDate = formatdatetime(now)
    sPath = "\\192.168.1.75\file server\Backend Machine Info Log\PT1\Status\Test Statistics\" & sFile
        For nVal = lbound(val) to ubound(val)
                val(nval) = "Val" & nval+1
                sVal = sVal & "`" & val(nVal) & "` varchar(100), "
                sValInsert = sValInsert & "`" & val(nVal) & "`, "
        next
    sVal = left(sVal,len(sVal)-2)
    sValInsert = left(sValInsert,len(sVal)-2)
    set conn = createobject("ADODB.Connection")

'<initialize>    
        
'<statTable>
    createString = _ 
        "CREATE TABLE IF NOT EXISTS `Ceramic_Data`.`" & sTable & "` (" & _ 
        "`Date` varchar(100), `Time` varchar(100), `" & sMach & "` varchar(100), " & _ 
        "`PartNo` varchar(100), `LotNo` varchar(100), `TileID` varchar(100), " & sVal & _
        ") ENGINE = InnoDB DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci;"
'<statTable>

msgbox sval
msgbox svalInsert
    
'<Insertstring>
    insertString = _ 
    "INSERT INTO " & sTable & " (`Date`, `Time`, `PartNo`, `LotNo`, `TileID`, " & sValInsert & _ 
        ") " & _
    "VALUES ('" & stext & " ON DUPLICATE KEY UPDATE `unixtime`= value(`unixtime`), " & _ 
        "`Date`= value(`Date`), `Time`= value(`Time`), `LotNo`= value(`LotNo`), " & _ 
        "`Product`= value(`Product`), `Pass`= value(`Pass`), `Open`= value(`Open`), " & _ 
        "`Short`= value(`Short`), `OpenShort`= value(`OpenShort`), " & _ 
        "`4Line Error`= value(`4Line Error`), `NonTest`= value(`NonTest`), " & _ 
        "`High Volt(V)`= value(`High Volt(V)`), `Continuity(Ohm)`= value(`Continuity(Ohm)`), " & _ 
        "`Total stroke`= value(`Total stroke`), `Step No`= value(`Step No`), `Machine`= value(`Machine`);"
'</Insertstring>


                conn.open conStringDB
                conn.execute(createString)
         '       conn.execute(insertString)
         '       conn.close

    
    set conn = nothing
    'set fs = nothing

    msgbox "Done Update!"
