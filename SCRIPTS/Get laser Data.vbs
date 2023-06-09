Option Explicit
main

Sub Main
    Dim cn
    Dim rs
    Dim conString
    Dim SQLString
    Dim field
    Dim sVal
    Dim sLotNo
    Dim sDriver
    Dim sServer
    Dim sPort
    Dim sDatabase
    Dim sUser

    sDriver = "" 'mySQL Driver
    sServer = "" 'Laser server IP
    sPort = "3306" 'Default
    sDatabase = "THEPI_SERVERNEW" 'Laser server table
    sUser = "root"  'Default user no password

    sLotNo = ucase(InputBox("Enter Lot No:" & vbcrlf & "Ex." & vbcrlf & _ 
            "0922J552(1x1),0223OIP2(1x2),0423NME2(2x2)" , "Lot Query", ""))  'Example Lots Only
    'sLotNo = InputBox("Enter Lot No", "Lot Query", "0223OIP2") '1x2
    'sLotNo = InputBox("Enter Lot No", "Lot Query", "0423NME2") '2x2
    if IsEmpty(sLotNO) = true then 
        msgbox "Empty LotNo"
        Exit sub
    end if

    conString = "Driver=" & sDriver & ";Server=" & sServer & ";Port=" & sPort & ";Database=" & sDatabase & ";User=" & sUser & ";"

    Set cn = CreateObject("ADODB.Connection")
        cn.Open conString
                    
        SQLString = "Select * FROM (" & _
                        "SELECT TileText01 as tile_ID FROM " & sLotNo & " Where MarkTime <> '' Union " & _
                        "SELECT TileText02 as tile_ID FROM " & sLotNo & " Where MarkTime <> '' Union " & _
                        "SELECT TileText03 as tile_ID FROM " & sLotNo & " Where MarkTime <> '' Union " & _
                        "SELECT TileText04 as tile_ID FROM " & sLotNo & " Where MarkTime <> '') tbla " & _
                    "WHERE left(tile_ID,3) = left((SELECT TileID FROM " & sLotNo & " LIMIT 1),3) " & _
                    "AND CHAR_LENGTH(tile_ID) = (SELECT CHAR_LENGTH(TileID) FROM " & sLotNo & " LIMIT 1) " & _
                    "AND tile_ID <> '' " & _
                    "ORDER BY Tile_Id ASC "

        Set rs = cn.Execute(SQLString)
        
        While Not rs.EOF
            ' Loop through each field
            For Each field In rs.Fields
                sVal = sVal & field.Value & vbCrLf
            Next
            rs.MoveNext
        Wend
        MsgBox sVal
        cn.Close
        
    Set cn = Nothing
    Set rs = Nothing
End Sub
Msgbox "DONE SQL"