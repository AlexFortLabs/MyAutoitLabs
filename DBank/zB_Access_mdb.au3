; hier ein Beispiel für das schreiben in eine Access mdb und das auslesen einer Access accdb über die Access.au3.
; Willst du direkt auf eine MS SQL Zugreifen, kannst du dir die sql_udf.au3 ansehen.
; Hier habe ich leider nur Beispiel zum lesen

;--------- Modify the variables below as desired or applicable -------
local $dbname = @ScriptDir & "\Report.mdb"
$tblname = "Reporting"
$fldname = "Kennziffer"
$fldname1 = "code"
$fldname2 = "Woher"
$fldname2 = "Sonstiges"


_access_write($w1, $w2, $w3, $w4)


Func _access_write($w1, $w2, $w3, $w4)






    $FULL_MDB_FILE_NAME = $dbname
    ;$SQL_CODE = "select * from Reporting"
    $CONN = ObjCreate("ADODB.Connection")
    $CONN.Open('Driver={Microsoft Access Driver (*.mdb)};Dbq=' & $FULL_MDB_FILE_NAME & ';')
    $RecordSet = ObjCreate("ADODB.Recordset")






    Local $sQuery = "INSERT INTO Reporting (`Kennziffer`,`code`,`Woher`,`Sonstiges`)" & _
                    "VALUES ('" & $w1 & "',"& _     ;
                    "'" & $w2 & "'," & _            ;
                    "'" & $w3 & "'," & _            ;
                    "'" & $w4 &  "')"                 ;
    ;ConsoleWrite($sQuery)
    $CONN.Execute($sQuery)
    $CONN.Close






EndFunc


; Hier noch ein Beispiel für das ausklesen einer Access accdb...


;$datenbankpfad = @ScriptDir & "\Test.accdb"
$datenbankpfad = "d:\Test.accdb"
$tabellenname = "Tabellentest"
$Spalte0 = "ID"    ;nicht ID Spalte
$Spalte1 = "Stick"    ;nicht ID Spalte
$Spalte2 = "Box"        ;nicht ID Spalte


$query = "SELECT * FROM " & $tabellenname & " WHERE Box = '12' AND Stick = '3'"
$strData1 = _ReadOneFld($query, $datenbankpfad,$Spalte0)


ConsoleWrite($query& @CRLF &$strData1)
MsgBox(0,"",$query & @CRLF & "Ergebnis: " & $strData1)


Func _ReadOneFld($_sql, $_datenbankpfad, $_field)
    Dim $_output
    $adoCon = ObjCreate("ADODB.Connection")
    $adoCon.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & $datenbankpfad & ";")
    $adoRs = ObjCreate("ADODB.Recordset")
    $adoRs.CursorType = 1
    $adoRs.LockType = 3
    $adoRs.Open($_sql, $adoCon)


    With $adoRs
        If .RecordCount Then
            While Not .EOF
                $_output = $_output & .Fields($_field).Value
                .MoveNext
            WEnd
        EndIf
    EndWith
    $adoCon.Close
    Return $_output
EndFunc


;---------------------- schreiben in eine Access accdb --------


$s_db_pwd = "12345678"


_access_write()


Func _access_write()


	$s_dbname = "D:\TestDB.accdb"
	$datenbankpfad = $s_dbname


	$s_data01 = @username
	$s_data02 = @ComputerName






	$adoCon = ObjCreate("ADODB.Connection")
    $adoCon.Open("Provider=Microsoft.ACE.OLEDB.12.0; Jet OLEDB:Database Password=" & $s_db_pwd & "; Data Source=" & $datenbankpfad & ";")
	$adoRs = ObjCreate("ADODB.Recordset")
    ;$adoRs.CursorType = 1
    ;$adoRs.LockType = 3


	Local $sQuery = "INSERT INTO TB1 (`feld1`,`feld2`)" & _
					"VALUES ('" & $s_data01 & "',"& _ 	;User
					"'" & $s_data02  & "')" 			;PC


	$adoCon.Execute($sQuery)
	$adoCon.Close


EndFunc


;-------- update auf eine access accdb ----------
_access_write_update()


Func _access_write_update()


	$dbname = "D:\Data\scripte\MS_AccessCom\access_01\db_test.accdb"
	$tblname = "epayslip"
	$s_data07 = _Now()


	$FULL_MDB_FILE_NAME = $dbname
	;$SQL_CODE = "select * from epayslip"
	$CONN = ObjCreate("ADODB.Connection")
	;$CONN.Open('Driver={Microsoft Access Driver (*.mdb)};Dbq=' & $FULL_MDB_FILE_NAME & ';')
	$CONN.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & $FULL_MDB_FILE_NAME & ";")
	$RecordSet = ObjCreate("ADODB.Recordset")


	;Local $sQuery = "UPDATE " & $tblname & " SET check = 1, date = " & $s_data07 & " WHERE user = " & "'" & @UserName & "'"
	Local $sQuery = "UPDATE " & $tblname & " SET " & $tblname & ".[check] = 1, " & $tblname & ".[date] = '" & $s_data07 & "' WHERE " & $tblname & ".[user] = " & "'" & @UserName & "'"
	;ConsoleWrite($sQuery)
	$CONN.Execute($sQuery)
	$CONN.Close


EndFunc