#VARIANT# ==============================================================
;~ $DSN1="DSN=MS Access Database"		; MS Access Database

;~ $adoCon = ObjCreate ("ADODB.Connection")
;~ $adoCon.Open ($DSN1)
;~ $adoRs = ObjCreate ("ADODB.Recordset") ; Create a Record Set to handles SQL Records - SELECT SQL
;~ $adoRs.Open("SELECT * from dbo_DevicesALL order by HostName,Type,Modell", $adoCon)

#VARIANT# ==============================================================
; AutoIt-Access Database 2007 Accdb
; First let’s first create some variables that will hold the database file name (whether .mdb or .accdb), the table name and the query to execute:

$dbname = "\\fileserver-fhu\edv\Inventory-Datenbank\NetworkCommon.mdb"
$tblname = "dbo_DevicesALL"
$query = "SELECT * FROM " & $tblname & " WHERE Location1 = DE"		; The & is simply a concatenation of the strings.
;$query = "SELECT * FROM " & $tblname ;& " WHERE Location1 = DE"		; The & is simply a concatenation of the strings.
Dim $_output

;Let’s set the variable for the one field that we want to retrieve from the database.
Local $title

; Then create the connection to the ADODB:
$adoCon = ObjCreate("ADODB.Connection")

; Then set the Provider. There is a different Provider for each file extension. A .mdb file will have its own Provider and a .accdb file will have another.
; Here is the Provider for a .mdb file:
$adoCon.Open("Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & $dbname)

; Here is the Provider for .accdb file:
;$adoCon.Open ("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $dbname)

; Wir erstellen das Objekt Recordset, setzen einige erforderliche Optionen und führen dann die Abfrage aus:
;~ $adoRs = ObjCreate ("ADODB.Recordset")
;~ $adoRs.CursorType = 1
;~ $adoRs.LockType = 3
;~ $adoRs.Open ($query, $adoCon)

$adoRs = $adoCon.Execute($query)
While Not $adoRs.EOF
	$_output = $_output & $adoRs.fields("title").value & @CRLF
	$adoRs.MoveNext
WEnd
$adoCon.Close
MsgBox(0,"Devicess List",$_output)


#VARIANT# ==============================================================
;~ $dbname = "C:\Users\vk\Documents\db\Marketing\articleSubmissionsTutorialRef.mdb"
;~ $tblname = "articles"
;~ $fldname = "UserName"
;~ $query = "SELECT * FROM " & $tblname & " WHERE articleID = '4'"
;~ Dim $_output
;~ $adoCon = ObjCreate("ADODB.Connection")
;~ $adoCon.Open("Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & $dbname)
;~ ;$adoCon.Open("Driver={Microsoft Access Driver (*.accdb)}; DBQ=" & $dbname)
;~ $adoRs = $adoCon.Execute($query)
;~ While Not $adoRs.EOF
;~ 	$_output = $_output & $adoRs.fields("title").value & @CRLF
;~ 	$adoRs.MoveNext
;~ WEnd
;~ $adoCon.Close
;~ MsgBox(0,"Guest List",$_output)