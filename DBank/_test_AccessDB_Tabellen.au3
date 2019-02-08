;---------------------------------------------------------------------------------------------
; Abfrage erstellt eine Liste allen Tabellen aus dem Access-DB.
;
; Test OK
; Language: AutoIt3
; ODBC 3.51 Und höher
; Win_10
;---------------------------------------------------------------------------------------------

Func ListTables()
$Err=ObjEvent("AutoIt.Error","nothing")

$conn = ObjCreate("ADODB.Connection")
$cat = ObjCreate("ADOX.Catalog")
$tbl = ObjCreate("ADOX.Table")


;~ $result=$conn.Open("Driver={Microsoft Access Driver (*.mdb)};" & _
;~       "DBQ=" & @ScriptDir & "\access.mdb" )

$dbname = "\\fileserver-fhu\edv\Inventory-Datenbank\NetworkCommon.mdb"
$result=$conn.Open("Driver={Microsoft Access Driver (*.mdb)};" & "DBQ=" & $dbname )

$cat.ActiveConnection = $conn
$TableName=""
For $h=1 to $cat.Tables.Count-1
  $TableName=$TableName & $cat.Tables.Item($h).Name & @CR
Next

$conn.Close
$conn=0
$cat=0
$tbl=0
  MsgBox(64,"Gefundene Tabellen:",$TableName)
EndFunc
;---------------------------------------------------------------------------------------------
ListTables()  ; Starte die Abfrage