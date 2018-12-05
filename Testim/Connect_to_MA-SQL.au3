; This should get you connected to MS SQL without an ODBC.

Local $sStrNanchnameOK = "Demur"
Local $sStrVornameOK = "Mur"
Local $nPersNr = 197
Local $sStrADUser = "DemurMu"
Local $sPasswd = "geheim1"
Local $sStrADUseBIG = "DEMURMU"

Local $bGefunden = False
Local $sEmail = "m.demur@funkegruppe.de"

;Local $sMySQL = ("SELECT * FROM benutzer WHERE login_ad LIKE 'kuchtaan'")
Local $sMySQL = ("SELECT * FROM benutzer WHERE login_ad LIKE '" & $sStrADUser &"'")
ConsoleWrite($sMySQL & @LF)
Local $sValues = ("'" & $sStrNanchnameOK & "', '" & $sStrVornameOK & "', '" & $nPersNr & "', '" &  $sEmail  & "', '" &  $sStrADUser  & "', '" &  $sPasswd  & "', '" &  $sStrADUseBIG & "', '" &  $sPasswd & "', 'A', 'benutzer', 'on'")
ConsoleWrite($sValues & @LF)

; Only for SQL server
$objConn = ObjCreate("ADODB.Connection")
$objConn.Open("Provider='sqloledb';Data Source='sql1';Initial Catalog='benutzer';User ID='intranet_allgemein';Password='Gt_Fu_57kS';")
;$rsCustomers = $objConn.Execute("SELECT * FROM Customers")
;$rsCustomers = $objConn.Execute("SELECT * FROM benutzer WHERE login_ad LIKE 'kuchtaan'")
$rsCustomers = $objConn.Execute($sMySQL)
With $rsCustomers
    While Not .EOF
        ConsoleWrite(.Fields("nachname").Value & " - " & .Fields("login_ad").Value & " - " & .Fields("kennwort_ad").Value & @LF)
		$bGefunden = True
        .MoveNext
    WEnd
    .Close
EndWith

If $bGefunden Then
	MsgBox(64, "User Intranet", "Im Intanet ist User '" & $sStrADUser & "' bereits vorhanden. Angaben sind unverändert")
Else
	; "INSERT INTO employees (empName, empTitle) VALUES ('Josh Walker', 'Manager')"
	;$rsCustomers = $objConn.Execute("INSERT INTO benutzer (nachname, vorname, personalnummer, email, login_ad, kennwort_ad, login_as400) VALUES ($sStrNanchnameOK, $sStrVornameOK, $nPersNr, $sEmail, $sStrADUser, $sPasswd, $sStrADUseBIG & ")")
	$rsCustomers = $objConn.Execute("INSERT INTO benutzer (nachname, vorname, personalnummer, email, login_ad, kennwort_ad, login_as400, kennwort_as400, status, art, status_intranet) VALUES (" & $sValues & ")")
	MsgBox(64, "User Intranet", "Im Intanet ist User '" & $sStrADUser & "' neu hinterlegt")
EndIf

$objConn.Close