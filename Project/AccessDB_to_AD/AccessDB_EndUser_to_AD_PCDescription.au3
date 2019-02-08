;~ #Region ;**** Directives created by AutoIt3Wrapper_GUI ****
;~ #AutoIt3Wrapper_Icon=D:\temp\favicon.ico
;~ #AutoIt3Wrapper_Compile_Both=y
;~ #AutoIt3Wrapper_UseX64=y
;~ #EndRegion
;**** Directives created by AutoIt3Wrapper_GUI ****
;#AutoIt3Wrapper_AU3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
;#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
;---------------------------------------------------------------------------------------------
; Abfrage auf Tabellen im Access-DB liefert Werte für AD-Computer [Beschreibung].
;
; Test OK.
; Voraussetzung AD-Administrator & PowerSchell import-module ActiveDirectory
; Language: AutoIt3
; ODBC 3.51 Und höher
; Win_10
;---------------------------------------------------------------------------------------------
#include <AD.au3>

;~ ; Open Connection to the Active Directory

;~ _AD_Open()
;~ ;_AD_Open("funkegruppe\it", "geheim1")
;~ If @error Then Exit MsgBox(16, "Active Directory Skript", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)

;~ Global $iReply = MsgBox(308, "Active Directory Functions", "Script fügt neue Computerbeschreibung in AD aus Access DB." & @CRLF & @CRLF & "Sind Sie sicher, dass Sie das ändern möchten?")
;~ If $iReply <> 6 Then Exit

$dbname = "\\fileserver-fhu\edv\Inventory-Datenbank\NetworkCommon.mdb"
$tblname = "dbo_DevicesALL"
$query = "SELECT * FROM " & $tblname & " WHERE Type ='PCL' OR Type = 'PCD'" 				; Produktiv
;$query = "SELECT * FROM " & $tblname & " WHERE Device = 'PLPCD605' AND Type ='PCD'"		; (M1) - zum Testen verwenden
;$query = "SELECT * FROM " & $tblname & " WHERE Type ='PHM'"

Local $code
$adoCon = ObjCreate("ADODB.Connection")
$adoCon.Open("Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & $dbname) ;Use this line if using MS Access 2003 and lower
;$adoCon.Open("Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & $dbname)

;$adoCon.Open ("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $dbname) ;Use this line if using MS Access 2007 and using the .accdb file extension
$adoRs = ObjCreate ("ADODB.Recordset")

$adoRs.CursorType = 1
$adoRs.LockType = 3
$adoRs.Open ($query, $adoCon)

While Not $adoRs.EOF
	Local $sHost = $adoRs.Fields("HostName").value 								; Retrieve value by field name
	Local $sUser = StringReplace($adoRs.Fields("EndUser").value, " ", "_")		; Replace a blank space (' ') with a _ character.
	if StringLen($sUser) > 0 Then
		;MsgBox(0,"Gefunden","Host: " & $sHost & @CR & "User: " & $sUser, 0)
		; Enter parent computer to modify per PowerShell
		;$sCommand1 = 'powershell -Command import-module ActiveDirectory; "Get-ADComputer -Filter * | Select -Expand Name"'
		$sCommand2 = 'Set-ADComputer ' & $sHost & ' -Description ' & $sUser
		;$sCommand1 = 'powershell -Command import-module ActiveDirectory; ' & $sCommand2
		$sCommand3 = 'powershell -Command ' & $sCommand2
		;run("cmd /k " & $sCommand1 & " ; " & $sCommand2)
		;run("cmd /k " & $sCommand3)
		;RunWait($sCommand3, "", @SW_HIDE)
		;RunWait(@ComSpec & ' /k PowerShell.exe -Command "& {Start-Process PowerShell.exe -ArgumentList ' & "'-ExecutionPolicy Bypass -File " & '""' & $sPSScript & '"' & "' -Verb RunAs}")
		;Local $iPID = Run(@ComSpec & ' /k PowerShell.exe -Command ' & $sCommand2, @SW_HIDE, $STDOUT_CHILD)		; Produktiv


		;Local $iPID = Run(@ComSpec & ' /k PowerShell.exe -Command ' & $sCommand2 )				; (M1) - zum Testen verwenden
		;Local $iPID = Run('powershell.exe -executionpolicy bypass -windowstyle hidden -noninteractive -nologo ' & $sCommand2, @SW_HIDE, $STDOUT_CHILD)
		Local $iPID = Run('powershell.exe -executionpolicy bypass -windowstyle hidden -noninteractive -nologo ' & $sCommand2 )
		Sleep(1000)
		ProcessWaitClose($iPID)
		ConsoleWrite($sCommand3 & @CR )
	EndIf
	$adoRs.MoveNext
WEnd
; Close Connection to the Access-DB
$adoCon.Close

;~ ; Close Connection to the Active Directory
;~ _AD_Close()