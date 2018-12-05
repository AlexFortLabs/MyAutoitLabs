#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=D:\temp\favicon.ico
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;~ ' Funktion um das Password verschlüsselt in der User Registry zu speichern, damit ein manuelles anmelden entfällt.

;~         ' Folgende Grundregeln scheinen bei SoftM aktiv zu sein.
;~         ' Alle Zeichen werden in Großbuchstaben umgewandelt.
;~         ' Danach wird dezimal 139 zu dem dezimal Wert hinzugerechnet.
;~         ' Die entsprechenden Werte werden als Asci Werte gespeichert.
;~         ' Angehängt wird Hex B4 (Dezimal 180)  und Hex 88 (Deziaml 136)

;~         ' Dieser String wird als rg_sz Wert in der Registry gepseichert:
;~         ' HKCU\Software\SoftM\Global\Environment\Password.

;~         ' Idee ist es das Klartext Passwort aus der SoftM Sachbearbeitertabelle direkt auszulesen, bzw.
;~         ' zu einem späteren Zeitpunk dort zu aktualisieren. (Paswort synchronisation)

#include <MsgBoxConstants.au3>
#include <String.au3>

;~ Const $cSoftMServer = "FunkeI5"
;~ Const $cSoftMLibrary = "SMKDIFP"
Const $cSoftMCompany = "01"      	; Default Werte
Const $cSoftMDepartment = 1 	  	; Default Werte
Const $sEndeGelende = "´ˆ"			; PassEnde

;###############################################
; Angemeldeter User & Password sollen automatisch ermittelt werden
Local $sUserAngemeldet = StringUpper(@username)
Local $sUser = "123"
Local $sfPasword = "geheim"

;~ ; Verbindung zu DB2 / AS 400 ohne ODBC.

; Initialize COM error handler
$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")

$sqlCon = ObjCreate("ADODB.Connection")
$sqlCon.Mode = 16  			; Erlaubt im MultiUser-Bereich das öffnen anderer Verbindungen ohne Beschränkungen [Lesen/Schreiben/Beides]
$sqlCon.CursorLocation = 3 	; client side cursor Schreiben beim Clienten
$sqlCon.Open ("Provider=IBMDA400;Data Source=FUNKEI5;User Id=SOFTMOPR ;Password=softmopr;")
If @error Then
    MsgBox(0, "ERROR", "Verbindungsproblem zum Datenbank")
    Exit
EndIf

$sqlRs = ObjCreate("ADODB.Recordset")
If Not @error Then
	; User & Password sollen automatisch ermittelt werden
	$sqlRs.open ("select * from SMKDIFP.XSB00 WHERE SBUSPF LIKE " & "'" & $sUserAngemeldet & "'", $sqlCon)		; +Produktivumgebung 	(SMKDICP)
	;$sqlRs.open ("select * from SMKDIFT.XSB00 WHERE SBUSPF LIKE " & "'" & $sUserAngemeldet & "'", $sqlCon)		; +Testumgebung			(SMKDICT)
    If Not @error Then
        ;Loop bis Dateiende
        While Not $sqlRs.EOF
         ;Rufen Sie Daten aus den folgenden Feldern ab
			$sUser = $sqlRs.Fields ('SBSBNR' ).Value						; S - Peronalnummer			496
			;$sfAbt = $sqlRs.Fields ('SBABT' ).Value						; S - Abteilung				1
			;$sfFirma = $sqlRs.Fields ('SBFNR' ).Value						; A - Firma					01
			;$sfName = $sqlRs.Fields ('SBNAME' ).Value						; A - Sachbearbeiter-Name 	Kai Pelzhof
			;$sfNName = $sqlRs.Fields ('SBNNAM' ).Value						; A - Kurzname zur Anzeige	Pelzhof
			;$sfKurz = $sqlRs.Fields ('SBZEI' ).Value						; A - SB-Kürzel				Pelzhof
			;$sfNameN = $sqlRs.Fields ('SBNAMN' ).Value						; A - Nachname				Pelzhof
			$sfPasword = $sqlRs.Fields ('SBKENN' ).Value					; A - Kennwort 				**********  DELL1
			;$sfMail = $sqlRs.Fields ('SBEMAL' ).Value						; A - E-mail Adresse		k.pelzhof@funkegruppe.de
			;$sfBerechtigung = $sqlRs.Fields ('SBBSTU' ).Value				; S - Berechtigungs-Level 	>	10 / 70 / 90 / 91 / 99
			;$sfBerechtExtra = $sqlRs.Fields ('SBRE12' ).Value				; P - Berecht. Extr.		0 / 1 / 70 / 90 / 91 / 99
			;$sfMenu = $sqlRs.Fields ('SBMENU' ).Value						; A - Sartmenü				BASIS /IPRROOT / ILOROOT /IEKROOT / IAVROOT / IROOT
			;$sfTel = $sqlRs.Fields ('SBTELF' ).Value						; A - Telefon-Nummer		02388 / 3071-
			;$sfDurchW = $sqlRs.Fields ('SBDWHL' ).Value					; A - Durchwahlnummer		160
			;$sfProfiele = $sqlRs.Fields ('SBUSPF' ).Value					; A - Benutzerprofiel		PELZHOFKA

            MsgBox(16, "Eintrag gefunden für", "NR:  " & $sUser & @CRLF& "PASS:  " & $sfPasword)
            $sqlRs.MoveNext
        WEnd
	EndIf
EndIf
$sqlCon.close
;###############################################


;###############################################
Local $sCryptedPassword = ""

$sIn = StringUpper ($sfPasword)
$aCharacters = StringSplit($sIn, "")

For $i = 1 To $aCharacters[0]
    ;ConsoleWrite(Asc($aCharacters[$i]) + Asc($cCodeAdd) & " ")
	$sCryptedPassword = $sCryptedPassword & (Chr((Asc($aCharacters[$i]) + Asc("‹"))))
Next
$sCryptedPassword = $sCryptedPassword & $sEndeGelende				; Ergebnis in Registry key schreiben
MsgBox(0, "ASCII Codierung", $sUserAngemeldet & @CRLF & @CRLF & $sCryptedPassword)
;###############################################
DataToReg()

Func DataToReg()
	; Ergebnis in Registry key schreiben: RegWrite ( "keyname", "valuename", "type", value )
	RegWrite("HKEY_CURRENT_USER\Software\SoftM\Global\Environment", "Company", "REG_SZ", $cSoftMCompany)
	RegWrite("HKEY_CURRENT_USER\Software\SoftM\Global\Environment", "Department", "REG_SZ", $cSoftMDepartment)
    RegWrite("HKEY_CURRENT_USER\Software\SoftM\Global\Environment", "User", "REG_SZ", $sUser)
	RegWrite("HKEY_CURRENT_USER\Software\SoftM\Global\Environment", "Password", "REG_SZ", $sCryptedPassword)
EndFunc		;==> DataToReg
;###############################################

Func MyErrFunc()
  $HexNumber=hex($oMyError.number,8)
  Msgbox(0,"COM Test","We intercepted a COM Error !"       & @CRLF  & @CRLF & _
             "err.description is: "    & @TAB & $oMyError.description    & @CRLF & _
             "err.windescription:"     & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "         & @TAB & $HexNumber              & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "     & @TAB & $oMyError.scriptline     & @CRLF & _
             "err.source is: "         & @TAB & $oMyError.source         & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile       & @CRLF & _
             "err.helpcontext is: "    & @TAB & $oMyError.helpcontext _
            )
  SetError(1)  ; to check for after this function returns
Endfunc   ;==>MyErrFunc