; Verbindung zu DB2 / AS 400 ohne ODBC.

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

; Variablen neu
Local $bUserNeu = True
; Variablen von übergabe
Local $sPersNr = 196														; Personalnummer 	496

$sqlRs = ObjCreate("ADODB.Recordset")
If Not @error Then

;~ 	;$sqlRs.open ("select * from SMKDIFP.XSB00 WHERE SBSBNR LIKE " & $sPersNr, $sqlCon)		; +Produktivumgebung 	(SMKDICP)
;~ 	$sqlRs.open ("select * from SMKDIFT.XSB00 WHERE SBSBNR LIKE " & $sPersNr, $sqlCon)		; +Testumgebung			(SMKDICT)
;~     If Not @error Then
;~         ;Loop bis Dateiende
;~         While Not $sqlRs.EOF
;~          ;Rufen Sie Daten aus den folgenden Feldern ab
;~          $sfPersNr = $sqlRs.Fields ('SBSBNR' ).Value						; S - Peronalnummer		496
;~ 			$sfAbt = $sqlRs.Fields ('SBABT' ).Value							; S - Abteilung			1
;~ 			$sfFirma = $sqlRs.Fields ('SBFNR' ).Value						; A - Firma				01
;~          $sfName = $sqlRs.Fields ('SBNAME' ).Value						; A - Sachbearbeiter-Name 		Kai Pelzhof
;~ 			$sfNName = $sqlRs.Fields ('SBNNAM' ).Value						; A - Kurzname zur Anzeige			Pelzhof
;~ 			$sfKurz = $sqlRs.Fields ('SBZEI' ).Value						; A - SB-Kürzel		Pelzhof
;~ 			$sfNameN = $sqlRs.Fields ('SBNAMN' ).Value						; A - Nachname			Pelzhof
;~ 			$sfPasword = $sqlRs.Fields ('SBKENN' ).Value					; A - Kennwort 			**********
;~ 			$sfMail = $sqlRs.Fields ('SBEMAL' ).Value						; A - E-mail Adresse				k.pelzhof@funkegruppe.de
;~ 			$sfBerechtigung = $sqlRs.Fields ('SBBSTU' ).Value				; S - Berechtigungs-Level >	10 / 70 / 90 / 91 / 99
;~ 			;$sfBerechtExtra = $sqlRs.Fields ('SBRE12' ).Value				; P - Berecht. Extr.	0 / 1 / 70 / 90 / 91 / 99
;~ 			$sfMenu = $sqlRs.Fields ('SBMENU' ).Value						; A - Sartmenü			BASIS /IPRROOT / ILOROOT /IEKROOT / IAVROOT / IROOT
;~ 			$sfTel = $sqlRs.Fields ('SBTELF' ).Value						; A - Telefon-Nummer		02388 / 3071-
;~ 			$sfDurchW = $sqlRs.Fields ('SBDWHL' ).Value						; A - Durchwahlnummer			160
;~ 			$sfProfiele = $sqlRs.Fields ('SBUSPF' ).Value					; A - Benutzerprofiel		PELZHOFKA

;~             MsgBox(16, "Eintrag gefunden für", "NR:  " & $sfPersNr & @CRLF & "NAME:  " & $sfName & @CRLF & "Abteilung:  " & $sfAbt & @CRLF & "Firma:  " & $sfFirma & @CRLF & "Rechte:  " & $sfBerechtigung _
;~ 				& @CRLF & "Tel.:  " & $sfTel & @CRLF & "MENU:  " & $sfMenu & @CRLF & "PASS:  " & $sfPasword _
;~ 				& @CRLF & "Profil:  " & $sfProfiele)

;~ 			; Um Daten anzupassen
;~ 			; $sqlRs.FIELDS('"' & $sfPersNr & '"') = ".F." ; ADDED THIS LINE
;~             ; $sqlRs.Update  ; ADDED THIS LINE

;~ 		    $bUserNeu = False
;~             $sqlRs.MoveNext
;~         WEnd

		If $bUserNeu Then
			MsgBox(64, "User erstellen", "Kein Eintrag gefunden für NR:  " & $sPersNr & @CRLF & "<!> User wir erstellt <!>")

			$sAbteilung = 1
			$sFirma = "01"
			$sName = "Pjer Rischar"
			$sNameV2 = "Rischar, Pjer"
			$sKurzname = "Rischar"
			$sPasswd = "geheim"
			$sMail = "p.rischar@funkegruppe.de"
			$sRechte = 70
			$sMenu = "IEKROOT"
			$sProfiel = "RISCHARPJ"

			$bWaage	= False
			Switch MsgBox(3, 'In Logistik', "Waage User?")
				Case 6 ; JA
					; Code für JA
					$bWaage	= True
				Case 7 ; NEIN
					; Code für NEIN
					$bWaage	= False
				Case 2 ; CANCEL
					$bWaage	= False
					Exit
			EndSwitch

			; Sachbearbeiter SoftM erstellen
			$vRow = "SBSBNR, SBABT, SBFNR, SBNAME, SBNNAM, SBZEI, SBNAMN, SBKENN, SBEMAL, SBBSTU, SBMENU, SBUSPF"
			$vValue = $sPersNr & ", " & $sAbteilung & ", '" & $sFirma & "', '" & $sName & "', '" & $sKurzname & "', '" & $sKurzname & "', '" & $sKurzname & "', '" & StringUpper($sPasswd) & "', '" & _
						$sMail & "', " & $sRechte & ", '" & $sMenu & "', '" & $sProfiel & "'"

			;MsgBox(64, "Testumgebung", "INSERT INTO SMKDIFT.XSB00 (" & $vRow & ") VALUES(" & $vValue & ")")
			MsgBox(64, "Produktivumgebung", "INSERT INTO SMKDIFP.XSB00 (" & $vRow & ") VALUES(" & $vValue & ")")
			$sqlCon.execute ("INSERT INTO SMKDIFP.XSB00 (" & $vRow & ") VALUES(" & $vValue & ")")	; !!! Password wird nicht gespeichert !!!

			; wenn Sachbearbeiter aus Logistik-Waage dann (6+PersNr) für zweites Konto
			if $bWaage Then
				$vValue = 6000 + $sPersNr & ", " & $sAbteilung & ", '" & $sFirma & "', '" & $sName & "', '" & $sKurzname & "', '" & $sKurzname & "', '" & $sKurzname & "', '" & StringUpper($sPasswd) & "', '" & _
							$sMail & "', " & $sRechte & ", '" & $sMenu & "', '" & $sProfiel & "'"
				;MsgBox(64, "Testumgebung", "INSERT INTO SMKDIFT.XSB00 (" & $vRow & ") VALUES(" & $vValue & ")")
				MsgBox(64, "Profuktivumgebung", "INSERT INTO SMKDIFP.XSB00 (" & $vRow & ") VALUES(" & $vValue & ")")
				$sqlCon.execute ("INSERT INTO SMKDIFP.XSB00 (" & $vRow & ") VALUES(" & $vValue & ")")	; !!! Password wird nicht gespeichert !!!
			EndIf

			; Benutzer / Benutzerprofiel auf AS400 erstellen
			$CMD2 = "RMTCMD CRTUSRPRF USRPRF(" & $sProfiel & ") PASSWORD(" & $sPasswd & ") PWDEXP(*NO) INLMNU(*SIGNOFF) TEXT('" & $sNameV2 & "') SPCAUT(*NONE) JOBD(QGPL/QDFTJOBD) GRPPRF(SOFTM) OWNER(*GRPPRF)"
			MsgBox(0, "", "Profiel: " & $CMD2)
			RunWait(@ComSpec & " /q /c " & $CMD2,@ScriptDir,@SW_HIDE)
		EndIf
		$sqlCon.close
    ;EndIf
EndIf


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