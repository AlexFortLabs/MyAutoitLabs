#cs ----------------------------------------------------------------------------

 #INDEX# =======================================================================================================================
 Title .........: Quick User Management Tool
 AutoIt Version : 3.3.14.5
 Language ......: DE
 Description ...: Script enthält Funktionen zu Manipulation in Aktive Directory.
 ===============================================================================================================================

#ce ----------------------------------------------------------------------------
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=D:\temp\favicon.ico
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
;#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
;#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>


#include <ComboConstants.au3>
#include <EditConstants.au3>

#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <AD.au3>
#include <_qumvar.au3>

HotKeySet("{ESC}", "_Exit")

; *****************************************************************************

; Bau Verbindung zu Active Directory
;_AD_Open($SUserId, $SPassword,)
;_AD_Open("funkegruppe\idmin", "geheim")
_AD_Open()
If @error Then Exit MsgBox(16, "Active Directory Skript", "Unerwartetes Problem in der Funktion _AD_Open. @error: " & @CRLF & @CRLF & @error & ", @extended = " & @extended)

; Ermittelten FQDN für aktuellen User
Global $sFQDN = _AD_SamAccountNameToFQDN() 			; NT4 Domänen Username zwecks Kompatibilität

Global $iReply = MsgBox(308, "Active Directory Manipulation", "Dieses Skript erstellt einen neuen Benutzer in der angegebener OU." & @CRLF & @CRLF & _
        "Möchten Sie Änderungen in  Active Directory vornehmen?")
If $iReply <> 6 Then
	Exit
Else
	_Anwenderangaben()
EndIf

Func _IsNumberString($String)
; *****************************************************************************
    Return Not StringRegExp($String, "[^0-9]")
EndFunc

Func _StringReplace($sString)
; *****************************************************************************
	For $i In StringSplit("ÃŸ#ss|Ãœ#Ue|Ã„#Ae|Ã–#Oe|Ã¼#ue|Ã¤#ae|Ã¶#oe|&Uuml;#Ue|&uuml;#ue|&ouml;#oe|&Ouml;#Oe|&auml;#ae|&Auml;#Ae|&szlig#ss|ß#ss|Ä#Ae|Ö#Oe|Ü#Ue|ä#ae|ö#oe|ü#ue|Â#A|À#A|Á#A|â#a|á#a|à#a|Ê#E|È#E|É#E|ê#e|é#e|è#e|Î#I|Ì#I|Í#I|î#i|í#i|ì#i|Ô#O|Ò#O|Ó#O|ô#o|ó#o|ò#o|Û#U|Ù#U|Ú#U|û#U|ú#U|ù#U|,#|.#|-#|/#|\#|_#", "|", 2)
		$sString = StringReplace($sString, StringLeft($i, StringInStr($i, "#", 1)-1), StringTrimleft($i, StringInStr($i, "#", 1)), 0, 1)
	Next

	Return $sString

EndFunc   ;==>_StringReplace

Func _StringKurzen($sString, $nLang)
; *****************************************************************************
	Local $sSErgebnis = StringStripWS($sString, $STR_STRIPALL)
	if StringLen($sSErgebnis) > $nLang Then
		$sSErgebnis = StringMid($sSErgebnis, 1, $nLang)   ;  StringMid( "string", start [, count = -1] )
		;$sSErgebnis = StringTrimRight($sSErgebnis, $nLang)
	EndIf

	Return $sSErgebnis
EndFunc	  ;==>_StringKurzen

Func _Anwenderangaben()
; *****************************************************************************
; Personalnummer, Benutzername, Password werden abgefragt
; Frage nach ........
; *****************************************************************************

; Frage nach Nachname!
	$sStrNanchname = InputBox("Benutzer Nachname", "Bitte Nachname eingeben." & @CRLF & " z.B : Mustermann", "", "")
	$sStrNanchnameLang = _StringReplace($sStrNanchname)
	$sStrNanchnameOK = _StringKurzen($sStrNanchnameLang, 8)
	;MsgBox($MB_SYSTEMMODAL, "", $sStrNanchname & " ist umgewandelt in:" & @CRLF & @CRLF & $sStrNanchnameOK)

; Frage nach Vorname!
	$sStrVorname = InputBox("Benutzer Vorname", "Bitte Vorname eingeben."  & @CRLF & " z.B : Karl", "", "")
	$sStrVornameLang = _StringReplace($sStrVorname)
	$sStrVornameOK = _StringKurzen($sStrVornameLang, 2)
	;MsgBox($MB_SYSTEMMODAL, "", $sStrVorname & " ist umgewandelt in:" & @CRLF & @CRLF & $sStrVornameOK)

; Attribute abfragen
	$sDescription = InputBox("Benutzer Beschreibung", "Bitte eingeben." & @CRLF & " z.B : Azubi Vertrieb", $sDescription, "")

; AD-Benutzername Zusammensetzung
	$sStrADUser = $sStrNanchnameOK & $sStrVornameOK		; für AD
	$sStrADUserText = $sStrNanchname & "\, " & $sStrVorname
	$sStrADUserRoh = $sStrNanchname & ", " & $sStrVorname
	$sNickName = StringLower(_StringKurzen($sStrVornameLang, 1) & "." & $sStrNanchnameLang)
	;MsgBox($MB_SYSTEMMODAL, "Nicname", $sNickName)

; Soft-M und i5 Benutzername
	; Personal Nummer =
	Do
		$nPersNr = InputBox("Benutzer Personalnummer", "Tragen Sie bitte die Personalnummer ein", "123")
	Until _IsNumberString($nPersNr)

	; für Benutzerprofil im SoftM-Sachbearbeiter
	$sStrADUseBIG = StringUpper($sStrADUser)				; TOEWSRU
    ;MsgBox($MB_SYSTEMMODAL, "Benutzerprofil", $sStrADUseBIG)

; Frage nach passwort!
    $sPasswd = InputBox("Benutzer Kennwort", "Benutzer " & $sStrADUser & " Passwort eingeben." & @CRLF & "Maximal 9 Zeichen lang." & @CRLF & "Ohne Umlauten", "", "*M9")
    ;MsgBox($MB_SYSTEMMODAL, "", $sPasswd)

EndFunc 	;==>_Anwenderangaben

Func _PasswordSetzen($sUser, $sPasswd)
; *****************************************************************************
;
; Setzen oder löschen Password für User oder Computer
; *****************************************************************************
	Local $iValueModPasswd = _AD_SetPassword($sUser, $sPasswd)
	If $iValueModPasswd = 1 Then
		MsgBox(64, "Active Directory Password", "Passwort für '" & $sUser & "' gesetzt")
		;$Result = _AD_DisablePasswordChange($sUser)
		$Result = _AD_DisablePasswordExpire($sUser)
		;ConsoleWrite("Result: " & $Result & ", error: " & @error & ", extended: " & @extended)
		MsgBox(64, "Active Directory Password: " & $Result, "Error: " & @error & " extended: " & @extended)
	ElseIf @error = 1 Then
		MsgBox(64, "Active Directory Password", "User '" & $sUser & "' does not exist")
	Else
		MsgBox(64, "Active Directory Password", "Return code '" & @error & "' from Active Directory")
	EndIf

EndFunc 	;==>_PasswordSetzen

Func _AttributeSetzenUnivers($sUser, $sAttrubut, $sAttrWert)
; *****************************************************************************
;
; Kontoattributen setzen
; *****************************************************************************
	Local $iValueModDescript = _AD_ModifyAttribute($sUser, $sAttrubut, $sAttrWert)
	If $iValueModDescript = 1 Then
		MsgBox(64, "Active Directory Attribut", "Attribut '" & $sAttrubut & "' hinterlegt")
	ElseIf @error = 1 Then
		MsgBox(64, "Active Directory Attribut", "User '" & $sUser & "' does not exist")
	Else
		MsgBox(64, "Active Directory Attribut", "Return code '" & @error & "' from Active Directory")
	EndIf

EndFunc		;==>_AttributeSertzenUnivers

Func _GruppenMitglied($sUser, $sGroup)
; *****************************************************************************
;
; AD-User zu Gruppe hinzufügen
; *****************************************************************************
	Local $iValue = _AD_AddUserToGroup($sGroup, $sUser)
	If $iValue = 1 Then
		MsgBox(64, "Active Directory Gruppenmitglied", "User '" & $sUser & "' erfolgreich hinzugefügt zu '" & $sGroup & "'")
	ElseIf @error = 1 Then
		MsgBox(64, "Active Directory Gruppenmitglied", "Gruppe '" & $sGroup & "' nicht vorhanden")
	ElseIf @error = 2 Then
		MsgBox(64, "Active Directory Gruppenmitglied", "User '" & $sUser & "' nicht vorhanden")
	ElseIf @error = 3 Then
		MsgBox(64, "Active Directory Gruppenmitglied", "User '" & $sUser & "' bereits Mitglied in '" & $sGroup & "'")
	Else
		MsgBox(64, "Active Directory Gruppenmitglied", "Return code '" & @error & "' from Active Directory")
	EndIf

EndFunc		;==>_GruppenMitglied

Func _BasisOrdnerCreate($sUser)
; *****************************************************************************
;
; Ordner Besitzer & Vollzugriff für User (inklusive Unterordner)
; *****************************************************************************
	Local $sBasisFolder = $sHomeDirectory & $sUser
	DirCreate($sBasisFolder)
	If FileExists($sBasisFolder) Then
		;MsgBox(4096, "", $sBasisFolder & " wurde angelegt.")

		; icacls D:\Beispielordner /grant znil\Beispielgruppe:(CI)(OI)(M)
		Local $sCommand = "icacls " & $sBasisFolder & " /grant " & $sUser & ":(CI)(OI)(F)"
		Local $sCommandOwner = "icacls " & $sBasisFolder & " /setowner " & $sUser & " /T /C" 	; "takeown /U " & $sUser & " /F " & $sBasisFolder & " /R /D Y /SKIPSL"

		Local $iPID = Run(@ComSpec & " /C " & $sCommand, "", @SW_HIDE, 8)
		;MsgBox($MB_SYSTEMMODAL, "PID", $sCommand & ": " & $iPID)

		$iPID = Run(@ComSpec & " /C " & $sCommandOwner, "", @SW_HIDE, 8)
		;MsgBox($MB_SYSTEMMODAL, "PID", $sCommandOwner & ": " & $iPID)
	Else
		MsgBox(4096, "", $sBasisFolder & " Konnte nicht erstellt werden.")
	EndIf

EndFunc 	;==>_BasisOrdner

Func _add_SoftM_User()
; *****************************************************************************
; Benutzer für SoftM und AS400 erstellen
; per Verbindung zu DB2 / AS 400 ohne ODBC.
; *****************************************************************************
; Initialize COM error handler
Global $oMyError = ObjEvent("AutoIt.Error","MyErrFunc")

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
	; Sachbearbeiter SoftM erstellen
	$sName = $sStrVorname & " " & $sStrNanchname			; "Pjer Rischar"
	$sNameV2 = $sStrNanchname & ", " & $sStrVorname			; "Rischar, Pjer"
	$sKurzname = $sStrNanchname								; "Rischar"
	$sMail = $sNickName & "@" & $sDomainName
	$sProfiel = $sStrADUseBIG								; "RISCHARPJ"

	$vRow = "SBSBNR, SBABT, SBFNR, SBNAME, SBNNAM, SBZEI, SBNAMN, SBKENN, SBEMAL, SBBSTU, SBMENU, SBUSPF"
	$vValue = $nPersNr & ", " & $sAbteilung & ", '" & $sFirma & "', '" & $sName & "', '" & $sKurzname & "', '" & $sKurzname & "', '" & $sKurzname & "', '" & StringUpper($sPasswd) & "', '" & _
		$sMail & "', " & $sRechte & ", '" & $sMenu & "', '" & $sProfiel & "'"

	;MsgBox(64, "Testumgebung", "INSERT INTO SMKDIFT.XSB00 (" & $vRow & ") VALUES(" & $vValue & ")")
	MsgBox(64, "Produktivumgebung", "INSERT INTO SMKDIFP.XSB00 (" & $vRow & ") VALUES(" & $vValue & ")")
	$sqlCon.execute ("INSERT INTO SMKDIFP.XSB00 (" & $vRow & ") VALUES(" & $vValue & ")")	; !!! Password wird nicht gespeichert !!!

	; wenn Sachbearbeiter aus Logistik-Waage dann (6+PersNr) für zweites Konto
	if $bWaage Then
		$vValue = 6000 + $nPersNr & ", " & $sAbteilung & ", '" & $sFirma & "', '" & $sName & "', '" & $sKurzname & "', '" & $sKurzname & "', '" & $sKurzname & "', '" & StringUpper($sPasswd) & "', '" & _
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

EndFunc		;==>_add_SoftM_User

Func _add_Intranet_User()
; *****************************************************************************
; Benutzerdaten im Intranet ablegen
; per Verbindung zu MS-SQL without an ODBC.
; *****************************************************************************
Local $bGefunden = False
Local $sEmail = $sNickName & "@" & $sDomainName

Local $sMySQL = ("SELECT * FROM benutzer WHERE login_ad LIKE '" & $sStrADUser &"'")
Local $sValues = ("'" & $sStrNanchnameOK & "', '" & $sStrVornameLang & "', '" & $nPersNr & "', '" &  $sEmail  & "', '" &  $sStrADUser  & "', '" &  $sPasswd  & "', '" &  $sStrADUseBIG & "', '" &  $sPasswd & "', 'A', 'benutzer', 'on'")

; Only for SQL server
$objConn = ObjCreate("ADODB.Connection")
$objConn.Open("Provider='sqloledb';Data Source='sql1';Initial Catalog='benutzer';User ID='intranet_allgemein';Password='Gt_Fu_57kS';")
$rsCustomers = $objConn.Execute($sMySQL)

With $rsCustomers
    While Not .EOF
        $bGefunden = True
        .MoveNext
    WEnd
    .Close
EndWith

If $bGefunden Then
	MsgBox(64, "User Intranet", "Im Intanet ist User '" & $sStrADUser & "' bereits vorhanden. Angaben sind unverändert")
Else
	$rsCustomers = $objConn.Execute("INSERT INTO benutzer (nachname, vorname, personalnummer, email, login_ad, kennwort_ad, login_as400, kennwort_as400, status, art, status_intranet) VALUES (" & $sValues & ")")
	MsgBox(64, "User Intranet", "Im Intanet ist User '" & $sStrADUser & "' neu hinterlegt")
EndIf

$objConn.Close

EndFunc 	;==>_add_Intranet_User

Func _create_Postfach($sUser)
; *****************************************************************************
; Beta 1
; Create a mailbox for a user.
; *****************************************************************************
;~ 	_AttributeSetzenUnivers($sUser, "mail", $sNickName & "@" & $sDomainName)	; !!! Primäre Mailadresse des Benutzers
;~ 	_AttributeSetzenUnivers($sUser, "mailNickname", $sNickName)					; !!! Exchange Alias, welcher als Basis für Mailadressen/MAPI-Profils dient
;~ ; ExchangeQuotas setzen !!!!! Teste - Alle Angaben werden entfernt nach dem Postfach erstellt wird?
;~ 	_AttributeSetzenUnivers($sUser, "mDBStorageQuota", "503317")
;~ 	_AttributeSetzenUnivers($sUser, "mDBOverQuotaLimit", "524288")
;~ 	_AttributeSetzenUnivers($sUser, "mDBOverHardQuotaLimit", "629146")
;~ 	;_AttributeSetzenUnivers($sUser, "mDBUseDefaults", "FALSE")		; sonst bekommt Postfach Standard-Kontingenteinstellungen
;~ 	_AttributeSetzenUnivers($sUser, "mAPIRecipient", "TRUE")		; bleibt bestehen

	; Exchange Postfach erstellen
	_AD_CreateMailbox($sUser, "Mailbox Store (" & $mailServer & ")")			; DESVR-MAIL01

EndFunc 	;==>_create_Postfach

Func _add_Archiv_User()
; *****************************************************************************
; Pustekuchen
; *****************************************************************************
EndFunc 	;==>_add_Archiv_User

Func _prn_Login_Info()
; *****************************************************************************
; Pustekuchen
; Anmeldedaten an Drucker
; *****************************************************************************
EndFunc 	;==>_prn_Login_Info

; *****************************************************************************
; ab hier alle Eingaben zu OU und dem User kontrollieren und anpassen
; *****************************************************************************
#region ### START Koda GUI section ### Form=
Global $Form1_1 = GUICreate("Quick User Management", 716, 446, 298, 229)
GUICtrlCreateLabel("Überprüfe die :", 16, 8, 223, 17)
GUICtrlCreateLabel("Dein Benutzeranmeldename:", 16, 40, 223, 17)
GUICtrlCreateLabel("Prüfe Personal Nummer:", 16, 80, 215, 17)
Global $IOU = GUICtrlCreateInput($sParentOU, 240, 8, 459, 21)
Global $IUser = GUICtrlCreateInput($sStrADUser, 240, 40, 459, 21)
Global $IPersNr = GUICtrlCreateInput($nPersNr, 240, 72, 459, 21)
Global $Group1 = GUICtrlCreateGroup("Resourcen", 16, 112, 681, 121)
Global $CheckSoftM = GUICtrlCreateCheckbox("SoftM User", 48, 136, 97, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
Global $CheckArchiv = GUICtrlCreateCheckbox("Archiv User", 48, 168, 97, 17)
Global $CheckMail = GUICtrlCreateCheckbox("Mail Empfänger", 344, 136, 109, 20)
GUICtrlSetState(-1, $GUI_CHECKED)
Global $CheckWaage = GUICtrlCreateCheckbox("Waage User", 344, 168, 97, 17)
Global $CheckIntranet = GUICtrlCreateCheckbox("Intranet User", 48, 200, 97, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $ComboDMS = GUICtrlCreateCombo("Rechnungsprüfer", 16, 249, 345, 25)
GUICtrlSetData(-1, "DMS_Arbeitschutz|DMS_Arbeitsvorbereitung|DMS_Baumarkt|DMS_Betriebsrat|DMS_Buchhaltung|DMS_Controlling|DMS_Einkauf|DMS_Export_Tiefbau|DMS_Logistik|DMS_PL|DMS_QM_QS|DMS_Vertrieb_Tiefbau")
Global $ComboSoftmMenu = GUICtrlCreateCombo("SoftM-Menu auswählen", 384, 249, 313, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "IVEROOT|IAVROOT|IEKROOT|ILOROOT|IROOT")
Global $BOK = GUICtrlCreateButton("Benutzer erstellen", 8, 396, 121, 33)
Global $BCancel = GUICtrlCreateButton("Abrechen", 628, 396, 73, 33, $BS_DEFPUSHBUTTON)
GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

While 1
    Global $nMsg = GUIGetMsg()
    Switch $nMsg
        Case $GUI_EVENT_CLOSE, $BCancel
            Exit
        Case $BOK
			; Korrekturen festhalten
            $sOU = GUICtrlRead($IOU)
			$sUser = GUICtrlRead($IUser)
			; $nPersNr, SoftM, SoftM-menu, Intranet, Archiv, Postfach, Waageuser, DMS-Gruppen?
			$nPersNr = GUICtrlRead($IPersNr)

			if BitAnd(GUICtrlRead($CheckSoftM),$GUI_CHECKED) = $GUI_CHECKED then
				$bSoftM = True
				$sMenu = GUICtrlRead($ComboSoftmMenu)
				if BitAnd(GUICtrlRead($CheckWaage),$GUI_CHECKED) = $GUI_CHECKED then
					$bWaage = True
				EndIf
			EndIf

			if BitAnd(GUICtrlRead($CheckIntranet),$GUI_CHECKED) = $GUI_CHECKED then
				_add_Intranet_User()
			EndIf

			if BitAnd(GUICtrlRead($CheckArchiv),$GUI_CHECKED) = $GUI_CHECKED then
				_add_Archiv_User()
			EndIf

			if BitAnd(GUICtrlRead($CheckMail),$GUI_CHECKED) = $GUI_CHECKED then
				$bPostfach = True
				;_create_Postfach()
			EndIf

			$sDMS = GUICtrlRead($ComboDMS)
            ExitLoop
    EndSwitch
WEnd

; *****************************************************************************
; Benutzer erstellen und Parametern setzen
; *****************************************************************************
Global $iValueAD = _AD_CreateUser($sOU, $sUser, $sStrADUserText)
If $iValueAD = 1 Then
    MsgBox(64, "Active Directory Manipulation Ergebnis", "User '" & $sUser & "' in OU '" & $sOU & "' erfolgreich erstellt")
	Sleep(1000)
	_AttributeSetzenUnivers($sUser, "name", $sStrADUserRoh)							; +Strätker, Markus
	_AttributeSetzenUnivers($sUser, "displayName", $sStrADUserRoh)					; Strätker, Markus
	_AttributeSetzenUnivers($sUser, "description", $sDescription)					; == Logistik ==
	_AttributeSetzenUnivers($sUser, "givenName", $sStrVorname)						; Vorname des Anwenders - Markus
	_AttributeSetzenUnivers($sUser, "employeeID", $nPersNr)							; Personal Nummer z.B = 537
	; wenn Waage Logistik
	if $bWaage = True Then
		_AttributeSetzenUnivers($sUser, "employeeNumber", 6000 + $nPersNr)			; Personal Nummer Logistik Waage = 6537
	EndIf
	_AttributeSetzenUnivers($sUser, "userAccountControl", 512)						; 512 - NORMAL_ACCOUNT (Default user)
	_AttributeSetzenUnivers($sUser, "homeDirectory", $sHomeDirectory & $sStrADUser)
	_AttributeSetzenUnivers($sUser, "homeDrive", $sHomeDrive)
	_AttributeSetzenUnivers($sUser, "l", "Hamm")
	;_AttributeSetzenUnivers($sUser, "physicalDeliveryOfficeName", "Büro Halle 3") 	; Büro Halle 3
	_AttributeSetzenUnivers($sUser, "postalCode", "59071")
	_AttributeSetzenUnivers($sUser, "sn", $sStrNanchname)							; Nachname des Anwenders - Strätker
	_AttributeSetzenUnivers($sUser, "st", "NRW")
	_AttributeSetzenUnivers($sUser, "streetAddress", "Siegenbeckstraße 15")
	_AttributeSetzenUnivers($sUser, "telephoneNumber", "02388 3071-")
	_AttributeSetzenUnivers($sUser, "wWWHomePage", "www." & $sDomainName)
	_AttributeSetzenUnivers($sUser, "company", "Funke Kunststoffe GmbH")
	_AttributeSetzenUnivers($sUser, "countryCode", "276")
	_AttributeSetzenUnivers($sUser, "c", "DE")
	_AttributeSetzenUnivers($sUser, "co", "Deutschland")

	; Wenn Postfach benötigt
	If $bPostfach = True Then
		_create_Postfach($sUser)
	EndIf

	; Gruppenmitgliedschaft
	_GruppenMitglied($sUser, "Benutzer")
	_GruppenMitglied($sUser, "vDesktop-FHU-01")
	_GruppenMitglied($sUser, "appCRM")
	;_GruppenMitglied($sUser, "Benutzer SoftM DE")
	If $sDMS <> "Rechnungsprüfer" Then
		_GruppenMitglied($sUser, $sDMS)
		MsgBox(64, "DMS", "User '" & $sOU & "' als Rechnungsprüfer aufgenomen")
	EndIf

	_PasswordSetzen($sUser, $sPasswd)
	_BasisOrdnerCreate($sUser)
	If $bSoftM = True Then
		_GruppenMitglied($sUser, "Benutzer SoftM DE")
		_add_SoftM_User()
	EndIf

ElseIf @error = 1 Then
    MsgBox(64, "Active Directory Manipulation Ergebnis", "User '" & $sUser & "' ist bereits vorhanden")
ElseIf @error = 2 Then
    MsgBox(64, "Active Directory Manipulation Ergebnis", "OU '" & $sOU & "' ist nicht vorhanden")
ElseIf @error = 3 Then
    MsgBox(64, "Active Directory Manipulation Ergebnis", "Werte für CN (e.g. Lastname Firstname) nicht vorhanden")
ElseIf @error = 4 Then
    MsgBox(64, "Active Directory Manipulation Ergebnis", "Wert für $sAD_User nicht vorhanden")
Else
    MsgBox(64, "Active Directory Manipulation Ergebnis", "Return code '" & @error & "' from Active Directory")
EndIf

; *****************************************************************************
; Close Connection to the Active Directory
_AD_Close()

Func MyErrFunc()
; *****************************************************************************
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

Func _Exit()
; *****************************************************************************
	Exit
EndFunc