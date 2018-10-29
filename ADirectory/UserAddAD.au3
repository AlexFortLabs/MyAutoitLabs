#cs ----------------------------------------------------------------------------

 #INDEX# =======================================================================================================================
 Title .........: String
 AutoIt Version : 3.3.14.5
 Language ......: DE
 Description ...: Script enthält Funktionen zu Manipulation in Aktive Directory.
 ===============================================================================================================================

#ce ----------------------------------------------------------------------------

#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <AD.au3>

; *****************************************************************************
#AutoIt3Wrapper_AU3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y

; Bau Verbindung zu Active Directory
_AD_Open()
If @error Then Exit MsgBox(16, "Active Directory Skript", "Unerwartetes Problem in der Funktion _AD_Open. @error = " & @error & ", @extended = " & @extended)

Global $sStrNanchname				; Töws
Global $sStrNanchnameLang
Global $sStrNanchnameOK				; Toews
Global $sStrVorname					; Rudolf
Global $sStrVornameLang
Global $sStrVornameOK				; Rudolf
Global $sStrADUserRoh				; Töws, Rudolf
Global $sNickName					; r.toews
Global $sPasswd = "123456"
Global $sStrADUser = "Test-User"	; ToewsRu
Global $sStrADUserText = "Bianco Melanie"
Global $sDescription = "== Mitarbeiter Vertrieb =="
Global $sStrADUseBIG				; TOEWSRU
Global $sHomeDirectory = "\\funkegruppe.de\userdata\Userlaufwerk\"
Global $sHomeDrive = "P:"
Global $sDomainName = "funkegruppe.de"


; Ermittelten FQDN für aktuellen User
Global $sFQDN = _AD_SamAccountNameToFQDN()

; OU zusammensetzung
	;Global $sOU = StringMid($sFQDN, $iPos + 1)		; NO
	;Global $sOU = "OU=Standort-FR,DC=funkegruppe,DC=de"	;OK
	;Global $sOU = "CN=Users,DC=funkegruppe,DC=de"
	;Global $sParentOU = StringMid($sFQDN, StringInStr($sFQDN, ",OU=") + 1)		; OK
	;Global $sParentOU = "CN=" & $sStrNanchname & "\, " & $sStrVorname & ",CN=Users,DC=funkegruppe,DC=de"
Global	$sParentOU = "CN=Users,DC=funkegruppe,DC=de"

Global $iReply = MsgBox(308, "Active Directory Manipulation", "Dieses Skript erstellt einen neuen Benutzer in der angegebener OU." & @CRLF & @CRLF & _
        "Möchten Sie Änderungen in  Active Directory vornehmen?")
If $iReply <> 6 Then
	Exit
Else
	_UserNameSetzen()
EndIf

; *****************************************************************************

Func _StringReplace($sString)

	For $i In StringSplit("ÃŸ#ss|Ãœ#Ue|Ã„#Ae|Ã–#Oe|Ã¼#ue|Ã¤#ae|Ã¶#oe|&Uuml;#Ue|&uuml;#ue|&ouml;#oe|&Ouml;#Oe|&auml;#ae|&Auml;#Ae|&szlig#ss|ß#ss|Ä#Ae|Ö#Oe|Ü#Ue|ä#ae|ö#oe|ü#ue|Â#A|À#A|Á#A|â#a|á#a|à#a|Ê#E|È#E|É#E|ê#e|é#e|è#e|Î#I|Ì#I|Í#I|î#i|í#i|ì#i|Ô#O|Ò#O|Ó#O|ô#o|ó#o|ò#o|Û#U|Ù#U|Ú#U|û#U|ú#U|ù#U|,#|.#|-#|/#|\#|_#", "|", 2)
		$sString = StringReplace($sString, StringLeft($i, StringInStr($i, "#", 1)-1), StringTrimleft($i, StringInStr($i, "#", 1)), 0, 1)
	Next
	;ConsoleWrite($sString & @CR)

	Return $sString

EndFunc   ;==>_StringReplace

Func _StringKurzen($sString, $nLang)

	Local $sSErgebnis = StringStripWS($sString, $STR_STRIPALL)
	if StringLen($sSErgebnis) > $nLang Then
		$sSErgebnis = StringMid($sSErgebnis, 1, $nLang)   ;  StringMid( "исходный текст", номер первого символа[, кол-во символов])
		;$sSErgebnis = StringTrimRight($sSErgebnis, $nLang)
	EndIf

	Return $sSErgebnis

EndFunc	  ;==>_StringKurzen

Func _UserNameSetzen()
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
	$sStrADUseBIG = StringUpper($sStrADUser)				; TOEWSRU - für Benutzerprofil im SoftM-Sachbearbeiter
	; Display Ergebnis.
    ;MsgBox($MB_SYSTEMMODAL, "Benutzerprofil", $sStrADUseBIG)

; Frage nach passwort!
    $sPasswd = InputBox("Benutzer Kennwort", "Benutzer " & $sStrADUser & " Passwort eingeben." & @CRLF & "Maximal 9 Zeichen lang." & @CRLF & "Ohne Umlauten", "", "*M9")
	; Display Ergebnis.
    ;MsgBox($MB_SYSTEMMODAL, "", $sPasswd)

EndFunc 	;==>_UserNameSetzen

Func _PasswordSetzen($sUser, $sPasswd)
	; Sets or clears the password for a user or computer
	Local $iValueModPasswd = _AD_SetPassword($sUser, $sPasswd)
	If $iValueModPasswd = 1 Then
		;MsgBox(64, "Active Directory Password", "Passwort für '" & $sUser & "' gesetzt")
	ElseIf @error = 1 Then
		MsgBox(64, "Active Directory Password", "User '" & $sUser & "' does not exist")
	Else
		MsgBox(64, "Active Directory Password", "Return code '" & @error & "' from Active Directory")
	EndIf

EndFunc 	;==>_PasswordSetzen

Func _AttributeSetzenUnivers($sUser, $sAttrubut, $sAttrWert)
	; Change attribute
	Local $iValueModDescript = _AD_ModifyAttribute($sUser, $sAttrubut, $sAttrWert)
	If $iValueModDescript = 1 Then
		;MsgBox(64, "Active Directory Attribut", "Attribut '" & $sAttrubut & "' hinterlegt")
	ElseIf @error = 1 Then
		MsgBox(64, "Active Directory Attribut", "User '" & $sUser & "' does not exist")
	Else
		MsgBox(64, "Active Directory Attribut", "Return code '" & @error & "' from Active Directory")
	EndIf

EndFunc		;==>_AttributeSertzenUnivers

Func _GruppenMitglied($sUser, $sGroup)
	; Add user to group
	Local $iValue = _AD_AddUserToGroup($sGroup, $sUser)
	If $iValue = 1 Then
		;MsgBox(64, "Active Directory Gruppenmitglied", "User '" & $sUser & "' erfolgreich hinzugefügt zu '" & $sGroup & "'")
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
	; Ordner Besitzer & Vollzugriff für User (inklusive Unterordner)
	Local $sBasisFolder = $sHomeDirectory & $sUser
	DirCreate($sBasisFolder)
	If FileExists($sBasisFolder) Then
		MsgBox(4096, "", $sBasisFolder & " wurde angelegt.")

		; icacls D:\Beispielordner /grant znil\Beispielgruppe:(CI)(OI)(M)
		Local $sCommand = "icacls " & $sBasisFolder & " /grant " & $sUser & ":(CI)(OI)(F)"
		Local $sCommandOwner = "icacls " & $sBasisFolder & " /setowner " & $sUser & " /T /C" 	; "takeown /U " & $sUser & " /F " & $sBasisFolder & " /R /D Y /SKIPSL"

		Local $iPID = Run(@ComSpec & " /C " & $sCommand, "", @SW_HIDE, 8)
		MsgBox($MB_SYSTEMMODAL, "PID", $sCommand & ": " & $iPID)

		$iPID = Run(@ComSpec & " /C " & $sCommandOwner, "", @SW_HIDE, 8)
		MsgBox($MB_SYSTEMMODAL, "PID", $sCommandOwner & ": " & $iPID)

	EndIf

EndFunc 	;==>_BasisOrdner

; Eingaben zu OU und dem User
#region ### START Koda GUI section ### Form=
Global $Form1 = GUICreate("Active Directory Manipulation", 714, 124)
GUICtrlCreateLabel("Überprüfe die OU (FQDN):", 8, 10, 231, 17)
GUICtrlCreateLabel("und Benutzeranmeldename (samAccountName)", 8, 42, 231, 17)
Global $IOU = GUICtrlCreateInput($sParentOU, 241, 8, 459, 21)
Global $IUser = GUICtrlCreateInput($sStrADUser, 241, 40, 459, 21)
Global $BOK = GUICtrlCreateButton("Benutzer erstellen", 8, 72, 121, 33)
Global $BCancel = GUICtrlCreateButton("Abrechen", 628, 72, 73, 33, BitOR($GUI_SS_DEFAULT_BUTTON, $BS_DEFPUSHBUTTON))
GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

While 1
    Global $nMsg = GUIGetMsg()
    Switch $nMsg
        Case $GUI_EVENT_CLOSE, $BCancel
            Exit
        Case $BOK
            Global $sOU = GUICtrlRead($IOU)
            Global $sUser = GUICtrlRead($IUser)
            ExitLoop
    EndSwitch
WEnd

; Create a new user
Global $iValue = _AD_CreateUser($sOU, $sUser, $sStrADUserText)
If $iValue = 1 Then
    MsgBox(64, "Active Directory Manipulation Ergebnis", "User '" & $sUser & "' in OU '" & $sOU & "' erfolgreich erstellt")
	;Sleep(2000)
	_AttributeSetzenUnivers($sUser, "name", $sStrADUserRoh)							; +Strätker, Markus
	_AttributeSetzenUnivers($sUser, "displayName", $sStrADUserRoh)					; Strätker, Markus
	_AttributeSetzenUnivers($sUser, "description", $sDescription)					; == Logistik ==
	_AttributeSetzenUnivers($sUser, "givenName", $sStrVorname)						; Markus
	_AttributeSetzenUnivers($sUser, "userAccountControl", 512)						; 512 - NORMAL_ACCOUNT (Учетная запись по умолчанию. Обычная активная учетная запись)
	_AttributeSetzenUnivers($sUser, "homeDirectory", $sHomeDirectory & $sStrADUser)
	_AttributeSetzenUnivers($sUser, "homeDrive", $sHomeDrive)
	_AttributeSetzenUnivers($sUser, "l", "Hamm")
	;_AttributeSetzenUnivers($sUser, "mail", $sNickName & "@" & $sDomainName)
	;_AttributeSetzenUnivers($sUser, "mailNickname", $sNickName)
	;_AttributeSetzenUnivers($sUser, "physicalDeliveryOfficeName", "Büro Halle 3") 	; Büro Halle 3
	_AttributeSetzenUnivers($sUser, "postalCode", "59071")
	_AttributeSetzenUnivers($sUser, "sn", $sStrNanchname)							; Strätker
	_AttributeSetzenUnivers($sUser, "st", "NRW")
	_AttributeSetzenUnivers($sUser, "streetAddress", "Siegenbeckstraße 15")
	_AttributeSetzenUnivers($sUser, "telephoneNumber", "02388 3071-")
	_AttributeSetzenUnivers($sUser, "wWWHomePage", "www." & $sDomainName)
	_AttributeSetzenUnivers($sUser, "company", "Funke Kunststoffe GmbH")
	_AttributeSetzenUnivers($sUser, "countryCode", "276")
	_AttributeSetzenUnivers($sUser, "c", "DE")
	_AttributeSetzenUnivers($sUser, "co", "Deutschland")
	; ExchangeQuotas setzen
	_AttributeSetzenUnivers($sUser, "mDBStorageQuota", "503317")
	_AttributeSetzenUnivers($sUser, "mDBOverQuotaLimit", "524288")
	_AttributeSetzenUnivers($sUser, "mDBOverHardQuotaLimit", "629146")
	_AttributeSetzenUnivers($sUser, "mAPIRecipient", "TRUE")
	; Gruppenmitgliedschaft
	_GruppenMitglied($sUser, "Benutzer")
	_GruppenMitglied($sUser, "vDesktop-FHU-01")
	_GruppenMitglied($sUser, "Benutzer SoftM DE")
	_PasswordSetzen($sUser, $sPasswd)
	_BasisOrdnerCreate($sUser)

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


; Close Connection to the Active Directory
_AD_Close()
