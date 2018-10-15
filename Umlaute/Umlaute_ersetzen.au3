#cs ----------------------------------------------------------------------------

 #INDEX# =======================================================================================================================
 Title .........: String
 AutoIt Version : 3.3.14.5
 Language ......: DE
 Description ...: This module contains various functions for manipulating strings.
 ===============================================================================================================================

#ce ----------------------------------------------------------------------------

#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

; *****************************************************************************
#AutoIt3Wrapper_AU3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#include <AD.au3>

; Bau Verbindung zu Active Directory
_AD_Open()
If @error Then Exit MsgBox(16, "Active Directory Skript", "Unerwartetes Problem in der Funktion _AD_Open. @error = " & @error & ", @extended = " & @extended)

Global $sStrNanchname
Global $sStrNanchnameOK
Global $sStrVorname
Global $sStrVornameOK
Global $sPasswd
Global $sStrADUser = "Test-User"
Global $sStrADUserText = "Bianco Melanie"
Global $sStrADUseBIG

; Ermittelten FQDN für aktuellen User
Global $sFQDN = _AD_SamAccountNameToFQDN()

;Global $sOU = StringMid($sFQDN, $iPos + 1)		; NO
;Global $sOU = "OU=Standort-FR,DC=funkegruppe,DC=de"	;OK
;Global $sOU = "CN=Users,DC=funkegruppe,DC=de"

;Global $sParentOU = StringMid($sFQDN, StringInStr($sFQDN, ",OU=") + 1)		; OK
Global $sParentOU
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
	$sStrNanchnameOK = _StringReplace($sStrNanchname)
	$sStrNanchnameOK = _StringKurzen($sStrNanchnameOK, 8)
	MsgBox($MB_SYSTEMMODAL, "", $sStrNanchname & " ist umgewandelt in:" & @CRLF & @CRLF & $sStrNanchnameOK)

; Frage nach Vorname!
	$sStrVorname = InputBox("Benutzer Vorname", "Bitte Vorname eingeben."  & @CRLF & " z.B : Karl", "", "")
	$sStrVornameOK = _StringReplace($sStrVorname)
	$sStrVornameOK = _StringKurzen($sStrVornameOK, 2)
	MsgBox($MB_SYSTEMMODAL, "", $sStrVorname & " ist umgewandelt in:" & @CRLF & @CRLF & $sStrVornameOK)

; AD-Benutzername Zusammensetzung
	$sStrADUser = $sStrNanchnameOK & $sStrVornameOK		; für AD
	$sStrADUserText = $sStrNanchname & "\, " & $sStrVorname

	$sStrADUseBIG = StringUpper($sStrADUser)				; für Benutzerprofil im SoftM-Sachbearbeiter
; Display Ergebnis.
    MsgBox($MB_SYSTEMMODAL, "Benutzerprofil", $sStrADUseBIG)

; OU zusammensetzung
	;$sParentOU = "CN=" & $sStrNanchname & "\, " & $sStrVorname & ",CN=Users,DC=funkegruppe,DC=de"
	$sParentOU = "CN=Users,DC=funkegruppe,DC=de"

; Frage nach passwort!
    $sPasswd = InputBox("Benutzer Kennwort", "Bitte für Benutzer " & $sStrADUser & " Passwort eingeben." & @CRLF & "Maximal 9 Zeichen lang." & @CRLF & "Ohne Umlauten", "", "*M9")
; Display Ergebnis.
    MsgBox($MB_SYSTEMMODAL, "", $sPasswd)

EndFunc 	;==>_UserNameSetzen

;To search for a user name (full or partial) you can use something like:
;	Global $asUser = StringLeft(@UserName, 1)
;	Global $iResult = _AD_GetObjectsInOU($asUser, $sOU, "(&(objectCategory=user)(name="*loet*))", 2, "department,cn,distinguishedName,sAMAccountName")
;	MsgBox(64, "Active Directory Functions - Example 1", "This example returned " & $iResult & " records")
;This will return department, full name, FQDN and sAMAccountname for all users in the specified OU that have "water" somewhere in the field "name".

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
