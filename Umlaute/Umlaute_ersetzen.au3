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

; Frage nach Nachname!
	Local $sStrNanchname = InputBox("Benutzer Nachname", "Bitte Nachname eingeben." & @CRLF & " z.B : Mustermann", "", "")
	Local $sStrNanchnameOK = _StringReplace($sStrNanchname)
	$sStrNanchnameOK = _StringKurzen($sStrNanchnameOK, 8)
	MsgBox($MB_SYSTEMMODAL, "", $sStrNanchname & " ist umgewandelt in:" & @CRLF & @CRLF & $sStrNanchnameOK)

; Frage nach Vorname!
	Local $sStrVorname = InputBox("Benutzer Vorname", "Bitte Vorname eingeben."  & @CRLF & " z.B : Karl", "", "")
	Local $sStrVornameOK = _StringReplace($sStrVorname)
	$sStrVornameOK = _StringKurzen($sStrVornameOK, 2)
	MsgBox($MB_SYSTEMMODAL, "", $sStrVorname & " ist umgewandelt in:" & @CRLF & @CRLF & $sStrVornameOK)

; AD-Benutzername Zusammensetzung
	Local $sStrADUser = $sStrNanchnameOK & $sStrVornameOK		; für AD
	;Local $sStrADUser = StringTrimRight($sStrNanchnameOK, 7) & StringTrimRight($sStrVornameOK, 2)
	Local $sStrADUseBIG = StringUpper($sStrADUser)				; für Benutzerprofil im SoftM-Sachbearbeiter
	; Display Ergebnis.
    MsgBox($MB_SYSTEMMODAL, "Benutzerprofil", $sStrADUseBIG)

; Frage nach passwort!
    Local $sPasswd = InputBox("Benutzer Kennwort", "Bitte für Benutzer " & $sStrADUser & " Passwort eingeben." & @CRLF & "Maximal 9 Zeichen lang." & @CRLF & "Ohne Umlauten", "", "*M9")
; Display Ergebnis.
    MsgBox($MB_SYSTEMMODAL, "", $sPasswd)
