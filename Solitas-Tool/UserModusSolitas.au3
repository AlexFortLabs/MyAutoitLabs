#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         A.Fortowski

 Script Function:
	Durchsucht Unterordner nach bestimmter INI-Datei, ermittelt aus dem Abschnitt AS400 den Wert von gblnWinUserModus und schreibt Treffer in eine Textdatei.

#ce ----------------------------------------------------------------------------
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Program Files (x86)\AutoIt3\Icons\au3.ico
#AutoIt3Wrapper_Outfile=UserModusSolitas.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <File.au3>
#include <Array.au3>
#include <MsgBoxConstants.au3>

Local $sPath = '\\funkegruppe.de\userdata\Userprofile\'
Local $sFile1 = 'ISRCMN.INI'
Local $my_Date = 'DATUM:'&@MDAY&":"&@MON&":"&@YEAR&' UHR:'&@HOUR&':'&@MIN&':'&@SEC

; Erstellt eine konstante Variable im lokalen Bereich des Dateipfads der gelesen bzw. in den geschrieben werden soll.
Const $sFilePathErgebnis = FileOpen(@DesktopDir & "\solitas.txt", $FO_APPEND)
    If $sFilePathErgebnis = -1 Then
        MsgBox($MB_SYSTEMMODAL, "Fehler", "Kann solitas.txt nicht öffnen.")
		Exit
    EndIf

FileWrite($sFilePathErgebnis, "---------------------------------" &@CRLF)
FileWrite($sFilePathErgebnis, "+ " &$my_Date &" +" &@CRLF)
FileWrite($sFilePathErgebnis, "---------------------------------" &@CRLF)
FileWrite($sFilePathErgebnis, " " & @CRLF)

$aDirs = _FileListToArray($sPath, '*', 2) ; das Array $aDirs enthält die Namen der Unterverzeichnisse von $sPath
If Not @error Then
    For $i = 1 To $aDirs[0] ; alle gefundenen Unterverzeichnisse durchgehen
		$sFilePath = ($sPath & $aDirs[$i]& '\Anwendungsdaten\Solitas\' & $sFile1)
        If FileExists($sFilePath) Then
			$sRead = IniRead($sFilePath, "AS400", "gblnWinUserModus", "")
			If $sRead > 0 Then
				FileWriteLine($sFilePathErgebnis, $aDirs[$i] & ' ' & $sRead)
				;MsgBox($MB_SYSTEMMODAL, "File gefunden", $aDirs[$i] & ' ' & $sRead)
			EndIf
		EndIf
    Next
EndIf

; Schließt das Handle welches von FileOpen zurückgegeben wurde.
FileClose($sFilePathErgebnis)

;_ArrayDisplay($aDirs,"$aDirs")