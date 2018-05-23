#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Alexander Fortowski

 Script Function:
	VNC2Displays AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <File.au3>
#include <MsgBoxConstants.au3>

;#RequireAdmin

Opt("TrayIconHide", 1)

Func SetzeWert($sFilePath)

	;MsgBox($MB_SYSTEMMODAL, "", "Der Pfad lautet: " & $sFilePath)

    ; Schreibt den Wert von '1' in den Schlüssel 'secondary' und in die Sektion 'admin'.
    IniWrite($sFilePath, "admin", "secondary", 1)

    ; Liest in der ini-Datei den Wert von 'secondary' in der Sektion 'admin'.
    ;Local $sRead = IniRead($sFilePath, "admin", "secondary", "Default Value")

    ; Zeigt den Wert der von IniRead zurückgegeben wurde.
    ;MsgBox($MB_SYSTEMMODAL, "", "Der Wert von 'secondary' in der Sektion 'admin' lautet: " & $sRead)

 EndFunc

; Erstellt eine konstante Variable im lokalen Bereich des Dateipfades, welcher zum lesen bzw. schreiben verwendet wird.
;$ini_loc = 'C:\Program Files\uvnc bvba\UltraVNC\ultravnc.ini'			;W10
$ini_loc = 'C:\Program Files\UltraVNC\ultravnc.ini'						;W7
;$ini_loc = 'C:\Temp\ultravnc.ini'

$check = FileExists($ini_loc)
If $check = 1 then
	  SetzeWert($ini_loc)
Else
	  sgBox(4096, "Achtung", "Pfad zu Datei nicht gefunden!", 3)
EndIf