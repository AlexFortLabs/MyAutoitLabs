; Liest aus einer Excel Tabelle Daten AD Accounts herraus

#include <Array.au3>
#include <Excel.au3>
#include <AD.au3>


;;;Datei Öffnen
;soll die Datei unsichtbar geöffnet werden?
$sichtbar = 0
;Dateipfad; wird noch verscriptet
$exceldatei = "D:\Meldung_1kp.xls"
;Handler der geöffneten Datei
$dateivariable = _ExcelBookOpen($exceldatei,$sichtbar)
;zuerst wollen wir die Nutzer auslesen, das Worksheet heißt "Nutzer"
$worksheet = "Nutzer"

;Worksheet aktivieren, falls jemand das falsche Offen hatte und Gespeichert hat
_ExcelSheetActivate($dateivariable, $worksheet)

;Spalten in Array lesen
$anutzer = _ExcelReadSheetToArray($dateivariable,2,1,0,0,False) ;Direction is Vertical


;schliessen
_ExcelBookClose($dateivariable)

;debug infos
_ArrayDisplay($anutzer,"Nutzer")



 #comments-start
 Der 1. Teil ist abgearbeitet, wir haben nun alle Nutzer mit Daten als Array, jetzt kommt die Erstellung im ActiveDirectory

 #comments-End

 ;AD Variablen
$sAD_UserIdParam = "domain\ad"
$sAD_PasswordParam = "password"
$sAD_HostServerParam = "donau.domain.skb"
$sAD_DNSDomainParam = "DC=domain,DC=skb"
$sAD_ConfigurationParam = "DC=donau,DC=domain,DC=skb"
;ou wo der Nutzer angelegt werden soll
$ou = "OU=test,OU=Nutzer,OU=secret,DC=domain,DC=skb"
;AD Verbindung Öffnen
_AD_Open($sAD_UserIdParam, $sAD_PasswordParam, $sAD_DNSDomainParam, $sAD_HostServerParam, $sAD_ConfigurationParam)



;Alle Nutzer aus der Excel Tabelle ins AD eintragen

For $i = 0 To UBound($anutzer) -1   ; $i = Element 1 (Index=0) bis Zeilen (Index= Zeilen -1)

 $zeiger = $i+1
 $personalnummer = $anutzer[$zeiger][1]
 $nachname = $anutzer[$zeiger][2]
 $vorname = $anutzer[$zeiger][3]


;Beispiel einer AD Nutzererstellung in der OU Test
;_AD_CreateUser("OU=test,OU=Nutzer,OU=secret,DC=domain,DC=skb", 11033496, "Hans Maulwurf")
;SCHARF

_AD_CreateUser($ou, String($personalnummer), $vorname & " " & $nachname)


ConsoleWrite(@error & @CRLF)

;Msgbox(0,"",$vorname &  " " & $nachname & $personalnummer)
 Next
;_AD_ModifyAttribute(11033496, "Common-Name","10101010")





; MUSS IMMER SEIN!!!!
_AD_Close()