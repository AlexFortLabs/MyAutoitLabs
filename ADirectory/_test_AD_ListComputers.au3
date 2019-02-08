;  script that list computer accounts on Computers OU in active directory and display for each computer its full name , operating system

;#include <ADfunctions.au3>
#include <AD.au3>

;~ $strDNSDomain = "funkegruppe.de"

;~ Global $aComputers
;~ $sOU = $strDNSDomain
;~ _ADGetObjectsInOU($aComputers,$sOU,"(objectclass=computer)",2,"name,operatingSystem")
;~ _ArrayDisplay($aComputers)

; CN=DEPCD012,OU=Desktops,OU=Computers,OU=Standort-DE,DC=funkegruppe,DC=de

Global $aComputers
Global $sOU = "OU=Computers,OU=Standort-DE,DC=funkegruppe,DC=de"

_AD_Open()
$aComputers = _AD_GetObjectsInOU($sOU, "(objectclass=computer)", 2, "name,samaccountname,description")
_AD_Close()
_ArrayDisplay($aComputers)


