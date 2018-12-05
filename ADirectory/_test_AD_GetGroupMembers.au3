#include <ad.au3>
#include <array.au3>

Global $OU = "OU=admins,DC=funkegruppe,DC=de"		; hier nacharbeiten wegen $aGroups ERROR
Global $aGroups, $aGroupMembers

_AD_Open()
$aGroups = _AD_GetObjectsInOU($OU, "(objectcategory=group)", 1)
ConsoleWrite(@error & @CRLF)
ConsoleWrite("--- Found " & $aGroups[0] & " groups to process" & @CRLF)

;alphabetically sort array
_ArraySort($aGroups, 0, 1)

For $i = 1 To $aGroups[0]
    ;do stuff against each
    ;get group group membership
    $aGroupMembers = _AD_GetGroupMemberOf($aGroups[$i])
    ConsoleWrite("----- found: " & $aGroupMembers[0] & " groups" & @CRLF)
    ;do stuff against each
    For $d = 1 To $aGroupMembers[0]
        ConsoleWrite("------ check:" & $aGroupMembers[$d] & @CRLF)
    Next
Next
_AD_Close()