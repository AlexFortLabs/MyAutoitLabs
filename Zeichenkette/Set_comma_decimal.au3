; set AutoIt up to work with different countries and how they handle comma/decimals

Local $rKey = "HKCU\Control Panel\International"
Local $sThousands = RegRead($rKey, "sThousand"), $sDecimal = RegRead($rKey, "sDecimal")
If $sThousands = -1 Then $sThousands = RegRead($rKey, "sThousand")

MsgBox(0, "Money", _StringAddThousandsSep(123456789/100))


Func _StringAddThousandsSep($sText)
    If Not StringIsInt($sText) And Not StringIsFloat($sText) Then Return 0
    Local $aSplit = StringSplit($sText, "-" & $sDecimal)
    Local $iInt = 1, $iMod
    If Not $aSplit[1] Then
        $aSplit[1] = "-"
        $iInt = 2
    EndIf
    If $aSplit[0] > $iInt Then
        $aSplit[$aSplit[0]] = "." & $aSplit[$aSplit[0]]
    EndIf
    $iMod = Mod(StringLen($aSplit[$iInt]), 3)
    If Not $iMod Then $iMod = 3
    $aSplit[$iInt] = StringRegExpReplace($aSplit[$iInt], '(?<=\d{' & $iMod & '})(\d{3})', $sThousands & '\1')
    For $i = 2 To $aSplit[0]
        $aSplit[1] &= $aSplit[$i]
    Next
    Return $aSplit[1]
EndFunc   ;==>_StringAddThousandsSep