; Dynamisches Fensteraufbau für ja nach Anzahl Radio
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Array.au3>

Local $sPath_ini = @HomeDrive & "\Example.ini"
Local $iLeft = @DesktopWidth/2-115
Local $iTop = 10
Local $iLeftR =10
Local $iTopR = 35
Local $sRead = IniRead($sPath_ini, "Section", "Key", "NotFound")
Local $iRadio = StringSplit($sRead, ",")
; Ja nach dem wie viele Radio benötigt werden, dank Array können wir dynamisch darstellen
Local $iRadio = [5,"1 radio","2 radio","3 radio","4 Punkt","5 Punkt"]

Local $Window = GUICreate("Pass", 230, 174,$iLeft, $iTop)
Local $B_Show = GUICtrlCreateButton("Zeigen", 152, 144, 73, 17)
GUICtrlSetState(-1, $GUI_DISABLE)

Local $Group1 = GUICtrlCreateGroup("", 5, 28,220,90)
Local $count = $iRadio[0]
Local $aItem = []
Local $iItem
For $i = 1 To $count
    $iItem = GUICtrlCreateRadio($iRadio[$i], $iLeftR, $iTopR)
    _ArrayAdd($aItem,$iItem)
    $iTopR += 20
	If $iTopR > 100 Then
        $iTopR = 35
        $iLeftR += 75
    EndIf
Next
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUISetState(@SW_SHOW)

While 1
    $nMsg = GUIGetMsg()
    Switch $nMsg
    Case $GUI_EVENT_CLOSE; $nCancel
            Exit
    Case $aItem[1] to $aItem[$count]
           $vaar =GUICtrlRead($nMsg,1)
           GUICtrlSetState($B_Show,$GUI_ENABLE)
                      MsgBox(0,'',$vaar)
    EndSwitch
WEnd
