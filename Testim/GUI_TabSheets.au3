#include <GUIConstantsEx.au3>

Global $afTabSheets[4] = [True, False, True, False]

$Gui = GUICreate("Tab",450, 300)
$Tab1 = GUICtrlCreateTab(20, 24, 425, 201)
$TabSheet1 = GUICtrlCreateTabItem("Tab1")
$TabSheet2 = GUICtrlCreateTabItem("Tab2")
$TabSheet3 = GUICtrlCreateTabItem("Tab3")
$TabSheet4 = GUICtrlCreateTabItem("Tab4")
GUISetState()

While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            ExitLoop
        Case $Tab1
            If $afTabSheets[GUICtrlRead($Tab1)] = False Then
                _Lock_Tab()
            EndIf
    EndSwitch
WEnd

Func _Lock_Tab()
    For $i = 0 To 3
        If $afTabSheets[$i] = True Then
            GUICtrlSetState($TabSheet1 + $i, $GUI_SHOW)
            Return
        EndIf
    Next
EndFunc