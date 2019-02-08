#include <GUIConstantsEx.au3>
#include <TabConstants.au3>

Local $tab, $iCombo, $msg, $Gui

$Gui = GUICreate("- (Tab) - GUI")

GUISetBkColor(0xB4E1D3)
GUISetFont(9, 300)

$tab = GUICtrlCreateTab(10, 10, 380, 200)

GUICtrlCreateTabItem("Path")
GUICtrlSetImage(-1, "shell32.dll", -4, 0)
GUICtrlCreateLabel("Path", 40, 43, 270, 17)
GUICtrlCreateButton("OK", 314, 60, 46, 25)
GUICtrlCreateInput("C:\WINDOWS\system32", 40, 60, 275, 25)

GUICtrlCreateTabItem("Style")
GUICtrlSetImage(-1, "shell32.dll", -6, 0)
GUICtrlSetState(-1, $GUI_SHOW)
GUICtrlCreateLabel("Sel. Style", 20, 54, 250, 17)
$iCombo = GUICtrlCreateCombo("", 20, 70, 310, 120)
GUICtrlSetData(-1, "$GUI_SS_DEFAULT_TAB|$TCS_FIXEDWIDTH|$TCS_FIXEDWIDTH+$TCS_FORCEICONLEFT|$TCS_FIXEDWIDTH+$TCS_FORCELABELLEFT|$TCS_BOTTOM", "$GUI_SS_DEFAULT_TAB")

GUICtrlCreateTabItem("?")
GUICtrlSetImage(-1, "shell32.dll", -222, 0)
GUICtrlCreateLabel("Des", 20, 40, 120, 17)
GUICtrlCreateButton("OK", 300, 150, 70, 30)

GUICtrlCreateTabItem("")

GUICtrlCreateLabel('$TCS_MULTILINE - ' & @CRLF & '$TCS_BUTTONS - ' & @CRLF & '$TCS_FLATBUTTONS+$TCS_BUTTONS - ', 20, 230, 370, 100)

GUISetState()

While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            ExitLoop
        Case $tab
            ; отображает кликнутую вкладку
            WinSetTitle($Gui, "", "- (Tab) - GUI, - " & GUICtrlRead($tab))
        Case $iCombo
            GUICtrlSetStyle($tab, BitOR($GUI_SS_DEFAULT_TAB, Execute(GUICtrlRead($iCombo))))
    EndSwitch
WEnd