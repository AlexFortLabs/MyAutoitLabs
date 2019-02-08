#include <GUIConstants.au3>
#include <Array.au3>

#Include <GuiEdit.au3>
#Include <GuiStatusBar.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

Opt("GUIOnEventMode", 1)
Global  $MAIN, $CMD_WINDOW
Global  $IP_CONFIG, $OTHER, $BUTT_CLOSE
#Region ### START Koda GUI section ### Form=
$MAIN = GUICreate("CMD FUNCTIONS", 623, 449, 192, 114)
$CMD_WINDOW = GUICtrlCreateEdit("", 10, 10, 600, 289, BitOR($ES_AUTOVSCROLL,$ES_AUTOHSCROLL,$ES_READONLY,$WS_VSCROLL))
    GUICtrlSetFont($CMD_WINDOW, 12, 800, 0, "Times New Roman")
    GUICtrlSetColor($CMD_WINDOW, 0xFFFFFF)
    GUICtrlSetBkColor($CMD_WINDOW, 0x000000)
$IP_CONFIG = GUICtrlCreateButton("IP_CONFIG", 10, 310, 75, 25)
    GUICtrlSetOnEvent($IP_CONFIG, "_IP_CONFIGClick")
$OTHER = GUICtrlCreateButton("OTHER", 95, 310, 75, 25)
    GUICtrlSetOnEvent($OTHER, "_OTHERClick")
$BUTT_CLOSE = GUICtrlCreateButton("EXIT", 535, 310, 75, 25)
    GUICtrlSetOnEvent($BUTT_CLOSE, "_ExitNow")

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
While 1
    Sleep(100)
WEnd
Func _ExitNow()
    Exit
EndFunc
Func _IP_CONFIGClick()
    Local $foo = Run(@ComSpec & " /c " & "IPCONFIG /ALL",@SystemDir,@SW_HIDE,$STDOUT_CHILD)
    Local $line
    While 1
        $line = StdoutRead($foo)
        If @error Then ExitLoop
        If Not $line = "" Then GUICtrlSetData($CMD_WINDOW,$line)
    WEnd
EndFunc
Func _OTHERClick()
    GUICtrlSetData($CMD_WINDOW,"Create a CMD Function to be ran here")
EndFunc
