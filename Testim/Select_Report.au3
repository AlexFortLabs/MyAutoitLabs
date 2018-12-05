; Select Report

#include <file.au3>
#include <Date.au3>
#include <GUIConstantsEx.au3>
;calculate yesterdays date
$ydate = _DateAdd('d','-1', _NowCalcDate())
$ymonth = StringMid($ydate, 6,2)
$yyear = Stringmid($ydate, 3,2)
$yday = Stringright($ydate, 2)
;MsgBox(0, "Yesterdays day is", $yday)
;MsgBox(0, "Yesterdays month is", $ymonth)
;MsgBox(0, "Yesterdays year is", $yyear)

Local $PathSc = _PathFull(@ScriptFullPath)
;Variables and constants
;window name for the as400 session
$frmInformation1 = GUICreate("Options", 325, 80, 190, 110)
    GUICtrlCreateIcon($PathSc, -1, 275, 10, 32, 32)
    GUISetBkColor(0xffe4b5)
    GUICtrlSetDefColor(0xff1493)
    GUICtrlSetDefBkColor(0xffffff)
    $btnOK1 = GUICtrlCreateButton("Month Start", 250, 48, 70, 25, 0x0001)
    $btnOK2 = GUICtrlCreateButton("Month End", 250, 8, 70, 25, 0x0001)
    $btnOK3 = GUICtrlCreateButton("Medicare Start", 165, 8, 80, 25, 0x0001)
    $btnOK4 = GUICtrlCreateButton("Medicare End", 165, 48, 80, 25, 0x0001)
    $btnOK5 = GUICtrlCreateButton("Print Queue", 90, 48, 70, 25, 0x0001)
    $btnOK6 = GUICtrlCreateButton("About Me", 10, 48, 70, 25, 0x0001)
    GUICtrlSetCursor($btnOK1, 0)
    GUICtrlSetCursor($btnOK2, 0)
    GUICtrlSetCursor($btnOK3, 0)
    GUICtrlSetCursor($btnOK4, 0)
    GUICtrlSetCursor($btnOK5, 0)
    GUICtrlSetCursor($btnOK6, 0)
    $lblInfo1 = GUICtrlCreateLabel("Starstar Starstar Starstar", 10, 10, 150, 20)
    $lblInfofont = GUICtrlSetFont($lblInfo1, 12, 800)
    GUICtrlSetCursor($lblInfo1, 7)
    GUISetState(@SW_SHOW)
While 1
        $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $btnOK1
                $mstart = InputBox("Month Start", "First Day of the Month Closed?", $ymonth & "01" & $yyear, "" )
                Case $GUI_EVENT_CLOSE
            Exit
            Case $btnOK2
            $mend = InputBox("Month End", "Last Day of the Month Closed?", $ymonth & $yday & $yyear, "" )
        Case $btnOK3
            $medstart = InputBox("Medicare Start", "First Medicare Cycle Date?", "100113", "" )
        Case $btnOK4
            $medend = InputBox("Medicare End", "Last Day of the Medicare Cycle?", "093014", "" )
        Case $btnOK5
            $prtqueue = InputBox("Print Queue", "Month End Report Queue?", "MERPTSQ", "" )
        Case $btnOK6
            MsgBox(0,"Starstat Autoit","Autoit Forums")
            EndSwitch
    WEnd
    GUISetState(@SW_HIDE)
    Exit
    $windowname = "Session A - [24 x 80]"
AutoItSetOption("SendKeyDelay", 150)
sleep (10000)
WinActivate($windowname)
WinWaitActive($windowname)
send("{F3 5}")
sleep(1000)
; do monthly report processing
send("1{enter}")
send("11{enter}")
send("4{enter}")
send($prtqueue & "{NUMPADADD}{tab}Y{enter}")
;menu reset
send("{F3 5}")
; Run several reports
send("1{enter}")
send("8{enter}")
; Admissions Register
MsgBox (64, "Step 1", "Admissions Register 1" )
send("9{enter}")
send($mstart & "{NUMPADADD}")
send($mend & "{NUMPADADD}")
send("2{enter}")
send("{tab 2}" & $prtqueue & "{NUMPADADD}{tab}Y{enter}")
SLEEP(1000)
MsgBox (64, "Step 2", "Admissions Register 2" )
send("9{enter}")
send($mstart & "{NUMPADADD}")
send($mend & "{NUMPADADD}")
send("3{enter}")
send("{tab 2}" & $prtqueue & "{NUMPADADD}{tab}Y{enter}")
; Discharge Register
MsgBox (64, "Step 3", "Discharge Register 1" )
send("10{enter}")
send($mstart & "{NUMPADADD}")
send($mend & "{NUMPADADD}")
send("2{enter}")
send("{tab 2}" & $prtqueue & "{NUMPADADD}{tab}Y{enter}")
SLEEP(1000)
