#include <WinAPIEx.au3>
#include <GuiToolbar.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

Opt('MustDeclareVars', 1)

Global $hGUI, $iDummy, $iTB_BtnFirst = Random(10000, 15000, 1)

_Main_1()
Exit

Func _Main_1()
    Local $hToolbar, $a_IcoString[4][3] = [[130, '&User Neu'],[132, '&Postfach'],[134, '&Soft-M'],[136, '&Intranet']], _
            $ah_BitMap[4], $i_ID_TB_Btn, $a_Info, $i_W = 32, $i_H = 32 ; Wenn 16x16, dann wird Funktion _GUICtrlToolbar_SetBitmapSize nicht benötigt

    $hGUI = GUICreate(StringTrimRight(@ScriptName, 4), 400, 300)
    $iDummy = GUICtrlCreateDummy()
    $hToolbar = _GUICtrlToolbar_Create($hGUI)
    _GUICtrlToolbar_SetBitmapSize($hToolbar, $i_W, $i_H)
    For $i = 0 To 3
        $a_IcoString[$i][0] = _WinAPI_ShellExtractIcon('shell32.dll', $a_IcoString[$i][0], $i_W, $i_H)
        $a_Info = _WinAPI_GetIconInfo($a_IcoString[$i][0])
        $ah_BitMap[$i] = _WinAPI_CopyBitmap($a_Info[5])
        $a_IcoString[$i][2] = _GUICtrlToolbar_AddBitmap($hToolbar, 1, 0, $ah_BitMap[$i])
        $a_IcoString[$i][1] = _GUICtrlToolbar_AddString($hToolbar, $a_IcoString[$i][1])
        If $i = 3 Then _GUICtrlToolbar_AddButtonSep($hToolbar)
        _GUICtrlToolbar_AddButton($hToolbar, $iTB_BtnFirst + $i, $a_IcoString[$i][2], $a_IcoString[$i][1])
        _WinAPI_DestroyIcon($a_IcoString[$i][0])
        For $j = 4 To 5
            _WinAPI_DeleteObject($a_Info[$j])
        Next
    Next
    $a_IcoString = 0
    $a_Info = 0
    GUISetState()
    GUIRegisterMsg($WM_COMMAND, 'WM_COMMAND')
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                For $i = 0 To 3
                    _WinAPI_DeleteObject($ah_BitMap[$i])
                Next
                _GUICtrlToolbar_Destroy($hToolbar)
                GUIDelete($hGUI)
                ExitLoop
            Case $iDummy
                $i_ID_TB_Btn = GUICtrlRead($iDummy)
                MsgBox(64, 'Toolbar', 'Button ID: ' & $i_ID_TB_Btn & @LF & 'Button Text: "' & _GUICtrlToolbar_GetButtonText($hToolbar, $i_ID_TB_Btn) & _
                        '"' & @LF & ' has been pressed!', 0, $hGUI)
        EndSwitch
    WEnd
EndFunc   ;==>_Main_1

Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    Local $i_ID

    Switch $hWnd
        Case $hGUI
            $i_ID = BitAND($wParam, 0xFFFF)
            Switch $i_ID
                Case $iTB_BtnFirst To $iTB_BtnFirst + 3
                    GUICtrlSendToDummy($iDummy, $i_ID)
            EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND