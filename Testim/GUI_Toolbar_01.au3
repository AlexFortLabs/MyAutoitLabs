#include <WinAPIEx.au3>
#include <GuiToolbar.au3>
#include <GUIConstantsEx.au3>

Opt('MustDeclareVars', 1)

_Main()

Func _Main()
    Local $hGUI, $hToolbar, $i_TB_Btn = 1000, $a_IcoString[4][3] = [[130, '&First'],[132, '&Second'],[134, '&Third'],[136, '&Fourth']], _
            $ah_BitMap[4], $a_Info

    $hGUI = GUICreate(StringTrimRight(@ScriptName, 4), 400, 300)
    $hToolbar = _GUICtrlToolbar_Create($hGUI)
    For $i = 0 To 3
        $a_IcoString[$i][0] = _WinAPI_ShellExtractIcon('shell32.dll', $a_IcoString[$i][0], 16, 16)
        $a_Info = _WinAPI_GetIconInfo($a_IcoString[$i][0])
        $ah_BitMap[$i] = _WinAPI_CopyBitmap($a_Info[5])
        $a_IcoString[$i][2] = _GUICtrlToolbar_AddBitmap($hToolbar, 1, 0, $ah_BitMap[$i])
        $a_IcoString[$i][1] = _GUICtrlToolbar_AddString($hToolbar, $a_IcoString[$i][1])
        If $i = 3 Then _GUICtrlToolbar_AddButtonSep($hToolbar)
        _GUICtrlToolbar_AddButton($hToolbar, $i_TB_Btn, $a_IcoString[$i][2], $a_IcoString[$i][1])
        $i_TB_Btn += 1
        _WinAPI_DestroyIcon($a_IcoString[$i][0])
        For $j = 4 To 5
            _WinAPI_DeleteObject($a_Info[$j])
        Next
    Next
    $a_IcoString = 0
    $a_Info = 0
    GUISetState()
    Do
    Until GUIGetMsg() = $GUI_EVENT_CLOSE
    For $i = 0 To 3
        _WinAPI_DeleteObject($ah_BitMap[$i])
    Next
EndFunc   ;==>_Main