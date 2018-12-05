;  start Internet Explorer in Theatermode/Kioskmode

#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <WinAPI.au3>
#include <Constants.au3>
#include <IE.au3>

Dim $hGUI, $oIE, $sURL, $iLeft, $iTop, $iHeight, $iWidth

$sURL = "http://www.heise.de"
$iLeft = 0
$iTop = 0
$iWidth = 1024
$iHeight = 768

$hGUI = GUICreate("Test", $iWidth, $iHeight)

$oIE = IECreatePseudoEmbedded($iLeft, $iTop, $iWidth, $iHeight, $hGUI)
_IENavigate($oIE, $sURL)
GUISetState(@SW_SHOW, $hGUI)

While 1
  $msg = GUIGetMsg()
  If $msg = $GUI_EVENT_CLOSE Then
    _IEQuit($oIE)
    Exit
  EndIf
WEnd

Exit


Func IECreatePseudoEmbedded($i_Left, $i_Top, $i_Width, $i_Height, $h_Parent)

Local $o_IE, $h_HWND

$o_IE = ObjCreate("InternetExplorer.Application")
$o_IE.theatermode = True
$o_IE.fullscreen = True
$o_IE.statusbar = False

$h_HWND = _IEPropertyGet($o_IE, "hwnd")
_WinAPI_SetParent($h_HWND, $h_Parent)
_WinAPI_MoveWindow($h_HWND, $i_Left, $i_Top, $i_Width, $i_Height, False)
_WinAPI_SetWindowLong($h_HWND, $GWL_STYLE, $WS_POPUP + $WS_VISIBLE)

Return $o_IE

EndFunc ;==>IECreatePseudoEmbedded