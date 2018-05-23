#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         A.Fortowski

 Script Function:
	parsit' sajt bez otkrytija brauzera.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <IE.au3>
#include <MsgBoxConstants.au3>

;$oIE = _IECreate("https://www.heise.de/",0,0)
; Seitendarstelung
;$oIE = _IECreate("my.yahoo.com", 1, 1, 0)
; Codedarstellung
$oIE = _IECreate("my.yahoo.com", 0, 0)
$data = (_IEDocReadHTML($oIE))
$data &= ConsoleRead()
MsgBox(0, "", "Code: " & @CRLF & @CRLF & $data)
