#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <Constants.au3>

$str = ''
$ip = '10.49.231.10'

$VBS = Run(@ComSpec & '/c nbtstat -a' & $ip & ' ' & chr(34) & 'MAC' & chr(34), @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)

StdinWrite($VBS)

While 1
	$line = StdoutRead($VBS)
	If @error Then ExitLoop
	If $line <> " " Then
		$str = $line
	EndIf
Wend

While 1
	$line = StderrRead($VBS)
	If @error Then ExitLoop
		$str = $line
Wend

;MsgBox(0,»",$str)
;MsgBox(0,»»,$str)
MsgBox(0,'',$str)
