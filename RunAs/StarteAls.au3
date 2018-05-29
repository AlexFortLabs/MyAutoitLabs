#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Запуск программы от имени другого пользователя.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
$Username="Administrator"
$Password="stargatesg1"
$Domain="X03"
; Path to the File that must be executed
$Address="\\Serveripadress\netlogon\Scriptbeispiel\ipadresse_auslesen.exe"

RunAs($Username,$Domain,$Password,0,$Address)
