#cs ----------------------------------------------------------------------------
	Code-Snippet wie AutoIT mit dem HLLAPI-API auf einem 5250-Emulator funktioniert.
	Wenn Sie keine Ahnung haben, was HLLAPI oder ein 5250-Emulator ist, benötigen Sie dies wahrscheinlich nicht
#ce ----------------------------------------------------------------------------

#include <Array.au3>
$returncode=""
$function=1
$dataString=""
$dataLength=0

; NB: Your 5250 session must be started before this will work

$dll = DllOpen("ehlapi32.dll") ; change this to whatever dll your client provides

$result = connectSession("A") ; Connect to session "A". Usually displayed on the terminal session window title
if $result<>0 Then
    mb("Connection to open 5250 Session Failed (have you opened a session?)")
    Exit
EndIf

$result = getScreen() ; get the entire screen
$nice=format80($result)

$result = getStatusLine() ; get the status line
$nice=$nice & format80($result)

setNowait() ; don't wait during status checks, for the system to become available
$result=getStatus() ; get the status
; setTimewait()

$nice=$nice & @CRLF & "Status: >" & $result & "<"

MsgBox(1,"ret:" & $dll,$nice)

; sendReset()

setInputFieldText("User", "softmopr")
setInputFieldText("Password","softmopr")
;setInputFieldText("Menu","Salmon:)<img src='http://www.autoitscript.com/forum/public/style_emoticons/<#EMO_DIR#>/smile.png' class='bbc_emoticon' alt=':)' />")
; sendCR()

$result=waitDone(1000) ; wait for system to come back in one second

$loc = findField("Benutzer")
$loc = getNextUField($loc)
$text = getField($loc)   ; get the text from a field given the location.
mb(">" & $text & "<")

if $result>0 then sendReset()
; mb("Wait:" & $result)

DllClose($dll)


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
; Begin lots of functions to wrap around the HLLAPI
;
Func hllCall()
    return DllCall($dll, "none", "hllapi", "word*", $function, "str", $dataString, "word*", $dataLength, "word*", $returncode)
EndFunc

; wait for system to come back in a couple of milliseconds. 0-OK otherwise status error number returned
Func waitDone($milliseconds)
    $done=99;
    $begin = TimerInit()
    while ($done=99)
        $res = getStatus()
        if ($res=0) Then $done=0
        sleep(100)
        $dif = TimerDiff($begin)
        if ($dif > $milliseconds) Then $done=$res
    WEnd
    return $done
EndFunc

; convience function to find a label and set the entry field after it with a value. PADs blanks to max size of field
Func setInputFieldText($partialString,$text)
    $loc = findField($partialString)
    $loc = getNextUField($loc)
    $size = getFieldLength($loc)
    ;mb($size)
    setField($loc,StringLeft($text & "                                               ",$size))
EndFunc

; get the length of a field
Func getFieldLength($loc)
    $function=32
    $dataString="T "
    $dataLength=2
    $returncode=$loc
    $res = hllCall()
    return $res[3]
EndFunc

; get the location of the next entry field given a position
Func getNextUField($loc)
    $function=31
    $dataString="NU"
    $dataLength=2
    $returncode=$loc
    $res = hllCall()
    return $res[3]
EndFunc

; Get location of a particular substring in the entire screen
Func findField($partialText)
    ; dump out whole screen first
    $res = getScreen()

    ; find the text in the result
    return StringInStr($res,$partialText)
EndFunc

; set a certain field to a value
Func setField($pos,$text)
    $function=33
    $dataString=$text
    $dataLength=StringLen($text)
    $returncode=$pos
    return hllCall()
EndFunc

; get the text of a field
Func getField($pos)
    $function=34
    $dataString=""
    $dataLength=300
    $returncode=$pos
    $res = hllCall()
    return $res[2]
EndFunc

; get the status of the screen
Func getStatus()
    $function=4
    $res = hllCall()
    return $res[4]
EndFunc

; get the wait to no wait. i.e. calls to wait will immediately return if busy.
Func setNowait()
    $function=9
    $dataString="NWAIT"
    $dataLength=StringLen($dataString)
    return hllCall()
EndFunc

Func setTimewait()
    $function=9
    $dataString="TWAIT"
    $dataLength=StringLen($dataString)
    return hllCall()
EndFunc

; connect to a session
Func connectSession($session)
    $function=1
    $dataString=$session
    $res = hllCall()
    return $res[4]
EndFunc

; get the entire screen in a string
Func getScreen()
    $function=5
    $res = hllCall()
    return $res[2]
EndFunc

; get the status line
Func getStatusLine()
    $function=13
    $dataLength=110
    $res = hllCall()
    return $res[2]
EndFunc

; send a reset
Func sendReset()
    return sendText("@R")
EndFunc

; send a enter
Func sendCR()
    return sendText("@E")
EndFunc

; send text to screen
Func sendText($text)
    $function=3
    $dataString=$text
    $dataLength=StringLen($text)
    return hllCall()
EndFunc

;
; Some generic Utility functions
;
Func format80($string)
    $end="N"
    $count=0
    $answer=""
    while ($end="N")
        $answer = $answer & StringMid($string,($count * 80) + 1,80) & @CRLF
        $count = $count + 1
        if $count=25 Then $end="Y"
    WEnd
    return $answer
EndFunc

Func mb($string) ; quick and dirty message box
    msgbox(1,"a message",$string)
EndFunc