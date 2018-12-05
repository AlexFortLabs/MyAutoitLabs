; Multiwindows from AS400

$hwnd_EmuA = WinGetHandle("A - [24x80]", "")
$hwnd_EmuB = WinGetHandle("B - [24x80]", "")
$hwnd_EmuC = WinGetHandle("C - [24x80]", "")

;{stuff goes here}

WinActivate($hwnd_EmuA)
If Not WinWaitActive($hwnd_EmuA, "", 30) Then
   ConsoleWrite("Error:  Timeout waiting for Emulator A to activate." & @CRLF)
   Exit(101)
EndIf