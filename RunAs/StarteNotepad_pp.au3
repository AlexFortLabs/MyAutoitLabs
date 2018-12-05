#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; Startet Notepad++ mit übergebenem Parameter "C:\windows-version.txt"

If $CmdLine[0] > 0 Then
    If FileExists($CmdLine[1]) Then
        Run(@ProgramFilesDir & '\Notepad++\notepad++.exe "' & $CmdLine[1] & '"')
        Send("^a")
        Send("^a^s^+b")
        sleep (500)
        Send("^!s")
        sleep (500)
        Send("D:\work\NI_WI_MONTHLY_INSTALMENTS_c2base64-7.txt")
        sleep (500)
        Send("{ENTER}")
    EndIf
EndIf