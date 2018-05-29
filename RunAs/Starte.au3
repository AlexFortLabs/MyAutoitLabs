#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; z.B so

;ShellExecute ( "G:\Games\hl2.exe" , "-game cstrike", "G:\Games")

;Run("hl2.exe -game cstrike")
;Run("program.exe", "", @SW_HIDE)

;RunWait(@ProgramFilesDir & '\Program\123.exe /l "' & @ScriptDir & '\file.txt"')

$command = @ProgramFilesDir & "\Mozilla Firefox\firefox-new-tab"
$url = "http://www.funkegruppe.de"
Run($command & $url, "", @SW_MAXIMIZE)
Run($command & $url, "", @SW_MAXIMIZE)

;~ Задержка 5 минут чтобы Imacros успел скрипт отработать
;Sleep ( 300000 )

;~ Закрываем все окна броузера
;While 1
    ;Sleep(100)
    ;If ProcessExists("firefox.exe") Then
		;ProcessClose("firefox.exe")
    ;Else
		;ExitLoop
    ;EndIf
;WEnd