; ***************************************************************
; Example 1 - Send Data from Excel to AS400
; *****************************************************************

#include <Excel.au3>

Local $oExcel = _Excel_Open()

;Local $sWorkbook = @ScriptDir & "\Extras\_Excel1.xls"
Local $sWorkbook = "C:\ListenTXT\forAS400.xls"
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)

;НАХОДИМ ПОСЛЕДНИЙ ЗАПОЛНЕННЫЙ СТОЛБЕЦ
;$nb_columns = $oExcel.ActiveSheet.UsedRange.Columns.Count

; И СТОРОКУ
$nb_rows = $oExcel.ActiveSheet.UsedRange.Rows.Count



For $selle = 1 To $nb_rows 									; 65536
    WinActivate("Microsoft Excel -forAS400.xls") 			; Activate Excel
    $sCellValue = _Excel_RangeRead($oWorkbook, Default, "A" & $selle)	; Read data

	MsgBox(0, "", "Gelesen Wert: " & @CRLF & $sCellValue, 2)

    If $sCellValue = 9999999 Then ExitLoop; Do this to cells are empty not 4547535 (syntax for empty cell??
    Sleep(1000)
    WinActivate("Sitzung A - [24 x 80]"); activate AS400
    Send($sCellValue); Write the data from excel
    Sleep(500)
    Send("{TAB}")
    Sleep(500)
	Send("softmopr")
	Sleep(500)
	Send("{ENTER}")

    MouseClick("Left", 310, 61, 1); Mouseclick AS400 macro
    Sleep(500)
    Send("{ENTER}"); Confirm use of macro
    Sleep(3000); Wait for the macro to run, big sleep
Next
_Excel_BookClose($oExcel) ; And finally we close out
_Excel_Close($oExcel)
Exit
