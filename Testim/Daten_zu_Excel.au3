;

#include <Excel.au3>

Local $sFilePath1 =  @ScriptDir & "\" & "TestData.xlsx"		; Diese Datei sollte bereits im angegebenen Pfad existieren

ExelEdit($sFilePath1,'1','2','3','4')

; Funktion fügt neue Zelle (nach der letzter vorhandene)
Func ExelEdit($Filename,$ms2,$ms3,$ms4,$ms5)
;~    ;$Filename (z.B "\docs\Journal.xls") - Dateiname (per Default benutzt den Pfad vom Script, deswegen nur Dateiname und (oder) Pfad wie Unterverzeichnis + Dateiname),
;~    ;$ms - Zehlen für

; Datei öffnen
Local $oExcel = _Excel_Open()
Local $sWorkbook = $Filename
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, Default, Default, True)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File " & $sWorkbook & " does not exist")
    Exit
EndIf

; Finde Letzte befühlte Spalte
$nb_columns = $oExcel.ActiveSheet.UsedRange.Columns.Count

; Finde Letzte befühlte Zeile
$nb_rows = $oExcel.ActiveSheet.UsedRange.Rows.Count

; Schreibe Daten in Zeile (Optional)
;$oExcel.Activesheet.Cells($nb_rows+1, 1).Value = "eins"
;$oExcel.Activesheet.Cells($nb_rows+1, 2).Value = "zwei"
;$oExcel.Activesheet.Cells($nb_rows+1, 3).Value = "drei"

; oder aus dem Massiv
Local $myUsernameArray[5] = [$nb_rows,$ms2,$ms3,$ms4,$ms5] ;

For $z = 1 To  UBound($myUsernameArray)
	$oExcel.Activesheet.Cells($nb_rows+1, $z).Value = $myUsernameArray[$z-1]
Next

; Tabelle speichern
;$oWorkbook.Save()
_Excel_BookSave($oWorkbook)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Error saving workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave Example 1", "Workbook has been successfully saved as " & @ScriptDir & $Filename)

; Tabellen schlissen
_Excel_BookClose($oWorkbook, False)
_Excel_Close($oExcel)

EndFunc
