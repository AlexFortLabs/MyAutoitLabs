; this script will read data from an access database

; By Greencan

;Prerequisites:
; Using ODBC
; You need to create a system DSN (or file DSN) that points to the .mdb that you want to access.

#include <Array.au3>

Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

Opt("TrayIconDebug", 1)        ;0=no info, 1=debug line info
Opt("ExpandEnvStrings", 1)     ;0=don't expand, 1=do expand
Opt("ExpandVarStrings", 1)     ;0=don't expand, 1=do expand
Opt("GUIDataSeparatorChar","|") ;"|" is the default

; ODBC System DSN definition
$DSN="DSN=MyDevices";    MSAccess database as defined in ODBC

; call SQL
$out=getData($DSN)

; display array
; array element 0 will hold the row titles
_ArrayDisplay($out,"Result of query")

Exit
#FUNCTION# ==============================================================
Func getData($DSN)
    ; ODBC
    $adoCon = ObjCreate ("ADODB.Connection")
    $adoCon.Open ($DSN)
    $adoRs = ObjCreate ("ADODB.Recordset") ; Create a Record Set to handles SQL Records - SELECT SQL
    ;$adoRs2 = ObjCreate ("ADODB.Recordset") ; Create a Record Set to handles SQL Records - UPDATE SQL

    ; create the SQL statement
    $adoSQL = "SELECT CODE, SHORT_EXPLANATION, LONG_EXPLANATION FROM dbo_DevicesALL"		; from example_table

    $adoRs.CursorType = 2
    $adoRs.LockType = 3

    ; execute the SQL
    $adoRs.Open($adoSql, $adoCon)

    DIM $Result_Array[1][1]

    With $adoRs
        $dimension = .Fields.Count
        ConsoleWrite($dimension & @cr)
        ReDim $Result_Array[1][$dimension]
        ; Column header
        $Title = ""
        For $i = 0 To .Fields.Count - 1
            $Title = $Title &  .Fields( $i ).Name  & @TAB
            $Result_Array[0][$i] = .Fields( $i ).Name ; set the array elements
        Next
        ConsoleWrite($Title & @CR & "----------------------------------" & @CR)

        ; loop through the records
        $element = 1
        If .RecordCount Then
            While Not .EOF
                $element = $element + 1
                ReDim $Result_Array[$element][$dimension]
                $Item = ""
                For $i = 0 To .Fields.Count - 1
                    $Item = $Item &  .Fields( $i ).Value & @TAB
                    $Result_Array[$element - 1][$i] = .Fields( $i ).Value ; set the array element
                Next
                ConsoleWrite($Item & @CR)

                .MoveNext
            WEnd
        EndIf
    EndWith

    $adoCon.Close ; close connection
    return $Result_Array
EndFunc
#FUNCTION# ==============================================================
; Com Error Handler
Func MyErrFunc()
    dim $oMyRet
    $HexNumber = Hex($oMyError.number, 8)
    $oMyRet[0] = $HexNumber
    $oMyRet[1] = StringStripWS($oMyError.description,3)
    ConsoleWrite("### COM Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
    SetError(1); something to check for when this function returns
    Return
EndFunc ;==>MyErrFunc
#FUNCTION# ==============================================================