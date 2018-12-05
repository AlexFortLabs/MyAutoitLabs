; This should get you connected to an DB2 / AS 400 without an ODBC.

; Initialize COM error handler
$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")

$sqlCon = ObjCreate("ADODB.Connection")

$sqlCon.Mode = 16  ; shared
$sqlCon.CursorLocation = 3 ; client side cursor


$sqlCon.Open ("Driver={DB2}; IP=192.168.0.5; Port=[port_number]; Database=[database_name]; UID=SOFTMOPR; PWD=[password]")
If @error Then
    MsgBox(0, "ERROR", "Failed to connect to the database")
    Exit
EndIf

; See also Catalog "ADOX Catalog Example.au3"

$sqlRs = ObjCreate("ADODB.Recordset")
If Not @error Then
    $sqlRs.open ("select * from Table", $sqlCon)
    If Not @error Then
        ;Loop until the end of file
        While Not $sqlRs.EOF
            ;Retrieve data from the following fields
            $OptionName = $sqlRs.Fields ('name' ).Value
            $OptionVal = $sqlRs.Fields ('value' ).Value
            MsgBox(0, "Record Found", "Name:  " & $OptionName & @CRLF & "Value:  " & $OptionVal)
            $sqlRs.FIELDS('"' & $OptionName & '"') = ".F." ; ADDED THIS LINE
           ; $sqlRs.Update  ; ADDED THIS LINE
            $sqlRs.MoveNext
        WEnd
        $sqlRs.close
    EndIf
EndIf

Func MyErrFunc()
  $HexNumber=hex($oMyError.number,8)
  Msgbox(0,"COM Test","We intercepted a COM Error !"       & @CRLF  & @CRLF & _
             "err.description is: "    & @TAB & $oMyError.description    & @CRLF & _
             "err.windescription:"     & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "         & @TAB & $HexNumber              & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "     & @TAB & $oMyError.scriptline     & @CRLF & _
             "err.source is: "         & @TAB & $oMyError.source         & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile       & @CRLF & _
             "err.helpcontext is: "    & @TAB & $oMyError.helpcontext _
            )
  SetError(1)  ; to check for after this function returns
Endfunc