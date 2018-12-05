#Region - Constants and Variables
Global $objConnection = ObjCreate("ADODB.Connection")
Global $objRecordSet = ObjCreate("ADODB.Recordset")
Global $errADODB = ObjEvent("AutoIt.Error","_ErrADODB")
Global $adCurrentProvider = 0
Global $adCurrentDataSource = 0

Const $adOpenForwardOnly = 0
Const $adOpenKeyset = 1
Const $adOpenDynamic = 2
Const $adOpenStatic = 3
Const $adLockReadOnly = 1
Const $adLockPessimistic = 2
Const $adLockOptimistic = 3
Const $adLockBatchOptimistic = 4
Const $adProviderSQLOLEDB = "SQLOLEDB"
Const $adProviderMSJET4 = "Microsoft.Jet.OLEDB.4.0"
Const $adProviderMSJET12 = "Microsoft.ACE.OLEDB.12.0"
#EndRegion

#Region - Functions, Subs, Methods
Func _AddColumn($varTableName,$varColumnName,$varColumnType)
    If Not IsObj($objConnection) Then Return -1
    $strAddCol = "ALTER TABLE " & $varTableName & " ADD " & $varColumnName & " " & $varColumnType
    Return $objConnection.Execute($strAddCol)
EndFunc

Func _CloseConnection()
    Return $objConnection.Close
EndFunc

Func _CloseRecordSet()
    Return $objRecordSet.Close
EndFunc

Func _CreateDatabase($varDatabaseName=0)
    If $adCurrentProvider = $adProviderMSJET4 Then
        $objADOXCatalog = ObjCreate("ADOX.Catalog")
        $strProvider = "Provider=" & $adCurrentProvider & ";Data Source=" & $adCurrentDataSource
        $objADOXCatalog.Create($strProvider)
        $objADOXCatalog = 0
        Return 0
    ElseIf $varDatabaseName Then
        $objConnection.Execute("CREATE DATABASE " & $varDatabaseName)
        Return $objConnection.Execute("USE " & $varDatabaseName)
    EndIf
    Return 0
EndFunc

Func _CreateTable($varTableName,$arrFields)
    If Not IsObj($objConnection) Then Return -1
    If Not IsString($varTableName) Then Return -2
    If Not IsArray($arrFields) Then Return -3
    $varFields = ""
    $varFieldCount = Ubound($arrFields)-1
    For $x = 0 to $varFieldCount
        $varFields &= $arrFields[$x]
        If $x < $varFieldCount Then $varFields &= " ,"
    Next
    Return $objConnection.Execute("CREATE TABLE " & $varTableName & "(" & $varFields & ")")
EndFunc

Func _DropColumn($varTableName,$varColumnName)
    If Not IsObj($objConnection) Then Return -1
    $strDropCol = "ALTER TABLE " & $varTableName & " DROP COLUMN " & $varColumnName
    Return $objConnection.Execute($strDropCol)
EndFunc

Func _DropDatabase($varDatabaseName=0)
    If Not IsObj($objConnection) Then Return -1
    If $adCurrentProvider = $adProviderMSJET4 Then
        _CloseConnection()
        If MsgBox(4+16,"Are you sure?","Are you sure you want to delete" & @CRLF & $adCurrentDataSource & " ?" & @CRLF & @CRLF & "This Cannot Be Undone!") = 6 Then
            Return FileDelete($adCurrentDataSource)
        EndIf
    Else
        $objConnection.Execute("USE master")
        Return $objConnection.Execute("DROP DATABASE " & $varDatabaseName)
    EndIf
EndFunc

Func _DropTable($varTableName)
    If Not IsObj($objConnection) Then Return -1
    Return $objConnection.Execute("DROP TABLE " & $varTableName)
EndFunc

Func _GetRecords($varTable,$arrSelectFields,$arrWhereFields=0)
    If Not IsObj($objConnection) Then Return -1
    If Not IsObj($objRecordSet) Then Return -2
    _OpenRecordset($varTable,$arrSelectFields,$arrWhereFields)
    If Not $objRecordSet.RecordCount Or ($objRecordSet.EOF = True) Then Return -3
    Dim $arrRecords
    $arrRecords = $objRecordSet.GetRows()
    _CloseRecordSet()
    Return $arrRecords
EndFunc

Func _GetTablesAndColumns($varSystemTables=0)
    If Not IsObj($objConnection) Then Return -1
    $objADOXCatalog = ObjCreate("ADOX.Catalog")
    $objADOXCatalog.ActiveConnection = $objConnection
    Dim $arrTables[1][2]=[['Table Name','Columns Array']]
    Dim $arrColumns[1][2]=[['Column Name','Column Type']]
    For $objTable In $objADOXCatalog.Tables
        Local $varSkipTable = 0
        If Not $varSystemTables Then
            If StringInstr($objTable.Type,"SYSTEM") or StringInstr($objTable.Name,"MSys")=1 Then $varSkipTable = 1
        EndIf
        If Not $varSkipTable Then
            ReDim $arrTables[UBound($arrTables)+1][2]
            $arrTables[UBound($arrTables)-1][0] = $objTable.Name
            ReDim $arrColumns[1][2]
            For $objColumn in $objTable.Columns
                ReDim $arrColumns[UBound($arrColumns)+1][2]
                $arrColumns[UBound($arrColumns)-1][0] = $objColumn.Name
                $arrColumns[UBound($arrColumns)-1][1] = $objColumn.Type
            Next
            $arrTables[UBound($arrTables)-1][1] = $arrColumns
        EndIf
    Next
    $objADOXCatalog = 0
    Return $arrTables
EndFunc

Func _InsertRecords($varTableName,$arrFieldsAndValues)
    If Not IsObj($objConnection) Then Return -1
    If Not IsArray($arrFieldsAndValues) Then Return -2
    For $y = 1 To UBound($arrFieldsAndValues)-1
        $strInsert = "INSERT INTO " & $varTableName & " ("
        $varFieldCount = UBound($arrFieldsAndValues,2)-1
        For $x = 0 To $varFieldCount
            $strInsert &= $arrFieldsAndValues[0][$x]
            If $x < $varFieldCount Then $strInsert &= ","
        Next
        $strInsert &= ") VALUES ("
        For $x = 0 To $varFieldCount
            $strInsert &= "'" & $arrFieldsAndValues[$y][$x] & "'"
            If $x < $varFieldCount Then $strInsert &= ","
        Next
        $strInsert &= ");"
        $objConnection.Execute($strInsert)
        If $errADODB.number Then
            If Msgbox(4+16+256,"Insert Record Error","Statement failed:" & @CRLF & $strInsert & @CRLF & @CRLF & "Would you like to continue?") <> 6 Then Return -3
        EndIf
    Next
    Return 1
EndFunc

Func _OpenConnection($varProvider,$varDataSource,$varTrusted=0,$varInitalCatalog="",$varUser="",$varPass="")
    If Not IsObj($objConnection) Then Return -1
    $adCurrentDataSource = $varDataSource
    $adCurrentProvider = $varProvider
    If $adCurrentProvider = $adProviderMSJET4 Then
        If Not FileExists($adCurrentDataSource) Then
            If MsgBox(4+16,$adCurrentDataSource & " does not exist.","Would you like to attempt" & @CRLF & "to create it?") = 6 Then
                _CreateDatabase()
            Else
                Return 0
            EndIf
        EndIf
    EndIf
    $strConnect = "Provider=" & $adCurrentProvider & ";Data Source=" & $adCurrentDataSource & ";"
    If $varTrusted Then $strConnect &= "Trusted_Connection=Yes;"
    If $varUser Then $strConnect &= "User ID=" & $varUser & ";"
    If $varPass Then $strConnect &= "Password=" & $varPass & ";"
    $objConnection.Open($strConnect)
    If $varInitalCatalog Then
        Return $objConnection.Execute("USE " & $varInitalCatalog)
    Else
        Return 1
    EndIf
EndFunc

Func _OpenRecordset($varTable,$arrSelectFields,$arrWhereFields=0,$varCursorType=$adOpenForwardOnly,$varLockType=$adLockReadOnly)
    If Not IsObj($objConnection) Then Return -1
    If Not IsObj($objRecordSet) Then Return -2
    $strOpen = "SELECT "
    $varFieldCount = UBound($arrSelectFields)-1
    For $x = 0 to $varFieldCount
        $strOpen &= "[" & $arrSelectFields[$x] & "]"
        If $x < $varFieldCount Then $strOpen &= ", "
    Next
    $strOpen &= " FROM " & $varTable
    If IsArray($arrWhereFields) Then
        $strOpen &= " WHERE "
        $varFieldCount = UBound($arrWhereFields)-1
        For $x = 0 to $varFieldCount
            $strOpen &= $arrWhereFields[$x]
            If $x < $varFieldCount Then $strOpen &= ", "
        Next
    EndIf
    Return $objRecordSet.Open($strOpen,$objConnection,$varCursorType,$varLockType)
EndFunc

Func _ErrADODB()
    Msgbox(0,"ADODB COM Error","We intercepted a COM Error !"      & @CRLF  & @CRLF & _
             "err.description is: "    & @TAB & $errADODB.description    & @CRLF & _
             "err.windescription:"     & @TAB & $errADODB.windescription & @CRLF & _
             "err.number is: "         & @TAB & hex($errADODB.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $errADODB.lastdllerror   & @CRLF & _
             "err.scriptline is: "     & @TAB & $errADODB.scriptline     & @CRLF & _
             "err.source is: "         & @TAB & $errADODB.source         & @CRLF & _
             "err.helpfile is: "       & @TAB & $errADODB.helpfile       & @CRLF & _
             "err.helpcontext is: "    & @TAB & $errADODB.helpcontext _
            )
    Local $err = $errADODB.number
    If $err = 0 Then $err = -1
EndFunc
#EndRegion