; USER DATA
$s_FilePath = @ScriptDir & '\SampleDB.mdb'
$s_TableName = 'SampleTable'
$s_FieldName1 = 'SampleField1'
$s_FieldName2 = 'SampleField2'
$s_Record = 2
; PROGRAM DATA
$s_Query = ''
$Options = False
$ReadOnly = False
$UserID = 'admin'
$Password = 'Blah'
$Result = 0

If FileExists($s_FilePath) Then FileDelete($s_FilePath)
Sleep(1000)

; CREATE NEW DB
$Auth = ';uid=' & $UserID & ';pwd=' & $Password
$o_object = ObjCreate('DAO.DBEngine.36')
$o_object.CreateDatabase($s_FilePath, ';LANGID=0x0409;CP=1252;COUNTRY=0' &$Auth) 		; ";LANGID=0x0409;CP=1252;COUNTRY=0" for English, German, French, Portuguese, Italian, and Modern Spanish
$AccessRoot = $o_object.OpenDatabase($s_FilePath, $Options, $ReadOnly,$Auth)

; RUN QUERY
$s_Query = 'CREATE TABLE ' & $s_TableName & ' (p_ID AUTOINCREMENT, ' & _
                    $s_FieldName1 & ' Text, ' & _
                    $s_FieldName2 & ' Memo, ' & _
                    'PRIMARY KEY (p_ID))'
$AccessRoot.Execute($s_Query)
$s_Query = 'INSERT INTO ' & $s_TableName & ' (' & _
                    $s_FieldName1 & ', ' & _
                    $s_FieldName2 & _
                    ') VALUES ("First Record", "Blah blah")'
$AccessRoot.Execute($s_Query)
$s_Query = 'INSERT INTO ' & $s_TableName & ' (' & _
                    $s_FieldName1 & ', ' & _
                    $s_FieldName2 & _
                    ') VALUES ("Second record", "Blah Huh")'
$AccessRoot.Execute($s_Query)

; COUNT TABLE
$AccessTableCount = $AccessRoot.TableDefs.Count
For $i = 0 To $AccessTableCount - 1
    ; Dont count system db
    If BitAND($AccessRoot.TableDefs($i).Attributes, -2147483646) = 0 Then
        $Result = $Result + 1
    EndIf
Next
MsgBox(0, 'Number of tables', $Result)
For $i = 0 To $AccessTableCount - 1
    ; CHECK IF TABLE EXISTS
    If $AccessRoot.Tabledefs($i).Name = $s_TableName Then
        MsgBox(0, '', 'Table exists')
    EndIf
Next

; SELECT A TABLE
$AccessTableSelect = $AccessRoot.Tabledefs($s_TableName)

; COUNT FIELD
$AccessFieldCount = $AccessTableSelect.Fields.Count
MsgBox(0, 'Number of fields', $AccessFieldCount)
; CHECK IF FIELD EXISTS
For $i = 0 To $AccessFieldCount - 1
    If $AccessTableSelect.Fields($i).Name = $s_FieldName1 Then
        MsgBox(0, '', 'Field exists')
    EndIf
Next

; COUNT RECORD
$AccessRecordPointer = $AccessRoot.OpenrecordSet($s_TableName)

$AccessRecordPointer.MoveFirst
$AccessRecordCount = $AccessRecordPointer.RecordCount
MsgBox(0, 'Number of records', $AccessRecordCount)
; GET RECORD
$AccessRecordPointer.Move($s_Record - 1)
If $AccessRecordPointer.EOF <> -1 Or $AccessRecordPointer.BOF <> -1 Then
    For $i = 0 To $AccessFieldCount - 1
        ; Result here
        $Result = $AccessRecordPointer.Fields($i).Value
        MsgBox(0, 'Get value from record', 'Column ' & $i & ': ' & $Result)
    Next
    ; EDIT RECORD
    $AccessRecordPointer.Edit
    $AccessRecordPointer.Fields($s_FieldName1).Value = 'Modified item'
    ;...
    If @error <> 0 Then
        $AccessRecordPointer.CancelUpdate
    Else
        $AccessRecordPointer.Update
    EndIf
    ; DELETE RECORD ^^^^^ "Modified item"
    $AccessRecordPointer.Delete
EndIf
; ADD RECORD
$AccessRecordPointer.AddNew
$AccessRecordPointer.Fields($s_FieldName1).Value = 'New item'
$AccessRecordPointer.Fields($s_FieldName2).Value = 'New item'
;...
If @error <> 0 Then
    $AccessRecordPointer.CancelUpdate
Else
    $AccessRecordPointer.Update
EndIf

$AccessRecordPointer.Close

; RUN THE "SELECT" QUERY
$s_Query = 'SELECT * FROM ' & $s_TableName & ' WHERE ' & $s_FieldName1 & ' = "First Record"'
$AccessQueryData = $AccessRoot.OpenRecordSet($s_Query)
For $i = 1 To $AccessQueryData.Fields.Count - 1
    ; Result here
    $Result = $AccessQueryData.Fields($i - 1).Value
    MsgBox(0, 'Run the SELECT query', 'Column ' & $i & ': ' & $Result)
    If @error <> 0 Then ExitLoop
Next

$AccessRoot.Close