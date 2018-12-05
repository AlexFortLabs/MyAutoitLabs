; Demo Script zu _ADODB.au3
;
; Die Daten werden entweder auf einem SQL-Server oder auf einer lokalen MDB gespeichert.
; Hier ist eine einfache UDF für den Umgang mit ADODB.

#include <_ADODB.au3>
#include <Array.au3>   ;Only needed for This Demo Script

$adCurrentProvider = $adProviderMSJET4   ;Microsoft.Jet.OLEDB.4.0
$adCurrentDataSource = @ScriptDir & "\TEMPDB.MDB"

;Establish ADO Connection
_OpenConnection($adCurrentProvider,$adCurrentDataSource)

;Create Table
Dim $arrFields[3]=["Firstname VARCHAR(50)","Lastname VARCHAR(50)","AnotherField VARCHAR(50)"]  ;Init Fields for Table Creation
_CreateTable("TestTable",$arrFields)

;Insert Record
Dim $arrFieldsandValues[3][3]=[["Firstname","Lastname","AnotherField"],["John","Smith","Rulez"],["Dave","Thomas","Rulez"]]   ;Init Fields and Values for Insert
_InsertRecords("TestTable",$arrFieldsandValues)

;Retrieve Records from Table
Dim $arrSelectFields[3]=["Firstname","Lastname","AnotherField"]   ;Init Select Fields for Recordset
$arrRecords = _GetRecords("TestTable",$arrSelectFields)
_ArrayDisplay($arrRecords)

;Retrieve Records from Table with Where Clause
Dim $arrWhereFields[1]=["Firstname = 'Dave'"]
$arrRecords = _GetRecords("TestTable",$arrSelectFields,$arrWhereFields)
_ArrayDisplay($arrRecords)

;Capture Tables and Colums
$tables = _GetTablesAndColumns(1)   ;Param causes display if SYSTEM tables.  Default (no param) hides these tables
_ArrayDisplay($tables)   ;Display Tables
For $x = 1 to UBound($tables)-1   ;Display Table Columns
    _ArrayDisplay($tables[$x][1],"Table Name: " & $tables[$x][0])
Next

;Add Column to Table
_AddColumn("TestTable","Field2","VARCHAR(10)")

;Drop Column from Table
_DropColumn("TestTable","Field2")

;Drop Table
_DropTable("TestTable")

;Drop Database
_DropDatabase()   ;Closes ADO Connection first if using MS.JET.OLEDB.4.0

;Close ADO Connection
;_CloseConnection()