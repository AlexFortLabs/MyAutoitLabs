#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <ColorConstants.au3>
#include <FontConstants.au3>
#include <file.au3>


AutoItSetOption("MustDeclareVars", 1)


const $cConnect01       = "Provider=IBMDA400;Data Source=FUNKEI5;User Id=QSECOFR ;Password=QSECOFR;"
Const $cAdOpenStatic = 3
Const $cAdLockOptimistic = 3
Const $cAdCmdStoredProc = 4
Const $cAdVarChar = 200
Const $cAdInteger = 3

Const $cAdParamInput = 1
Const $cAdParamOutput = 2


Local $sJobCmd = ""


GrGetSoftSBS()




Func GrGetSoftSBS( )
Local $objConn
Local $objRs
Local $i

Local $sql1 = "SELECT * FROM SMKDIFT.XSB00 WHERE SBSBNR = 200"

    ; Instantiate objects
    $objConn = ObjCreate("ADODB.Connection")
    $objConn.Open($cCONNECT01)

   Msgbox( 4096, "Test", @error)


	$objRs = ObjCreate("ADODB.Recordset")

	$objRs.Open ($sql1, $objConn, $cAdOpenStatic, $cAdLockOptimistic)

	If @error Then
		 Msgbox ( 4096, "MyToolBox:", "Database connection problem!",10)
		Return False
	endIf

	if $objRs.eof then
		Msgbox ( 4096, "Message","TEST ONLY" ,10)
		Return False
	EndIf

	$objRs.MoveFirst()		; We only do the first Job found
      Msgbox( 4096, "Test", $objRs.Fields("SBNAME").value)

	$objRs.close
	$objConn.Close

	Return True

EndFunc




