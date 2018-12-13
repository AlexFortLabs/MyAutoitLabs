#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Program Files (x86)\AutoIt3\Icons\au3.ico
#AutoIt3Wrapper_Outfile=InfoStoreIni.exe
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Add_Constants=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****



#Include <Process.au3>
#Include <File.au3>
#include <MsgBoxConstants.au3>

AutoItSetOption("MustDeclareVars", 1)


Global $cINFOSTOREINI 		= "\\funkegruppe.de\userdata\Userprofile\"		;Hier kommen die abgeschlossenen Tasks rein
main()

Func Main()
Local $ret
Local $sEnvVar = EnvGet("USERNAME")
Local $sPfadEnd = "\Anwendungsdaten\Solitas\ISRCMN.INI"
Local $sPfadINI = $cINFOSTOREINI & $sEnvVar & $sPfadEnd
Local $sPfadDig1 = $cINFOSTOREINI & $sEnvVar & "\Anwendungsdaten\Solitas\SOLDSProvider.ini"
Local $sPfadDig2 = $cINFOSTOREINI & $sEnvVar & "\Anwendungsdaten\Solitas\SOLDSConnectorSet.ini"


if fileexists( $sPfadINI) then

	$ret = 	IniRead ( $sPfadINI, "AS400", "APPCSystem", "default" )

	 ; Display the valst4096ue returned by IniRead. Since there is no key stored the value will be the 'Default Value' passed to IniRead.
    MsgBox($MB_SYSTEMMODAL, "Info", "Die Variabl %ret% hat den Wert: " & @CRLF & @CRLF & $ret)

	$ret = 	IniWrite ( $sPfadINI, "Browser", "DefExpPath", "P:\ISR\Save")
	$ret = 	IniWrite ( $sPfadINI, "Browser", "DefExtArchPath", "P:\ISR\Archiv")
	$ret = 	IniWrite ( $sPfadINI, "Browser", "DefCheckInOutPath", "P:\ISR\Import")
	$ret = 	IniWrite ( $sPfadINI, "Browser", "EviorementLogPath", "P:\ISR\temp")

	$ret = 	IniWrite ( $sPfadINI, "AS400", "APPCSystem", "FunkeI5")
	$ret = 	IniWrite ( $sPfadINI, "AS400", "OverlayPath", "P:\ISR\Overlay\")
	$ret = 	IniWrite ( $sPfadINI, "AS400", "gblnWinUserModus", "1")

	$ret = 	IniWrite ( $sPfadINI, "ImageView", "EviorementLogPath", "P:\ISR\temp")

	$ret = 	IniWrite ( $sPfadINI, "DigSign", "DSBaseSetINI", $sPfadDig1)
	$ret = 	IniWrite ( $sPfadINI, "DigSign", "DSConSetINI", $sPfadDig2)

	MsgBox($MB_SYSTEMMODAL, "Info", "Die Variabl %sTest% hat den Wert: " & @CRLF & @CRLF & $sPfadINI)


Endif


;Example()


EndFunc





Func Example()
    ; Retrieve the value of the environment variable %APPDATA%.
    ; When you assign or retrieve an envorinment variable you do so minus the percentage signs (%).
    Local $sEnvVar = EnvGet("USERNAME")
	Local $sPfad = "C:\Users\"
	Local $sPfadEnd = "\AppData\Local\Temp\"
	Local $sPfadINI = $sPfad & $sEnvVar & $sPfadEnd
    ; Display the value of the environment variable %APPDATA%.
    MsgBox($MB_SYSTEMMODAL, "Info", "Die environment Variabl %APPDATA% hat den Wert: " & @CRLF & @CRLF & $sPfadINI ); This returns the same value as the macro @AppDataDir does.
EndFunc   ;==>Example
