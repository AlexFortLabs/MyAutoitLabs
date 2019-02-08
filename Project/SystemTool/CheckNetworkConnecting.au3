; #INDEX# ========================================================================
; Title .........: FehlerErmittlung.au3
; AutoIt Version : 0.1
; Language ......: Deutsch
; Description ...: Automatische Ermittlung von Netzwerk-Verbindungsproblemen
; Author ........: Alexander Fortowski
; ================================================================================

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=D:\temp\favicon.ico
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <InetConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>

#include <Inet.au3>		;$External=_GetIP()


AutoItSetOption ( "WinTitleMatchMode", 2 )
AutoItSetOption ( "WinDetectHiddenText", 1 )
AutoItSetOption ( "SendAttachMode", 1 )

; open free desktop

WinMinimizeAll ( )

$sBericht = ("Zusammenfassung zum Vorgang" & @CRLF & @CRLF)
$nBerichtID = 64

Sleep ( '1000' )
;MsgBox ( 64, 'Network Status', @IPAddress1 & @CR & @IPAddress2 & @CR & @IPAddress3 & @CR & @IPAddress4, 100 )

; Adresse per DHCP erneuern
;Run ( @ComSpec & " /c ipconfig /renew", @SW_HIDE)
;RunWait ( @ComSpec & " /c ipconfig /renew")
$MyCommand = 'ipconfig /renew && ping www.google.com && ping www.stackoverflow.com'
RunWait(@ComSpec & " /c " & $MyCommand, @SystemDir, @SW_Show)
MsgBox( 64, 'Network Status nach Update',  @IPAddress1 & @CR & @IPAddress2 & @CR & @IPAddress3 & @CR & @IPAddress4, 100 )

If StringLeft( @IPAddress1, 3 ) == '127' Then
	MsgBox ( 48, 'Network Status', 'Fehlerhafter Netzwerkadapter oder Kabel nicht angeschlossen', 100 )

Else
	If StringLeft( @IPAddress1, 3 ) == '169' Then
		MsgBox ( 48, 'Network Status', 'DHCP-Server nicht verfügbar oder Netzwerkkabel nicht angeschlossen', 100 )

Else
	If StringLeft( @IPAddress1, 8 ) == '192.168.' Or StringLeft( @IPAddress1, 7 ) == '172.16.' Or StringLeft( @IPAddress1, 6 ) == '10.49.' Then			; Private Netwerkbereiche Class ABC = '192.168.0/7.' OR '172.16.' OR '10.49.230/231.'
		;MsgBox ( 0, 'Network Status', 'Unzulässiger Adressbereich', 100 )
;------ address OK

		$ping = Ping("intranet.funkegruppe.de")
		If Not @error Then
			$sBericht &= ('Intranet-Server ist erreichbar' & @CR)
		Else
			$sBericht &= ('Intranet-Server antworten nicht' & @CR)
			$nBerichtID = 48
		EndIf

		;Local Const $NETWORK_ALIVE_WAN = 0x00000002
		;$flag = $NETWORK_ALIVE_WAN
		;$aRet = DllCall("sensapi.dll", "int", "IsNetworkAlive", "int*", 0)
		$aRet = _GetIP()
		;If $aRet[1] = 0 Then
		;If BitAND($aRet[1], $NETWORK_ALIVE_WAN) Then
		if $aRet = -1 Then
			$sBericht &= ('Internet nicht verfügbar' & @CR)
			$nBerichtID = 48
		Else
			$sBericht &= ('Internetzugang erlaubt' & @CR & 'Meine Externe ip Addresse ist: ' & $aRet & @CR)
			;$sBericht = ($sBericht & 'Internetzugang erlaubt' & @CR & 'Meine Externe ip Addresse ist: ' & _GetIp() & @CR)
		EndIf

;------
	Else
		$nBerichtID = 48
		MsgBox ( $nBerichtID, 'Network Status', 'Unbehandeltes Adressbereich', 100 )
	EndIf

EndIf

EndIf

$sBericht &= ( 'Network Status nach Update: ' & @IPAddress1 & @CR & 'PC Name: ' & @ComputerName & @CR & 'User: ' & @UserName & @CR & 'OS Version: ' & @OSVersion & @CR & 'Build: ' & @OSBuild & @CR & 'Service Pack: ' & @OSServicePack)
MsgBox($nBerichtID, "Network Check", $sBericht, 100)


Exit