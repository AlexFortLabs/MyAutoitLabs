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

Sleep ( '1000' )
;MsgBox ( 64, 'Network Status', @IPAddress1 & @CR & @IPAddress2 & @CR & @IPAddress3 & @CR & @IPAddress4, 100 )

; Adresse per DHCP erneuern
;RunWait ( @ComSpec & " /c ipconfig /renew" )
MsgBox( 64, 'Network Status nach Update',  @IPAddress1 & @CR & @IPAddress2 & @CR & @IPAddress3 & @CR & @IPAddress4, 100 )

If StringLeft( @IPAddress1, 3 ) == '127' Then
	MsgBox ( 48, 'Network Status', 'Fehlerhafter Netzwerkadapter oder Kabel nicht angeschlossen', 100 )

Else
	If StringLeft( @IPAddress1, 3 ) == '169' Then
		MsgBox ( 48, 'Network Status', 'Server nicht verfügbar oder Netzwerkkabel nicht angeschlossen', 100 )

Else
	If StringLeft( @IPAddress1, 8 ) == '192.168.' Then			; '192.168.0.' OR '10.49.231.'
		;MsgBox ( 0, 'Network Status', 'Unzulässiger Adressbereich', 100 )
;------ address OK
		; Variant 1
		;$inByteSize = InetGet('http://intranet.funkegruppe.de/', @MyDocumentsDir & '\in_tmp.html', $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
		;MsgBox(0, "", "The total download size: " & $inByteSize)
		;FileDelete( @MyDocumentsDir & '\tmp_in.html' )

		; Variant 2
		$ping = Ping("intranet.funkegruppe.de")
		;If $inByteSize > 0 Then
		If Not @error Then
			; Variant 1
			;$exByteSize = InetGet('http://www.google.com/', @MyDocumentsDir & '\ex_tmp.html', $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
			;FileDelete( @MyDocumentsDir & '\tmp_ex.html' )
			;MsgBox(0, "", "The total download size: " & $exByteSize)
			; Variant 2
			$aRet = DllCall("sensapi.dll", "int", "IsNetworkAlive", "int*", 0)

			;If $exByteSize > 0 Then
			If $aRet[1] = 0 Then
				MsgBox (64, 'Network Status', 'Intranet-Server ist in Ordnung.' & @CR & 'Internetzugang erlaubt' & @CR & 'Meine Externe ip Addresse ist: ' & _GetIp(), 100 )
			Else
				MsgBox ( 48, 'Network Status', 'Intranet-Server ist in Ordnung.' & @CR & 'Internet nicht verfügbar', 100 )
			EndIf

		Else
			MsgBox ( 48, 'Network Status', 'Intranet-Server antworten nicht', 100 )
		EndIf
;------
	Else
		MsgBox ( 48, 'Network Status', 'Unbehandeltes Adressbereich', 100 )
	EndIf

	EndIf

EndIf

MsgBox( 64, 'System Info', 'PC Name: ' & @ComputerName & @CR & 'User: ' & @UserName & @CR & 'OS Version: ' & @OSVersion & @CR & 'Build: ' & @OSBuild & @CR & 'Service Pack: ' & @OSServicePack, 100 )

Exit