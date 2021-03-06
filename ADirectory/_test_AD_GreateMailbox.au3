#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=favicon.ico
#AutoIt3Wrapper_Outfile_x64=_test_AD_GreateMailbox.exe
#AutoIt3Wrapper_UseX64=Y
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; *****************************************************************************
; Example 1
; Create a mailbox for a user.
; *****************************************************************************
#include <AD.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

; Open Connection to the Active Directory

_AD_Open()
;_AD_Open("funkegruppe\it", "geheim1")
If @error Then Exit MsgBox(16, "Active Directory Example Skript", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)

Global $iReply = MsgBox(308, "Active Directory Functions - Example 1", "This script creates a mailbox for a user." & @CRLF & @CRLF & _
		"Are you sure you want to change the Active Directory?")
If $iReply <> 6 Then Exit

Global $sUser, $sMailbox, $sIStore, $sServer, $sGroup, $sServerGroup
; Get the forms data for the current user
Global $sHomeMDB = _AD_GetObjectAttribute(@UserName, "HomeMDB")
If @error = 0 Then
	Global $aTemp = StringSplit($sHomeMDB, ",")
	$sMailbox = StringMid($aTemp[1], 4)
	$sIStore = StringMid($aTemp[2], 4)
	$sServer = StringMid($aTemp[4], 4)
	$sGroup = StringMid($aTemp[6], 4)
	$sServerGroup = StringMid($aTemp[8], 4)
EndIf
; Enter the necessary data
#region ### START Koda GUI section ### Form=
Global $Form1 = GUICreate("Active Directory Functions - Example 1", 514, 252)
GUICtrlSetFont(-1, 10, 800, 0, "Arial")
GUICtrlSetColor(-1, 0x33cc00)
GUICtrlSetBkColor(-1, 0xffffff)
GUISetBkColor("0xA0A0A0")
GUICtrlCreateLabel("User account (FQDN or samAccountName):", 8, 10, 231, 17)
GUICtrlCreateLabel("Mailbox storename:", 8, 42, 231, 17)
GUICtrlCreateLabel("Information store:", 8, 74, 231, 17)
GUICtrlCreateLabel("EMail server:", 8, 106, 231, 17)
GUICtrlCreateLabel("Administrative group in Exchange:", 8, 138, 231, 17)
GUICtrlCreateLabel("Exchange Server Group:", 8, 170, 231, 17)
Global $lUser = GUICtrlCreateInput(@UserName, 241, 8, 259, 21)
Global $IMailbox = GUICtrlCreateInput($sMailbox, 241, 40, 259, 21)
Global $IIStore = GUICtrlCreateInput($sIStore, 241, 72, 259, 21)
Global $IServer = GUICtrlCreateInput($sServer, 241, 104, 259, 21)
Global $IGroup = GUICtrlCreateInput($sGroup, 241, 136, 259, 21)
Global $IServerGroup = GUICtrlCreateInput($sServerGroup, 241, 168, 259, 21)
Global $BOK = GUICtrlCreateButton("Create mailbox", 8, 200, 121, 33)
Global $BCancel = GUICtrlCreateButton("Cancel", 428, 200, 73, 33, BitOR($GUI_SS_DEFAULT_BUTTON, $BS_DEFPUSHBUTTON))
GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

While 1
	Global $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $BCancel
			Exit
		Case $BOK
			$sUser = GUICtrlRead($lUser)
			$sMailbox = GUICtrlRead($IMailbox)
			$sIStore = GUICtrlRead($IIStore)
			$sServer = GUICtrlRead($IServer)
			$sGroup = GUICtrlRead($IGroup)
			$sServerGroup = GUICtrlRead($IServerGroup)
			ExitLoop
	EndSwitch
WEnd

; Create Mailbox
;Global $mailServer = $sServer		; DESVR-MAIL01
;Global $iValue = _AD_CreateMailbox($sUser, $sMailbox, $sIStore, $sServer, $sGroup, $sServerGroup)
;Global $iValue = _AD_CreateMailbox($sUser, "Mailbox Store (" & $sServer & ")")
;Global $iValue = _AD_CreateMailbox($sUser, "smtp.funkegruppe.de") 	;"Mailbox Store (" & $sServer & ")")
;Global $iValue = _AD_CreateMailboxPS($sUser, "desvr-mail-fws.funkegruppe.de") 	; per PShell - habe Code 7
;Global $iValue = _AD_CreateMailbox($sUser, "/o=FUNKEGRUPPE/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Configuration/cn=Servers/cn=DESVR-MAIL01")
Global $iValue = _AD_CreateMailbox($sUser, "Mailbox Store (" & $sMailbox & ")", "First Storage Group", $sServer, "First Administrative Group", $sServerGroup)

If $iValue = 1 Then
	MsgBox(64, "Active Directory Functions - Example 1", "Mailbox for User '" & $sUser & "' successfully created")
ElseIf @error = 1 Then
	MsgBox(64, "Active Directory Functions - Example 1", "User '" & $sUser & "' does not exist")
ElseIf @error = 2 Then
	MsgBox(64, "Active Directory Functions - Example 1", "User '" & $sUser & "' already has a mailbox defined")
Else
	MsgBox(64, "Active Directory Functions - Example 1", "Return code '" & @error & "' from Active Directory")
EndIf

; Close Connection to the Active Directory
_AD_Close()