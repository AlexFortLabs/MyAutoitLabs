; _AD_GetAllOUs should return an array of ALL OUs in your AD

#AutoIt3Wrapper_AU3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#include <AD.au3>

; Open Connection to the Active Directory
_AD_Open()
If @error Then Exit MsgBox(16, "Active Directory Example Skript", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get a list of all OUs in the Active Directory
; *****************************************************************************
Global $aOUs = _AD_GetAllOUs()
If @error > 0 Then
MsgBox(64, "Active Directory Functions - Example 1", "No OUs could be found")
Else
_ArrayDisplay($aOUs, "Active Directory Functions - Example 1 - All OUs found in the Active Directory")
EndIf

;------------------------------
; Make your changes below
;------------------------------
;To get the FQDN for your user account you would use:
Global $sOU_user = _AD_SamAccountNameToFQDN(@UserName)

; To get the OU where your computer is a member of you need the following code:
Global $sOU_computer = _AD_SamAccountNameToFQDN(@Computername & "$")

Global $iPos1 = StringInStr($sOU_user, ",")
Global $iPos2 = StringInStr($sOU_computer, ",")
	Global $sOU1 = StringMid($sOU_user, $iPos1 + 1)
	Global $sOU2 = StringMid($sOU_computer, $iPos2 + 1)
	MsgBox(0, "User and Computer", "Username exist in the OU " & $sOU1 & @CR & " Computer in the OU " & $sOU2)

_AD_close()

