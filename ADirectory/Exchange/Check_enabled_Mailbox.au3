;  check attribute HomeMDB to see if a user is mailbox enabled

#Include <ad.au3>
; *****************************************************************************
;~ #Include <adfunctions.au3>

;~ Global $ADUsertab[1]

;~ ; Get all Users in OU
;~ _ADGetObjectsInOU($ADUserTab, "CN=Users,DC=funkegruppe,DC=de",  "(&(objectCategory=user))", 2, "sAMAccountName")

;~ ; Check if mail enabled
;~ _ArrayDisplay($ADUserTab)
;~ For $i = 1 To $ADUserTab[0]
;~   ConsoleWrite($ADUserTab[$i] & "  " & _ADGetObjectAttribute($ADUserTab[$i],"HomeMDB") & @CRLF)
;~ Next
; *****************************************************************************

_AD_Open()

; list users with no HomeDirectory
;$a = _AD_GetObjectsInOU("users with no HomeDirectory", "(&(objectcategory=user)(objectclass=user)(!homedirectory=*))",2,"distinguishedname,sAMAccountName,homeDirectory")
;_ArrayDisplay($a)

; list users with no group membership
;$a = _AD_GetObjectsInOU("users with no group membership", "(&(objectcategory=user)(objectclass=user)(!memberof=*))",2,"distinguishedname,sAMAccountName,memberof")
;_ArrayDisplay($a)

; ++ Returns a list of users where the common name starts with "Xxxxx".
;~ $sUser = "Funke"
;~ $aObjects = _AD_GetObjectsInOU("", "(&(objectcategory=person)(objectclass=user)(cn=" & $sUser & "*))", 2, "sAMAccountName,distinguishedName,displayname", "displayname")
;~  _ArrayDisplay($aObjects)
; +++ Ende

; -- Set Random Password for AD-User
;$iSet = _AD_SetPassword($sUsername, Random(000001, 999998, 1)) ; set random password
;    If @error Then _Message("ERROR - Error setting random password - " & @error, True)

; +++ User ist Mitglied in folgenden Gruppen
;_AD_Open()
;$ADGroup = _AD_RecursiveGetMemberOf(@UserName)
;~ $ADGroup = _AD_RecursiveGetMemberOf("FunkeCh")
;~ If IsArray($ADGroup) Then
;~   _ArrayDisplay($ADGroup)
;~ EndIf
;_AD_Close()
; +++ Ende


ConsoleWrite("@AutoItX64: " & @AutoItX64 & @CRLF)
ConsoleWrite("@CPUArch: " & @CPUArch & @CRLF)
ConsoleWrite("@OSArch: " & @OSArch & @CRLF)

; *****************************************************************************
; Example 1
; Get a list of disabled accounts
; *****************************************************************************
;~ Global $aProperties = _AD_GetObjectsDisabled()
;~ _ArrayDisplay($aProperties)

;~ Global $aDisabled[1]
; *****************************************************************************
; Example 2
; Get a list of disabled accounts
; *****************************************************************************
;~ $aDisabled = _AD_GetObjectsDisabled()
;~ If @error > 0 Then
;~     MsgBox(64, "Active Directory Functions - Example 1", "No disabled user accounts could be found")
;~ Else
;~     _ArrayDisplay($aDisabled, "Active Directory Functions - Example 1 - Disabled User Accounts")
;~ EndIf


;$array = _AD_GetObjectsInOU("", "(&(objectcategory=user)(description=#resc#*)(userAccountControl:1.2.840.113556.1.4.803:=" & $ADS_UF_ACCOUNTDISABLE & "))", 2, "distinguishedname,description")
;_ArrayDisplay($array)


_AD_Close()

;FindUser("MaestroFm", "funkegruppe.de")

Func FindUser($username, $domain)
_AD_Open()
    $aDC = _AD_ListDomainControllers()
    _ArrayDisplay($aDC)

_AD_Close()
    For $x = 0 To UBound($aDC) - 1
        MsgBox('', 'Domain Controller Checking...', $aDC[$x][0])
        _AD_Open('', '', '', $aDC[$x][0])
        If Not @error Then
            $oUser = ObjGet("WinNT://" & $domain & "/" & $UserName)
            If _AD_IsObjectLocked($UserName) Or $oUser.IsAccountLocked = True Then
                MsgBox('', 'LOCKED on DC', $aDC[$x][0])
            ElseIf @error = 1 Then
                MsgBox('', 'User does not exist on DC', $aDC[$x][0])
            Else
                MsgBox('', 'NOT LOCKED on DC', $aDC[$x][0])
            EndIf
        Else
            MsgBox('', 'ERROR', @error)
        EndIf
        _AD_Close()
    Next
_ArrayDelete($aDC, 0)
EndFunc   ;==>FindUser


;Global $aResult = _AD_ListSchemaVersions()
;Global $aResult = _AD_GetObjectClass(@ComputerName & "$")
;Global $aResult = _AD_GetObjectClass("CN=fortowskal,OU=User,DC=funkegruppe,DC=de" & "$")
;ConsoleWrite($aResult[4])


;Global $aTemp = _AD_ListRootDSEAttributes()
;_ArrayDisplay($aTemp, "RootDSEAttributes")



