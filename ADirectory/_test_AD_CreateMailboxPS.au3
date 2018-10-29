#include <AD.au3>

_AD_Open()
Local $vResult = _AD_CreateMailboxPS("TrutanAD", "https://owa.funkegruppe.de")
If @error > 0 Then
    If @error < 7 Then
        MsgBox(0, "Error!", "_AD_CreateMailboxPS returned @error = " & @error & ", @extended = " & @extended)
    Else
        _ArrayDisplay($vResult, "_AD_CreateMailboxPS returned @error = " & @error & ", @extended = " & @extended)
    EndIf
Else
    _ArrayDisplay($vResult, "Successful!")
EndIf
_AD_Close()
Exit

; #FUNCTION#====================================================================================================================
; Name...........: _AD_CreateMailboxPS
; Description ...: Creates a mailbox for a user using PowerShell
; Syntax.........: _AD_CreateMailboxPS($sUser, $sURI[, $sSessionParam = Default[, $sMailboxParam = Default]])
; Parameters ....: $sUser         - User account (SamAccountName or FQDN) for which you want to create the mailbox
;                  $sURI          - Specifies a URI that defines the connection endpoint for the session. The URI must be fully qualified.
;                                   Example: http://YourExchangeServerNameGoesHere.CompanyName.com
;                  $sSessionParam - Optional: One or multiple additional parameters for the PowerShell "Session" command e.g. " -Authentication Kerberos"
;                  $sMailboxParam - Optional: One or multiple additional parameters for the PowerShell "Enable-Mailbox" command (see parameter $sSessionParam)
; Return values .: Success - Zero based one-dimensional array holding the StdOut messages written by PowerShell
;                  Failure - "", sets @error to:
;                  |1 - $sUser does not exist
;                  |2 - $sUser already has a mailbox
;                  |3 - $sUser is invalid (empty)
;                  |4 - $sURI is invalid (empty)
;                  |5 - Run returned an error (PowerShell could not be started). @extended is set to the @error returned by Run
;                  |6 - Writing to StdIn returned an error. @extended is set to the @error returned by StdinWrite
;                  Failure - Zero based one-dimensional array holding the StdErr messages written by PowerShell, sets error to:
;                  |7 - PowerShell has written some error messages to StdErr.
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _AD_CreateMailboxPS($sUser, $sURI, $sSessionParam = Default, $sMailboxParam = Default)

    Local $aResult
    If StringStripWS($sUser, $STR_STRIPALL) = "" Then Return SetError(3, 0, "")
    If StringStripWS($sURI, $STR_STRIPALL) = "" Then Return SetError(4, 0, "")
    If $sSessionParam = Default Then $sSessionParam = ""
    If $sMailboxParam = Default Then $sMailboxParam = ""
    If Not _AD_ObjectExists($sUser) Then Return SetError(1, 0, "")
    Local $sProperty = "sAMAccountName"
    If StringMid($sUser, 3, 1) = "=" Then $sProperty = "distinguishedName" ; FQDN provided
    $__oAD_Command.CommandText = "<LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain & ">;(" & $sProperty & "=" & $sUser & ");ADsPath;subtree"
    Local $oRecordSet = $__oAD_Command.Execute ; Retrieve the ADsPath for the object
    Local $sLDAPEntry = $oRecordSet.fields(0).Value
    Local $oUser = __AD_ObjGet($sLDAPEntry) ; Retrieve the COM Object for the object
    If $oUser.HomeMDB <> "" Then Return SetError(2, 0, "")
    Local $sCMD = "Powershell -Command -", $sSTDOUT = "", $sSTDERR = "", $bError = False
    $iPID = Run($sCMD, @SystemDir, @SW_MAXIMIZE, $STDIN_CHILD + $STDOUT_CHILD + $STDERR_CHILD)
    If $iPID = 0 Or @error Then Return SetError(5, @error, "")
    $sCMD = "$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri " & $sURI & " " & $sSessionParam & _
        ";Import-PSSession $Session" & _
        ";Enable-Mailbox -Identity " & $sUser & " " & $sMailboxParam & _
        ";Remove-PSSession $Session"
    StdinWrite($iPID, $sCMD)
    If @error Then Return SetError(6, @error, "")
    StdinWrite($iPID)
    ; Process STDOUT
    While 1
        $sOutput = StdoutRead($iPID)
        If @error Then ExitLoop
        If $sOutput <> "" Then $sSTDOUT = $sSTDOUT & $sOutput
    WEnd
    ; Process STDERR
    While 1
        $sOutput = StderrRead($iPID)
        If @error Then ExitLoop
        If $sOutput <> "" Then $sSTDERR = $sSTDERR & $sOutput
        $bError = True
    WEnd
    If $bError Then
        $aResult = StringSplit($sSTDERR, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)
        Return SetError(7, 0, $aResult)
    Else
        $aResult = StringSplit($sSTDOUT, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)
        Return $aResult
    EndIf
EndFunc   ;==>_AD_CreateMailboxPS