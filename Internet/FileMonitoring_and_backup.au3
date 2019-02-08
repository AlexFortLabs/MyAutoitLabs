; ein Scrip (wenn möglich mit AUTOIT) das diesen ordner überwacht und alle neuen dateien im abstand von 5 minuten als e-mail(1und1) Versendet inkl. der PDF mit im Anhang,
; und dann die PDF in einen extra ordner als Backup verschiebt

#NoTrayIcon
#Include <file.au3>
; Variablen ------------------
global $CHECKDIR = "C:\Temp\source"
global $OUTDIR = "C:\Temp\target"
; Intervall in dem der Ordner überprüft werden soll (Sekunden)
Const $CHECKINTERVALL = 5
; Mail Daten
Global $FROMEMAIL = ""
Global $TOEMAIL = ""
Const $FROMNAME = "Dein Name"
Const $MAILSERVER = "SMTP.SERVER.DE"
Const $SMTPUSER = "SMTPUSERNAME"
Const $SMTPPASS = "SMTPPASSWORD"
Const $SMTPPORT = "25"
Const $SMTPUSESSL = 1
;-----------------------------
While 1
	$search = FileFindFirstFile($CHECKDIR & "\*.pdf")
	if $search <> -1 then
		While 1
			$file = FileFindNextFile($search)
			If @error Then ExitLoop
			local $szDrive, $szDir, $szFName, $szExt
			_PathSplit($file, $szDrive, $szDir, $szFName, $szExt)
			$pdfFile = $CHECKDIR & "\" & $file
			_INetSmtpMailCom($MAILSERVER,$FROMNAME,$FROMEMAIL,$TOEMAIL,"Neues Fax","Im Anhang finden sie das Fax als PDF",$pdfFile,"","","Normal",$SMTPUSER,$SMTPPASS,$SMTPPORT,$SMTPUSESSL)
			FileMove($pdfFile,$OUTDIR & "\",9)
		WEnd
	EndIf
	FileClose($search)
	Sleep($CHECKINTERVALL * 1000)
WEnd



; Variables for Function _INetSmtpMailCom
;##################################
;~ $SmtpServer = ""                            ; address for the smtp-server to use - REQUIRED
;~ $FromName = "User"                          ; name from who the email was sent
;~ $FromAddress = ""                           ; address from where the mail should come
;~ $ToAddress = ""                             ; destination address of the email - REQUIRED
;~ $Subject = "testsubject"                       ; subject from the email - can be anything you want it to be
;~ $Body = "This Is The Body"                  ; the messagebody from the mail - can be left blank but then you get a blank mail
;~ $AttachFiles = ""                           ; the file you want to attach- leave blank if not needed
;~ $CcAddress = ""                             ; address for cc - leave blank if not needed
;~ $BccAddress = ""                            ; address for bcc - leave blank if not needed
;~ $Importance = "Normal"                      ; Send message priority: "High", "Normal", "Low"
;~ $Username = ""                              ; username for the account used from where the mail gets sent - REQUIRED
;~ $Password = ""                              ; password for the account used from where the mail gets sent - REQUIRED
;~ $IPPort = 25                               ; port used for sending the mail

;~ $ssl=0                                   ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance="Normal", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)
    Local $objEmail = ObjCreate("CDO.Message")
    $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
    $objEmail.To = $s_ToAddress
    Local $i_Error = 0
    Local $i_Error_desciption = ""
    If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
    If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
    $objEmail.Subject = $s_Subject
    If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
        $objEmail.HTMLBody = $as_Body
    Else
        $objEmail.Textbody = $as_Body & @CRLF
    EndIf
    If $s_AttachFiles <> "" Then
        Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
        For $x = 1 To $S_Files2Attach[0]
            $S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
            ConsoleWrite('@@ Debug(62) : $S_Files2Attach = ' & $S_Files2Attach & @LF & '>Error code: ' & @error & @LF) ;### Debug Console
            If FileExists($S_Files2Attach[$x]) Then
                $objEmail.AddAttachment ($S_Files2Attach[$x])
            Else
                ConsoleWrite('!> File not found to attach: ' & $S_Files2Attach[$x] & @LF)
                SetError(1)
                Return 0
            EndIf
        Next
    EndIf
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    If Number($IPPort) = 0 then $IPPort = 25
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
    ;Update settings
    $objEmail.Configuration.Fields.Update
    ; Set Email Importance
    Switch $s_Importance
        Case "High"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "High"
        Case "Normal"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Normal"
        Case "Low"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Low"
    EndSwitch
    $objEmail.Fields.Update
    ; Sent the Message
    $objEmail.Send
    If @error Then
        SetError(2)
    EndIf
    $objEmail=""
EndFunc   ;==>_INetSmtpMailCom
