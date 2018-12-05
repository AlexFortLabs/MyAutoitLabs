; Benutzer / benutzerprofiel auf AS400 erstellen
;#RequireAdmin
;MsgBox(0, "", "IsAdmin: " & IsAdmin())

; logon before rmtcmd will work
$CMD1 = "CWBLOGON funkei5 /u softmopr /p softmopr"
$CMD2 = "RMTCMD CRTUSRPRF USRPRF(RISCHARPJ) PASSWORD(geheim) INLMNU(*SIGNOFF) TEXT('Rischar, Pjer') SPCAUT(*NONE) JOBD(QGPL/QDFTJOBD) GRPPRF(SOFTM) OWNER(*GRPPRF)"
; To clear all userids from the cache:
$CMD3 = "CWBLOGON /c"

RunWait(@ComSpec & " /q /c " & $CMD2,@ScriptDir,@SW_SHOW)
