; AD
Global $sStrNanchname				; Töws
Global $sStrNanchnameLang
Global $sStrNanchnameOK				; Toews
Global $sStrVorname					; Rudolf
Global $sStrVornameLang				; Rudolf
Global $sStrVornameOK				; Ru
Global $sStrADUserRoh				; Töws, Rudolf
Global $sNickName					; r.toews
Global $sPasswd = "123456"
Global $sStrADUser = "Test-User"	; ToewsRu
Global $sStrADUserText = "Bianco Melanie"
Global $sDescription = "== Mitarbeiter Vertrieb =="
Global $sStrADUseBIG				; TOEWSRU
Global $sHomeDirectory = "\\funkegruppe.de\userdata\Userlaufwerk\"
Global $sHomeDrive = "P:"

Global $sDomainName = "funkegruppe.de"
Global $sParentOU = "CN=Users,DC=funkegruppe,DC=de"

; HomeMDB = 		; LDAP-Pfad zur Datenbank
; msExchMailboxServerName = 			; Auch dieses Feld sollte "gefüllt" sein, weil einiger Code das noch abprüft, aber maßgeblich für den zuständigen Server ist der aktive Server für die HomeMDB

Const $mailServer = "DESVR-MAIL01"

; SoftM / AS400
Global $nPersNr = 123
Const $sAbteilung = 1
Const $sFirma = "01"
Const $sRechte = 70
Global $sMenu = "IVEROOT" 			; IEKROOT ...IROOT

Global $bSoftM = False
Global $bWaage = False
Global $bPostfach = False