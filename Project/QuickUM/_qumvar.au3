; #VARIABLES# ====================================================================
; AD
Global $sStrNanchname				; Töns
Global $sStrNanchnameLang
Global $sStrNanchnameOK				; Toens
Global $sStrVorname					; Rudolf
Global $sStrVornameLang				; Rudolf
Global $sStrVornameOK				; Ru
Global $sStrADUserRoh				; Töns, Rudolf
Global $sNickName					; r.toens
Global $sPasswd = "123456789"
Global $sStrADUser = "Test-User"	; ToensRu
Global $sStrADUserText = "Toens Rudolf"
Global $sDescription = "== Mitarbeiter Vertrieb =="
Global $sStrADUseBIG				; TOENSRU
Global $sHomeDirectory = "\\funkegruppe.de\userdata\Userlaufwerk\"
Global $sHomeDrive = "P:"

Global $sBericht = ("Zusammenfassung zum Vorgang" & @CRLF & @CRLF)
Global $nVorgang = 0 ; und progressbar-saveposition

Global $sDomainName = "funkegruppe.de"
Global $sParentOU = "CN=Users,DC=funkegruppe,DC=de"

; HomeMDB = 		; LDAP-Pfad zur Datenbank
; msExchMailboxServerName = 			; Auch dieses Feld sollte "gefüllt" sein, weil einiger Code das noch abprüft, aber maßgeblich für den zuständigen Server ist der aktive Server für die HomeMDB
Global $sExchHomeServer = "/o=FUNKEGRUPPE/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Configuration/cn=Servers/cn=DESVR-MAIL01"

Const $sMailServer = "DESVR-MAIL01"

; SoftM / AS400
Global $nPersNr = 123
Const $sAbteilung = 1
Const $sFirma = "01"
Const $sRechte = 70
Global $sMenu = "IVEROOT" 			; IEKROOT ...IROOT

; GUI Form
Global $idProgressBar1
Global $bCheckSoftM = True
Global $bCheckWaage = False
Global $bCheckPostfach = True
Global $bCheckCRM = True
; ================================================================================