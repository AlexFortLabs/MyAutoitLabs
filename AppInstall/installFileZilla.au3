#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         A.Fortowski

 Script Function:
	Skript installiert FileZilla und kopiert die Files „Filezilla.xml“ (allgemeine Programmoptionen wie AutoUpdate)
	sowie „Sitemanager.xml“ (gespeicherte FTP Sites) in das Userprofil des aktuell angemeldeten Benutzers.
	Wenn das Programm in einer Tasksequenz installiert werden soll, könnte man hier die Files in das Default User Profile kopieren.

#ce ----------------------------------------------------------------------------

;ScriptOptionen festlegen
Opt("ExpandEnvStrings",1)	;Umgebungsvariablen interpretieren

;Filezilla Installer Silent auftrufen
$exitcode = RunWait("FileZilla_Server-0_9_60_2.exe /S")

;Filezilla Default Einstellungen vorgeben
FileCopy("Filezilla.xml","%userprofile%\AppData\Roaming\Filezilla\Filezilla.xml",8)	;Programmoptionen
FileCopy("sitemanager.xml","%userprofile%\AppData\Roaming\Filezilla\sitemanager.xml",8)	;vorgegebene FTP-Server

; Liefert eine Installation den Rückgabewert 3010 - „Neustart erforderlich“, gibt das Skript diesen Wert auch an den SCCM zurück
; exit 3010
exit $exitcode
