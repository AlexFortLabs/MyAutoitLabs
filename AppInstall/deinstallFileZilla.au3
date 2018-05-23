#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         A.Fortowski

 Script Function:
	Das Deinstallationsskript ruft den Uninstaller mit dem Parameter /S auf und wartet anschließend darauf,
	dass der Prozess au_.exe geschlossen wird. Der Timeout beträgt 60 Sekunden.
	Liefert eine Installation den Rückgabewert 3010 - „Neustart erforderlich“, gibt das Skript diesen Wert auch an den SCCM zurück

#ce ----------------------------------------------------------------------------

;ScriptOptionen festlegen
Opt("ExpandEnvStrings",1)	;Umgebungsvariablen interpretieren

$exitcode = RunWait("%programfiles%\FileZilla FTP Client\uninstall.exe /S")
ProcessWaitClose("au_.exe",60)

exit $exitcode

