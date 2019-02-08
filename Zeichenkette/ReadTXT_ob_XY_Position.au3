; eine Zeile AB einer bestimmten Stelle auslesen.

;Datei einlesen
$f = FileOpen("C:\Temp\info.txt")

; z.B. zweite Zeile der Datei auslesen
$line = FileReadLine($f,2)

; alles hinter dem 10 Zeichen in der Zeile in der Variablen $str speichern
$str = StringMid($line,11)

; String anzeigen
msgbox (0,"",$str)

;FileHandle schlieﬂen
FileClose($f)