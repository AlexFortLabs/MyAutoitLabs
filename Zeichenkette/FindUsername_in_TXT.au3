; txt nach username durchsuchen

$Datei1 = 'Test1.txt' ;Ort der TXT-Datei
$String1 = FileRead($Datei1) ;lesen der TXT-Datei
$Suchstring = @username ;nach was gesucht werden soll

If StringInStr($String1, $Suchstring) Then
	MsgBox(0,"Melde", "Sie stehen in Test1.txt!")
Else
	MsgBox(0, "Melde", "Sie stehen nirgendwo drin!")
EndIf