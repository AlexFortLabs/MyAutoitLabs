; Abfrage Yes, No, Cancel

$bWaage	= False

Switch MsgBox(3, 'In Logistik', "Waage User?")
	Case 6 ; JA
		; Code f�r JA
		$bWaage	= True
	Case 7 ; NEIN
		; Code f�r NEIN
		$bWaage	= False
	Case 2 ; CANCEL
		$bWaage	= False
		Exit
EndSwitch
