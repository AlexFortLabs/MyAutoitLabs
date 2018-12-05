; Abfrage Yes, No, Cancel

$bWaage	= False

Switch MsgBox(3, 'In Logistik', "Waage User?")
	Case 6 ; JA
		; Code für JA
		$bWaage	= True
	Case 7 ; NEIN
		; Code für NEIN
		$bWaage	= False
	Case 2 ; CANCEL
		$bWaage	= False
		Exit
EndSwitch
