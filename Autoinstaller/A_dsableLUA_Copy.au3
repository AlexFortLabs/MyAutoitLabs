#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         A.Fortowski

 Script Function:
	Script als Admin aus Netzfreigabe starten.

#ce ----------------------------------------------------------------------------

; #RequireAdmin

;otkljuchaem kontrol' uchetnyh zapisej
;RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA", "REG_DWORD", "00000000")

If DirCopy("\auto_install", "C:\auto_install", 1) = 1 then
   ;If FileCopy("C:\auto_install\B_install_soft.exe", @StartupDir & "\B_install_soft.exe") = 1 then
	  ;-Shutdown(6)
	  $iPID = Run("C:\auto_install\B_install_soft.exe")
	  If $iPID < 1 Then
		 MsgBox(16,"","Programm B_install_soft.exe konnte nicht gestartet werden" & @CRLF & "Beenden!")
		 Exit
	  Else
		 Sleep(1000)
	  EndIF
   ;EndIf
EndIf
