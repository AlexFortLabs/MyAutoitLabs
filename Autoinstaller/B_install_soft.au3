#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         A.Fortowski

 Script Function:
	Script führt automatische Installation mit Bereinigung durch

#ce ----------------------------------------------------------------------------

; #RequireAdmin

; in diesem Modul befindet sich _TempFile() Funktion.
#Include <File.au3>

$Folder = "C:\auto_install"

; Posle togo, kak budet vyzvana jeta funkcija, obratnogo puti uzhe ne budet. Kak tol'ko skript zakonchit svoju rabotu, EXE-fajl budet bezvozvratno udalen.
; Script sozdaet v Temp papke BAT fajl, zatem zapuskaet ego. V svoju ochered', BAT fajl zhdet zavershenija osnovnoj programmy, posle chego udaljaet ee i sebja.
Func _ScriptDestroy()
   If StringRight(@ScriptFullPath,3) <> "au3" Then
	  $sTemp = _TempFile(@TempDir, '~', '.bat')
	  $sPath = FileGetShortName(@ScriptFullPath)
	  $hFile = FileOpen($sTemp, 2)
	  FileWriteLine($hFile, '@echo off')
	  FileWriteLine($hFile, ':loop')
	  FileWriteLine($hFile, 'del ' & $sPath)
	  FileWriteLine($hFile, 'if exist ' & $sPath & ' goto loop')
	  FileWriteLine($hFile, 'del ' & $sTemp)
	  FileClose($hFile)
	  Run($sTemp, '', @SW_HIDE)
   EndIf
 EndFunc   ;==>_ScriptDestroy

Run($Folder & "\setup-lightshot.exe", $Folder)
MsgBox(0, '', 'Wait...Lightshot')
;Sleep(5000)
Run("VNC-Viewer-6.17.1113-Windows.exe")
MsgBox(0, '', 'Wait...VNC')
;Sleep(5000)
Run("FileZilla_3.32.0_win64-setup_bundled.exe")
;Run($Folder & "\FileZilla-Server.exe", $Folder)
MsgBox(0, '', 'Wait...FileZilla')
;Sleep(5000)
Run("ownCloud-2.4.1.9270-setup.exe")
MsgBox(0, '', 'Wait...ownCloud')
;Sleep(5000)

;Run("WinDjView.exe")
;Run("TotalCom_7.50.exe")
;Run("KMPlayer_2.9.4.1435.exe")
;Run("KLiteMega_5.61.exe")
;Run("FoxitReader_3.0.1506.exe")
;Run("Everest_5.30.1977.exe")
;Run("AIMP_2.60.530.exe")
;Run("3DMark06.exe")

;DirCreate("C:\dx")
;Run("directx_Jun2010_redist.exe")
;WinWaitActive("DirectX June 2010 Redist","read the following")
;Send("{TAB}")
;send("{space}")
;WinWaitActive("DirectX June 2010 Redist"," type the location where")
;ControlSend("DirectX June 2010 Redist","","Edit1","C:\dx")
;ControlClick("DirectX June 2010 Redist","type the location where","Button2")
;Sleep(10000)
;Run("C:\dx\DXSETUP.exe")
;WinWaitActive("Установка Microsoft(R) DirectX(R)","Мастер установки DirectX")
;ControlClick("Установка Microsoft(R) DirectX(R)","Мастер установки DirectX","Button1")
;ControlClick("Установка Microsoft(R) DirectX(R)","Мастер установки DirectX","Button4")
;WinWaitActive("Установка Microsoft(R) DirectX(R)","Установка исполняемого модуля")
;ControlClick("Установка Microsoft(R) DirectX(R)","Установка исполняемого модуля","Button4")
;WinWaitActive("Установка Microsoft(R) DirectX(R)","Установка завершена")
;send("{ENTER}")
;FileDelete(@DesktopDir & "\Vit Registry Fixer.lnk")
;FileDelete(@DesktopDir & "\Vit Disk Cleaner.lnk")
;FileDelete(@StartupDir & "\Total Commander.lnk")
;FileDelete(@DesktopCommonDir & "\Vit Registry Fixer.lnk")
;FileDelete(@DesktopCommonDir & "\Vit Disk Cleaner.lnk")
;FileDelete(@StartupCommonDir & "\Total Commander.lnk")
;FileDelete(@StartupDir & "\script.lnk")
;RunWait("Office_2010_SP3_Full.exe")
;sleep(120000)
;Sleep(800)
;MsgBox(0, '', 'Wait...')

; Script soll Ordner hinter sich aufraeumen
;DirRemove("C:\auto_install",1)
; V kakom meste vyzyvat' _ScriptDestroy() po bol'shemu schetu ne imeet znachenija, t.k. jeto tol'ko iniciiruet otlozhennoe udalenie, do zavershenija skripta.
_ScriptDestroy()

;Shutdown(6)
