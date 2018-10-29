#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=mail.ico
#AutoIt3Wrapper_Outfile=getinboxsize.exe
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Comment=Supported: Outlook 2003/2007/2010, Thunderbird, Outlook Express
#AutoIt3Wrapper_Res_Description=Get inboxsizes from various mailclients
#AutoIt3Wrapper_Res_Fileversion=0.1.0.68
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Bastian Schulte
#AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         Bastian Schulte

 Script Function:
  Get mailboxsize from Thunderbird, Outlook Express & Outlook (2003, 2007, 2010)

#ce ----------------------------------------------------------------------------

#Include <File.au3>
#Include <Array.au3>

Global $check_ou2003
Global $check_ou2007
Global $check_ou2010

;$CmdLine[1]: Display size in MB or KB
;$CmdLine[2]: Targetserver for collecting information of mailbox sizes in one place
If $CmdLine[0] < 1 Then
	ConsoleWrite("Description:" & @CRLF)
	ConsoleWrite("Retrieve sizes of mailboxes on local clients." & @CRLF)
	ConsoleWrite("---------------" & @CRLF)
	ConsoleWrite("Developed by Bastian Schulte" & @CRLF)
	ConsoleWrite("Website: http://autohaus-it.bs101.de" & @CRLF & @CRLF)
	ConsoleWrite("Usage:" & @CRLF)
	ConsoleWrite("getinboxsize [-kb|-mb] [\\server\smb-share\mailboxsizes.txt]" & @CRLF & @CRLF)
	ConsoleWrite("-kb,-mb: choose wether to display mailbox size in kilobytes or megabytes" & @CRLF)
	ConsoleWrite("second parameter (opt): upload output to one server (SMB/local file only,full path to targetfile)")
	Exit
EndIf

;Thunderbird
$check_tbird = FileExists (@AppDataDir & "\Thunderbird\profiles.ini")
If $check_tbird = 1 Then
	$use_tb = "Thunderbird"
	$tb_profile = IniRead(@AppDataDir & "\Thunderbird\profiles.ini","Profile0","Path","")
	$tb_profile_mod = StringReplace($tb_profile,"/","\")
	$tb_prof_dir = @AppDataDir & "\Thunderbird\" & $tb_profile_mod & "\Mail"
	$tb_size = DirGetSize($tb_prof_dir,0)
	If $CmdLine[1] = "-mb" Then
		$tb_size_mb = Round($tb_size / 1024 / 1024 )
	ElseIf $CmdLine[1] = "-kb" Then
		$tb_size_mb = Round($tb_size / 1024 )
	EndIf
	ConsoleWrite(@UserName & " (Thunderbird):" & $tb_size_mb & @CRLF)
ElseIf $check_tbird = 0 Then
	ConsoleWrite("Thunderbird profile not found" & @CRLF)
EndIf

;Outlook Express
If @OSVersion = "WIN_XP" Or @OSVersion = "WIN_2000" Then
	$oe_userid = RegRead("HKEY_CURRENT_USER\Identities","Default User ID")
	$oe_profile = RegRead("HKEY_CURRENT_USER\Identities\" & $oe_userid & "\Software\Microsoft\Outlook Express\5.0","Store Root")
	$oe_profile_mod = StringReplace($oe_profile,"%UserProfile%",@UserProfileDir)
	$check_oe = FileExists($oe_profile_mod)
	If $check_oe = 1 Then
		$use_oe = "OutlookExpress"
		$oe_size = DirGetSize($oe_profile_mod,0)
		If $CmdLine[1] = "-mb" Then
			$oe_size_mb = Round($oe_size / 1024 / 1024 )
		ElseIf $CmdLine[1] = "-kb" Then
			$oe_size_mb = Round($oe_size / 1024 )
		EndIf
		ConsoleWrite(@UserName & " (Outlook Express):" & $oe_size_mb & @CRLF)
	ElseIf $check_oe = 0 Then
		ConsoleWrite("Outlook Express profile not found" & @CRLF)
	EndIf
Else
	$check_oe = 0
	ConsoleWrite("Windows is newer than XP -> next" & @CRLF)
EndIf

;Outlook
For $y=1 To 20
	$office_version = RegEnumKey("HKCU\Software\Microsoft\Office",$y)
;Outlook 2003
	If $office_version = "11.0" Then
		$check_ou2003 = FileExists(@UserProfileDir & "\Lokale Einstellungen\Anwendungsdaten\Microsoft\Outlook\Outlook.pst")
		If $check_ou2003 = 1 Then
			$use_ou2003 = "Outlook2003"
			$ou2003_size = DirGetSize(@UserProfileDir & "\Lokale Einstellungen\Anwendungsdaten\Microsoft\Outlook",0)
			If $CmdLine[1] = "-mb" Then
				$ou2003_size_mb = Round($ou2003_size / 1024 / 1024 )
			ElseIf $CmdLine[1] = "-kb" Then
				$ou2003_size_mb = Round($ou2003_size / 1024 )
			EndIf
			ConsoleWrite(@UserName & " (Outlook 2003):" & $ou2003_size_mb & @CRLF)
		ElseIf $check_ou2003 = 0 Then
			ConsoleWrite("Outlook 2003 PST file not found" & @CRLF)
		EndIf
;Outlook 2007
	ElseIf $office_version = "12.0" Then
		$check_ou2007 = FileExists(@UserProfileDir & "\Lokale Einstellungen\Anwendungsdaten\Microsoft\Outlook\Outlook.pst")
		If $check_ou2007 = 1 Then
			$use_ou2007 = "Outlook2007"
			$ou2007_size = DirGetSize(@UserProfileDir & "\Lokale Einstellungen\Anwendungsdaten\Microsoft\Outlook",0)
			If $CmdLine[1] = "-mb" Then
				$ou2007_size_mb = Round($ou2007_size / 1024 / 1024 )
			ElseIf $CmdLine[1] = "-kb" Then
				$ou2007_size_mb = Round($ou2007_size / 1024 )
			EndIf
			ConsoleWrite(@UserName & " (Outlook 2007):" & $ou2007_size_mb & @CRLF)
		ElseIf $check_ou2007 = 0 Then
			ConsoleWrite("Outlook 2007 PST file not found" & @CRLF)
		EndIf
;Outlook 2010
	ElseIf $office_version = "14.0" Then
		$aCheck_ou2010 = _FileListToArray(@MyDocumentsDir & "\Outlook-Dateien","*.pst",1)
		If Not @error Then
			For $j=1 To $aCheck_ou2010[0]
				$check_ou2010 = FileExists(@MyDocumentsDir & "\Outlook-Dateien\" & $aCheck_ou2010[$j])
				If $check_ou2010 = 1 then ExitLoop
			Next
		ElseIf @error Then
			$check_ou2010 = 0
		EndIf
		If $check_ou2010 = 1 Then
			$use_ou2010 = "Outlook2010"
			$ou2010_size = DirGetSize(@MyDocumentsDir & "\Outlook-Dateien",0)
			If $CmdLine[1] = "-mb" Then
				$ou2010_size_mb = Round($ou2010_size / 1024 / 1024 )
			ElseIf $CmdLine[1] = "-kb" Then
				$ou2010_size_mb = Round($ou2010_size / 1024 )
			EndIf
			ConsoleWrite(@UserName & " (Outlook 2010):" & $ou2010_size_mb & @CRLF)
		ElseIf $check_ou2010 = 0 Then
			ConsoleWrite("Outlook 2010 PST file not found" & @CRLF)
		EndIf
	ElseIf @error <> 0 then
		ExitLoop
	EndIf
Next


;Server upload
If $CmdLine[0] = 2 Then
	$check_outfile = FileExists($CmdLine[2])
		If $check_outfile = 0 Then
			_FileCreate($CmdLine[2])
				If @error Then
					ConsoleWrite ("error creating " & $CmdLine[2] & @CRLF)
					ConsoleWrite ("possible errors: 1=cannnot open file, 2=cannot write to file")
					Exit
				EndIf
			$logfile_handle = FileOpen($CmdLine[2],1)
			FileWrite($logfile_handle,"#Mailboxsizes of users from " & @LogonDomain & @CRLF & "#started collecting data at " & @MDAY & "." & @MON & "." & @YEAR & @CRLF & @CRLF)
			FileWrite($logfile_handle,"#computername|username|mailclient|mailboxsize" & @CRLF)
			If $check_tbird = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_tb & "|" & $tb_size_mb & @CRLF)
			ElseIf $check_tbird = 0 Then
			EndIf
			If $check_oe = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_oe & "|" & $oe_size_mb & @CRLF)
			ElseIf $check_oe = 0 Then
			EndIf
			If $check_ou2003 = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_ou2003 & "|" & $ou2003_size_mb & @CRLF)
			ElseIf $check_ou2003 = 0 Then
			EndIf
			If $check_ou2007 = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_ou2007 & "|" & $ou2007_size_mb & @CRLF)
			ElseIf $check_ou2007 = 0 Then
			EndIf
			If $check_ou2010 = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_ou2010 & "|" & $ou2010_size_mb & @CRLF)
			ElseIf $check_ou2010 = 0 Then
			EndIf
			FileClose($logfile_handle)
		ElseIf $check_outfile = 1 Then
			$logfile_handle = FileOpen($CmdLine[2],1)
			If $check_tbird = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_tb & "|" & $tb_size_mb & @CRLF)
			ElseIf $check_tbird = 0 Then
			EndIf
			If $check_oe = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_oe & "|" & $oe_size_mb & @CRLF)
			ElseIf $check_oe = 0 Then
			EndIf
			If $check_ou2003 = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_ou2003 & "|" & $ou2003_size_mb & @CRLF)
			ElseIf $check_ou2003 = 0 Then
			EndIf
			If $check_ou2007 = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_ou2007 & "|" & $ou2007_size_mb & @CRLF)
			ElseIf $check_ou2007 = 0 Then
			EndIf
			If $check_ou2010 = 1 Then
				FileWrite($logfile_handle,@ComputerName & "|" & @UserName & "|" & $use_ou2010 & "|" & $ou2010_size_mb & @CRLF)
			ElseIf $check_ou2010 = 0 Then
			EndIf
			FileClose($logfile_handle)
		EndIf
EndIf