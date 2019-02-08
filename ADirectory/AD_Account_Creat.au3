#cs
Version 1.2
Added Show tips for Mailbox, Remote Access and Search on Main GUI
Added Show tips for User Name, Search on Search GUI
Changed Title of the Main GUI
#ce


#include <GUIConstants.au3>
#include <Array.au3>
#Include <GuiCombo.au3>
#Include <GuiEdit.au3>
#Include <GuiStatusBar.au3>
#Include <GuiList.au3>

FileInstall("C:\Documents and Settings\admin\Desktop\Working Scripts\AD\User Create\usercreationHelp.chm",@ScriptDir & "\")

Dim $strUserDN,$Form1,$locateUserName_input,$locateFirstName_input,$locateLastName_input,$locateADLocation_input
Dim $Home,$locateTemplates_listbox,$strTemplateCNLocate,$selectedTemplate
Dim $homeDirectory

Local $oMyError = ObjEvent("AutoIt.Error", "ComError")

Opt("GUIOnEventMode", 1)  ; Change to OnEvent mode

#Region ### START Koda GUI section ### Form=c:\documents and settings\admin\desktop\scripting\koda\forms\tabs.kxf
$Form1_1 = GUICreate("User Creation", 413, 305, -1, -1)
GUISetIcon("prometheus.ico")
$PageControl1 = GUICtrlCreateTab(8, 8, 396, 256)

GUISetOnEvent($GUI_EVENT_CLOSE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_RESTORE, "SpecialEvents")

;TabSheet 1
$TabSheet1 = GUICtrlCreateTabItem("Copy Template")
GUICtrlSetState(-1,$GUI_SHOW)
$templateUser_combo = GUICtrlCreateCombo("", 224, 40, 161, 25)
$templateUser_lbl = GUICtrlCreateLabel("Select Template User", 104, 40, 106, 17)

   Local $objCommand = ObjCreate("ADODB.Command")
   Local $objConnection = ObjCreate("ADODB.Connection")

   $objConnection.Provider = "ADsDSOObject"
   $objConnection.Open ("Active Directory Provider")
   $objCommand.ActiveConnection = $objConnection

   Local $strBase = "<LDAP://dc=testdomain,dc=com>"
   Local $strFilter = "(&(objectCategory=person)(objectClass=user)(cn=Template*))"
   Local $strAttributes = "cn,samAccountName"
   Local $strQuery = $strBase & ";" & $strFilter & ";" & $strAttributes & ";subtree"

   $objCommand.CommandText = $strQuery
   $objCommand.Properties ("Page Size") = 100
   $objCommand.Properties ("Timeout") = 30
   $objCommand.Properties ("Cache Results") = False

   Local $objRecordSet = $objCommand.Execute

   While Not $objRecordSet.EOF
    $strtemplateCN = $objRecordSet.Fields("samAccountName").value
    GuiCtrlSetData($templateUser_combo,$strTemplateCN)

    $objRecordSet.MoveNext
   Wend

   $objConnection.Close
   $objConnection = ""
   $objCommand = ""
   $objRecordSet = ""


;TabSheet 2


$TabSheet2 = GUICtrlCreateTabItem("Copy Existing User")

$userName2_input = GUICtrlCreateInput("", 24, 56, 89, 21)
$userName2_lbl = GUICtrlCreateLabel("User Name to Copy", 24, 40, 96, 17)
$displayName2_input = GUICtrlCreateInput("", 168, 56, 193, 21)
$userName2_btn = GUICtrlCreateButton("...", 120, 56, 27, 17, 0)
GUICtrlSetTip(-1, "Click to Show Display Name of User")
GUICtrlSetOnEvent($userName2_btn,"see_displayName")




GUICtrlCreateTabItem("")




$email_chkbox = GUICtrlCreateCheckbox("Create Mailbox", 32, 136, 97, 17)
GUICtrlSetTip(-1, "Check to Enable Mailbox Account")
$vpn_btn = GUICtrlCreateCheckbox("Enable Remote Access",170,136,150,17)
GUICtrlSetTip(-1, "Check for VPN Access")
$createUser_btn = GUICtrlCreateButton("&Create User", 130, 225, 139, 25, 0,$WS_EX_CLIENTEDGE )
$firstName_input = GUICtrlCreateInput("", 32, 100, 121, 21)
$Input2 = GUICtrlCreateInput("", 160, 100, 25, 21)
$lastName_input = GUICtrlCreateInput("", 192, 100, 185, 21)
$firstName_lbl = GUICtrlCreateLabel("First Name", 32, 80, 54, 17)
$mI_lbl = GUICtrlCreateLabel("M.I.", 160, 80, 22, 17)
$lastName_lbl = GUICtrlCreateLabel("Last Name", 192, 80, 55, 17)



$Button2 = GUICtrlCreateButton("&Exit", 246, 272, 75, 25, 0)
$Button3 = GUICtrlCreateButton("&Help", 328, 272, 75, 25, 0)
$existingUser_btn = GUICtrlCreateButton("&Search", 25, 272, 75, 25, 0)
GUICtrlSetTip(-1, "Click to Search for Existing User")

GUISetState(@SW_SHOW)
GUICtrlSetOnEvent($createUser_btn,"create_user")


GUICtrlSetOnEvent($button2,"set_exit")
GUICtrlSetOnEvent($button3,"get_help")
GUICtrlSetOnEvent($existingUser_btn,"locate_user")
$homeMDB_combo = GUICtrlCreateCombo("", 24, 184, 361, 25)
GUICtrlSetState(-1,$GUI_HIDE)

    $IADS = ObjGet("LDAP://RootDSE")
    $strDefaultNC = $IADS.Get("defaultnamingcontext")
    $strConfigNC = $IADS.Get("configurationNamingContext")

    ; Open the connection.
    $theConnection = ObjCreate("ADODB.Connection")
    $theCommand = ObjCreate("ADODB.Command")
    $theRecordSet = ObjCreate("ADODB.Recordset")
    $theConnection.Provider = ("ADsDSOObject")
    $theConnection.Open ("ADs Provider")

    ; Build the query to find the private MDBs. Use the first
    ; one if any are found.
    $strQuery = "<LDAP://" & $strConfigNC & ">;(objectCategory=msExchPrivateMDB);name,adspath;subtree"

    $theCommand.ActiveConnection = $theConnection
    $theCommand.CommandText = $strQuery
    $theRecordSet = $theCommand.Execute
    While Not $theRecordSet.EOF


     $firstMDB = String($theRecordSet.Fields("ADsPath").Value)
     GUICtrlSetData($homeMDB_combo,$firstMDB)
     $theRecordSet.Movenext
    WEnd



While 1
    if GUICtrlRead($email_chkbox) = $GUI_CHECKED Then
        GUICtrlSetState($homeMDB_combo,$GUI_SHOW)

 ElseIf GUICtrlRead($email_chkbox) = $GUI_UNCHECKED Then
        GUICtrlSetState($homeMDB_combo,$GUI_HIDE)
    EndIf
    Sleep(100)
WEnd



Func SpecialEvents()

    Select
        Case @GUI_CTRLID = $GUI_EVENT_CLOSE
            Exit

        Case @GUI_CTRLID = $GUI_EVENT_MINIMIZE

        Case @GUI_CTRLID = $GUI_EVENT_RESTORE

    EndSelect

EndFunc

Func set_exit()
 Exit
EndFunc


Func see_displayName()
    GUICtrlSetData($displayName2_input,"")

    $UserObj = ""
    $oMyError = ObjEvent("AutoIt.Error", "")
    Const $ADS_NAME_INITTYPE_GC = 3
    Const $ADS_NAME_TYPE_NT4 = 3
    Const $ADS_NAME_TYPE_1779 = 1

    $oMyError = ObjEvent("AutoIt.Error", "ComError")
    $objRootDSE = ObjGet("LDAP://RootDSE")
    $strNTName = GUICtrlRead($userName2_input)
    If @error Then
     MsgBox(0, 'username', 'Username does not exist or not able to communicate with ' & @LogonDomain)
    Else
    ; DNS domain name.
     $objTrans = ObjCreate("NameTranslate")
     $objTrans.Init ($ADS_NAME_INITTYPE_GC, "")
     ;$objTrans.Set ($ADS_NAME_TYPE_1779, @LogonDomain)
     $objTrans.Set ($ADS_NAME_TYPE_NT4, @LogonDomain & "\" & $strNTName)
     $strUserDN = $objTrans.Get ($ADS_NAME_TYPE_1779)
     $UserObj = ObjGet("LDAP://" & $strUserDN)
      If @error Then
        MsgBox(0, "User Name Error", "Username does not exist in the " & @LogonDomain & " Domain, Please enter a different username")
      Else
        GUICtrlSetData($displayName2_input,$UserObj.givenName & " " & $UserObj.sn)

      EndIf
    EndIf



EndFunc



Func clear_control()
 GUICtrlSetData($templateUser_combo,"")
EndFunc


Func create_User()
 $oMyError = ObjEvent("AutoIt.Error", "ComError")

 If GUICtrlRead($templateUser_combo) = "" AND GUICtrlRead($userName2_input) = "" Then
  MsgBox(64,"Error","Please Enter a Template or User Name")
  Return
 ElseIf GUICtrlRead($templateUser_combo) <> "" AND GUICtrlRead($userName2_input) <> "" Then
  MsgBox(64,"Error","Remove Template OR User Name")
  Return
 ElseIf GUICtrlRead($templateUser_combo) = "" AND GUICtrlRead($userName2_input) <> "" Then
  $strTemplateUser = GUICtrlRead($userName2_input)
  ;MsgBox(64,"username",$strTemplateUser)
 ElseIf GUICtrlRead($templateUser_combo) <> "" AND GUICtrlRead($userName2_input) = "" Then
  $strTemplateUser = GUICtrlRead($templateUser_combo)
  ;MsgBox(64,"template",$strTemplateUser)
 EndIf

 $Progress_GUI = GUICreate("Creating User", 283, 100, -1, -1)

 $GUI_Progress = _GUICtrlProgressMarqueeCreate(50, 25, 200, 20)

 GUISetState()

 _GUICtrlProgressSetMarquee($GUI_Progress)

   $firstName = StringLeft(GUICtrlRead($firstName_input),1)
   $lastName = StringLeft(GUICtrlRead($lastName_input),5)
   $samAccountName = StringLower($firstName & $lastName)

     if _CheckAccount($samAccountName) then
      $counter = 2
      While _CheckAccount($samAccountName,$counter)
       $counter += 1
      WEnd
       $samAccountName &= $counter

     endif

   get_attributes()

    $strDestinationDN = StringSplit($strUserDN, ',')
    $destinationDN = _ArrayToString($strdestinationDN,",",2)

   $objTemplate = ObjGet("LDAP://cn=" & $strTemplateUser & "," & $destinationDN)
   $template = ("cn=" & $strTemplateUser & "," & $destinationDN)
   ;MsgBox(64,"template",$template)
   $strNewUserDN = ("cn=" & $samAccountName & "," & $destinationDN)

   $homeDirectory = $objTemplate.homeDirectory
   If $homeDirectory = "" Then
    MsgBox(0, 'Home Directory Error', 'Unable to set the Home Direcotry correctly, Please manually create the home folder')
   Else

    $string = stringsplit($homeDirectory, '\')
    $fullString= $string[1] & "\" & $string[2] & "\" & $string[3] & "\" &$string[4] & "\" & $string[5]

    ;MsgBox(64,"FullString",$fullstring)
    $home = StringReplace($fullString,$string[5],$samAccountName)
    ;MsgBox(64,"Home",$home)
   EndIf

   Const $ADS_PROPERTY_APPEND = 3
   Const $ADS_UF_NORMAL_ACCOUNT = 512
   Const $createHomeDirectory = TRUE

   $objParent = ObjGet("LDAP://" & $destinationDN)
   $objUser = ObjGet("LDAP://" & $strUserDN)
   $objUser = $objParent.Create("user", "cn=" & $samAccountName)

   $objUser.Put ("samAccountName",$samAccountName)
   $objUser.SetInfo()

   $objUser.SetPassword("Xjdksie!45")
   $objUser.SetInfo()

   $objUser.Put ("userAccountControl", $ADS_UF_NORMAL_ACCOUNT)
   $objUser.AccountDisabled = 0
   $objUser.SetInfo()

   $objUser.Put ("givenName",GUICtrlRead($firstName_input))

   if GUICtrlRead($Input2) <> "" Then
     $objUser.Put ("initials",GUICtrlRead($Input2))
    Else
   EndIf
   $objUser.Put ("sn",GUICtrlRead($lastName_input))
   $objUser.Put ("displayName",GUICtrlRead($lastName_input) & ", " & GUICtrlRead($firstName_input) & " " & GUICtrlRead($Input2))
   $objUser.SetInfo()
    If @error Then
      MsgBox(0, 'Attribute Error', 'Unable to set the Name attributes correctly, Please check manually after ID is finished')
     Else
    EndIf

   $objUser.Put ("userPrincipalName",$samAccountName & "@testdomain.com")
   $objUser.SetInfo()
    If @error Then
      MsgBox(0, 'Attribute Error', 'Unable to set the User Principle Name correctly, Please check manually after ID is finished')
     Else
    EndIf
   $objUser.Put ("department",$objTemplate.department)
   $objUser.SetInfo()
    If @error Then
      MsgBox(0, 'Attribute Error', 'Unable to set the Department attribute correctly, Please check manually after ID is finished')
     Else
    EndIf
   $objUser.Put ("streetAddress",$objTemplate.streetAddress)
   $objUser.Put ("l",$objTemplate.l)
   $objUser.Put ("st",$objTemplate.st)
   $objUser.Put ("c",$objTemplate.c)
   $objUser.Put ("postalCode",$objTemplate.postalCode)
   $objUser.SetInfo()
   If @error Then
     MsgBox(0, 'Attribute Error', 'Unable to set the Address Tab attributes correctly, Please check manually after ID is finished')
    Else
   EndIf
   $objUser.Put ("company",$objTemplate.company)
   $objUser.SetInfo()
   If @error Then
     MsgBox(0, 'Attribute Error', 'Unable to set the Company attribute correctly, Please check manually after ID is finished')
    Else
   EndIf
   $objUser.Put ("physicalDeliveryOfficeName",$objTemplate.physicalDeliveryOfficeName)
   ;$objUser.Put ("description",$objTemplate.description)
   $objUser.SetInfo()
   If @error Then
     MsgBox(0, 'Attribute Error', 'Unable to set the Office attribute correctly, Please check manually after ID is finished')
    Else
   EndIf
   ;$objUser.Put ("homeDrive",$objTemplate.homeDrive)
   ;$objUser.Put ("homeDirectory",$home)

   ;$objUser.SetInfo()
   ;If @error Then
    ; MsgBox(0, 'username', 'Unable to set Home Drive, Please check manually after ID is finished')
    ;Else
   ;EndIf

   ;Create the Home Directory
   $fso = ObjCreate("Scripting.FileSystemObject")
   If Not $fso.FolderExists($home) Then
    $fldUserHomedir = $fso.CreateFolder($home)
   Sleep(5000)

   RunWait(@ComSpec & ' /c cacls.exe "' & $home & '" /T /E /G "' & "testdomain\" & $samAccountName & ':f" ', '', @SW_SHOW)

   EndIf

   $strGroup=$objTemplate.getex("memberof")
    for $strmember in $strGroup
     $objmemberOU = ObjGet("LDAP://" & $strMember)
     $objMemberOU.PutEX($ADS_PROPERTY_APPEND,"member", _ArrayCreate($strNewUserDN))
     $objmemberOU.SetInfo()
    next


  if GUICtrlRead($vpn_btn) = $GUI_CHECKED Then
    $objGroupDN = ObjGet("LDAP://cn=Radius Enabled Dial Access,OU=Groups,OU=US,DC=testdomain,DC=com")
    $objGroupDN.add("LDAP://" & $strNewUserDN)
  EndIf


  if GUICtrlRead($email_chkbox) = $GUI_CHECKED Then
    $LDAPhomeMDB = StringTrimLeft(GUICtrlRead($homeMDB_combo),7)
    ; Build and write the users Exchange attributes
      $objUser.createMailbox($LDAPhomeMDB)
      $objUser.SetInfo
   EndIf

 RunWait("csvde.exe " & "-f " & '"C:\UserCreation\' & $samAccountName & '.csv"' & " -d " & '"dc=testdomain,dc=com"' & " -r " & '"(&(objectClass=user)(cn=' & $samAccountName & '))"')

 MsgBox(64,"User Created", "User ID " & $samAccountName & " has been created")
 $oMyError = ObjEvent("AutoIt.Error", "")
 GUIDelete($progress_GUI)


EndFunc

Func _CheckAccount($samAccountName,$counter = "")

    Local $objCommand = ObjCreate("ADODB.Command")
    Local $objConnection = ObjCreate("ADODB.Connection")

    $objConnection.Provider = "ADsDSOObject"
    $objConnection.Open ("Active Directory Provider")
    $objCommand.ActiveConnection = $objConnection


    Local $strBase = "<LDAP://dc=testdomain,dc=com>"
    Local $strFilter = "(&(objectCategory=person)(objectClass=user)(cn="& $samAccountName & $counter & "))"
    Local $strAttributes = "distinguishedName"
    Local $strQuery = $strBase & ";" & $strFilter & ";" & $strAttributes & ";subtree"
    $objCommand.CommandText = $strQuery
    $objCommand.Properties ("Page Size") = 100
    $objCommand.Properties ("Timeout") = 30
    $objCommand.Properties ("Cache Results") = False

    Local $objRecordSet = $objCommand.Execute

    if $objRecordSet.EOF Then

    Else
     $strDN = $objRecordSet.Fields("distinguishedName").value
     $strDestinationOU = StringSplit($strDN, ',')

     $destinationString = _ArrayToString($strdestinationOU,",",2)

     $objSamAccountName = ObjGet("LDAP://cn=" & $samAccountName & $counter & "," & $destinationstring)


      If IsObj($objSamAccountName) Then
       Return True
      Else
       Return False
      EndIf

    EndIf

    $objConnection.Close
    $objConnection = ""
    $objCommand = ""
    $objRecordSet = ""

EndFunc

Func get_attributes()
    Const $ADS_NAME_INITTYPE_GC = 3
    Const $ADS_NAME_TYPE_NT4 = 3
    Const $ADS_NAME_TYPE_1779 = 1

    If GUICtrlRead($templateUser_combo) = "" AND GUICtrlRead($userName2_input) <> "" Then
     $strNTName = @LogonDomain & "\" & GUICtrlRead($userName2_input)
    ElseIf GUICtrlRead($templateUser_combo) <> "" AND GUICtrlRead($userName2_input) = "" Then
     $strNTName = @LogonDomain & "\" & GUICtrlRead($templateUser_combo)
    EndIf

    $objTrans = ObjCreate("NameTranslate")
    $objTrans.Init ($ADS_NAME_INITTYPE_GC, "")
    $objTrans.Set ($ADS_NAME_TYPE_NT4, $strNTName)
    $strUserDN = $objTrans.Get($ADS_NAME_TYPE_1779)
    $objUser = ObjGet("LDAP://" & $strUserDN)

    $strDestinationDN = StringSplit($strUserDN, ',')
    $destinationDN = _ArrayToString($strdestinationDN,",",2)
EndFunc

;COM Error function
Func ComError()
 If IsObj($oMyError) Then
  $HexNumber = Hex($oMyError.number, 8)
  SetError($HexNumber)
 Else
  SetError(1)
 EndIf
 Return 0
EndFunc ;==>ComError

Func _GUICtrlProgressMarqueeCreate($i_Left, $i_Top, $i_Width = Default, $i_Height = Default)

    Local Const $PBS_MARQUEE = 0x08

    Local $v_Style = $PBS_MARQUEE

    Local $h_Progress = GUICtrlCreateProgress($i_Left, $i_Top, $i_Width, $i_Height, $v_Style)
    If $h_Progress = 0 Then
        SetError(1)
        Return 0
    Else
        SetError(0)
        Return $h_Progress
    EndIf

EndFunc;==>_GUICtrlProgressMarqueeCreate
Func _GUICtrlProgressSetMarquee($h_Progress, $f_Mode = 1, $i_Time = 100)

    Local Const $WM_USER = 0x0400
    Local Const $PBM_SETMARQUEE = ($WM_USER + 10)

    Local $var = GUICtrlSendMsg($h_Progress, $PBM_SETMARQUEE, $f_Mode, Number($i_Time))
    If $var = 0 Then
        SetError(1)
        Return 0
    Else
        SetError(0)
        Return $var
    EndIf

EndFunc;==>_GUICtrlProgressSetMarquee

Func get_Help()
 ;MsgBox(64,"",@LogonDNSDomain & "  " & @LogonDomain)
 ShellExecute('"' & @ScriptDir & '\usercreationHelp.chm"')

EndFunc

Func locate_user()

  $Form1 = GUICreate("Locate Existing User", 517, 326, 193, 115,$DS_SETFOREGROUND,$WS_EX_TOPMOST)
  $locateCancel_btn = GUICtrlCreateButton("&Cancel", 336, 270, 75, 25, 0)
  $locateHelp_btn = GUICtrlCreateButton("&Help", 424, 270, 75, 25, 0)
  $locateSearch_btn = GUICtrlCreateButton("&Search",80,5,50,20)
  GUICtrlSetTip(-1, "Click to Search for User Name")
  $locateUserName_input = GUICtrlCreateInput("", 8, 32, 121, 21)
  GUICtrlSetTip(-1, "Enter User Name")
  $locateUserName_lbl = GUICtrlCreateLabel("User Name", 8, 8, 57, 17)
  $locateLastName_input = GUICtrlCreateInput("", 136, 88, 177, 21)
  $locateFirstName_lbl = GUICtrlCreateLabel("First Name", 8, 64, 54, 17)
  $locateFirstName_input = GUICtrlCreateInput("", 8, 88, 105, 21)
  $locateLastName_lbl = GUICtrlCreateLabel("Last Name", 136, 64, 55, 17)
  $locateADLocation_input = GUICtrlCreateInput("", 72, 136, 393, 21)
  $locateADLocation_lbl = GUICtrlCreateLabel("AD Location", 8, 136, 63, 17)
  $locateOK_btn = GUICtrlCreateButton("&OK", 248, 270, 75, 25, 0)
  $locateTemplate_lbl = GUICtrlCreateLabel("Template", 8, 168, 48, 17)
  $locateTemplates_listbox = GUICtrlCreateList("", 72, 176, 121, 97)
  GUICtrlSetOnEvent($locateCancel_btn, "cancel_form1")
  GUICtrlSetOnEvent($locateHelp_btn, "get_help")
  GUICtrlSetOnEvent($locateSearch_btn, "locate_search")
  GUICtrlSetOnEvent($locateOK_btn, "populate_template")
  GUISwitch($Form1)
  GUISetState(@SW_SHOW,$Form1)

EndFunc

Func cancel_form1()
    ;GUISetState(@SW_SHOW, $GuiWindow)
    GUISwitch($Form1_1)
    GUIDelete($Form1)
EndFunc

Func locate_search()
    $UserObj = ""
    $oMyError = ObjEvent("AutoIt.Error", "")

    GUICtrlSetData($locateTemplates_listbox,"")
    GUICtrlSetData($locateLastName_input,"")
    GUICtrlSetData($locateFirstName_input,"")
    GUICtrlSetData($locateADLocation_input,"")


    Const $ADS_NAME_INITTYPE_GC = 3
    Const $ADS_NAME_TYPE_NT4 = 3
    Const $ADS_NAME_TYPE_1779 = 1

    $oMyError = ObjEvent("AutoIt.Error", "ComError")
    $objRootDSE = ObjGet("LDAP://RootDSE")
    $strNTName = GUICtrlRead($locateUserName_input)

    If $strNTName <> "" Then



     If @error Then
      MsgBox(0, 'username', 'Username does not exist or not able to communicate with ' & @LogonDomain)
     Else


     ; DNS domain name.


      $objTrans = ObjCreate("NameTranslate")
      $objTrans.Init ($ADS_NAME_INITTYPE_GC, "")
      ;$objTrans.Set ($ADS_NAME_TYPE_1779, "testdomain")
      $objTrans.Set ($ADS_NAME_TYPE_NT4, @LogonDomain & "\" & $strNTName)
      $strUserDN = $objTrans.Get ($ADS_NAME_TYPE_1779)
      $UserObj = ObjGet("LDAP://" & $strUserDN)
      $strDestinationDN = StringSplit($strUserDN, ',')
      $destinationDN = _ArrayToString($strdestinationDN,",",2)


         If @error Then
           MsgBox(0, 'User Name Error', 'Username does not exist, Please enter a different username')
         Else


           GUICtrlSetData($locateFirstName_input,$UserObj.givenName)
           GUICtrlSetData($locateLastName_input,$UserObj.sn)
           GUICtrlSetData($locateADLocation_input,$destinationDN)

           Local $objCommand = ObjCreate("ADODB.Command")
           Local $objConnection = ObjCreate("ADODB.Connection")

           $objConnection.Provider = "ADsDSOObject"
           $objConnection.Open ("Active Directory Provider")
           $objCommand.ActiveConnection = $objConnection

           Local $strBase = "<LDAP://" & $destinationDN & ">"
           Local $strFilter = "(&(objectCategory=person)(objectClass=user)(cn=Template*))"
           Local $strAttributes = "cn,samAccountName"
           Local $strQuery = $strBase & ";" & $strFilter & ";" & $strAttributes & ";subtree"



           $objCommand.CommandText = $strQuery
           $objCommand.Properties ("Page Size") = 100
           $objCommand.Properties ("Timeout") = 30
           $objCommand.Properties ("Cache Results") = False

           Local $objRecordSet = $objCommand.Execute

           While Not $objRecordSet.EOF
            $strtemplateCNLocate = $objRecordSet.Fields("samAccountName").value
            GuiCtrlSetData($locateTemplates_listbox,$strTemplateCNLocate)

            $objRecordSet.MoveNext
           Wend

           $objConnection.Close
           $objConnection = ""
           $objCommand = ""
           $objRecordSet = ""
         EndIf
           $UserObj = ""
           $oMyError = ObjEvent("AutoIt.Error", "")
      EndIf
    Else
     MsgBox(0, 'User Name Error', 'Username Field Cannot Be Blank')
    EndIf
   EndFunc

Func populate_template()
 $selectedTemplate = GUICtrlRead($locateTemplates_listbox)
 If $selectedTemplate <> "" Then
  ;MsgBox(64,"",$selectedTemplate)
  GUICtrlSetData($templateUser_combo,$selectedTemplate)
  GUIDelete($Form1)
 Else
  MsgBox(64,"Select A Template","Please Select A Template From The List")
 EndIf
EndFunc
