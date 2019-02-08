; List computers accounts on Active Directory OU

#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <AD.au3>
#include <TreeviewConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>
#include <GuiTreeView.au3>

Global $sTitle = "Active Direcory OU Treeview"
#region ### START Koda GUI section ### Form=
Global $hMain = GUICreate($sTitle, 743, 683, -1, -1)
Global $hTree = GUICtrlCreateTreeView(6, 6, 600, 666, -1, $WS_EX_CLIENTEDGE)
Global $bExit = GUICtrlCreateButton("Exit", 624, 8, 97, 33)
Global $bExpand = GUICtrlCreateButton("Expand", 624, 56, 97, 33)
Global $bCollapse = GUICtrlCreateButton("Collapse", 624, 104, 97, 33)
Global $bSelect = GUICtrlCreateButton("Select OU", 624, 152, 97, 33)
#endregion ### END Koda GUI section ###

;------------------------------
; Make your changes below
;------------------------------
_AD_Open()
Global $sOU = _AD_SamAccountNameToFQDN(@ComputerName & "$")
Global $iPos = StringInStr($sOU, ",")
$sOU = StringMid($sOU, $iPos + 1) ; FQDN of the OU where to start

;Global $sCategory = "computer" ; Category to select when counts should be displayed or full LDAP query
; If you just want to display all computers with e.g Windows 7 modify the following statement:
Global $sCategory = "(&(objectcategory=computer)(operatingSystem=Windows 7*))" ; Category to select when counts should be displayed or full LDAP query


;Global $sText = " (Computer: %)" ; Text to use for the counts to display. % is replaced with the count
; If you just want to display all .... modify the following statement:
Global $sText = " (Laptops: %)" ; Text to use for the counts to display. % is replaced with the count

Global $iScope = 1 ; Scope for the LDAP search to calculate the count. 1 = Only the OU, 2 = OU + sub-OUs
Global $bCount = True ; 1 = Display the object count right to the OU
Global $bDisplay = True ; True = Display objects in the OU as well (selection = $sCategory)
;------------------------------
; Make your changes above
;------------------------------
Global $aTreeView = _AD_GetOUTreeView($sOU, $hTree, True, $bCount, $bDisplay, $sCategory, $sText, $iScope)
If @error <> 0 Then MsgBox(16, "Active Direcory OU Treeview", "Error creating list of OUs starting with '" & $sOU & "'." & @CRLF & _
        "Error returned by function _AD_GetALLOUs: @error = " & @error & ", @extended =  " & @extended)
GUISetState(@SW_SHOW)

While 1
    $Msg = GUIGetMsg()
    Switch $Msg
        Case $GUI_EVENT_CLOSE, $bExit
            _AD_Close()
            Exit
        Case $bExpand
            _GUICtrlTreeView_Expand($hTree)
        Case $bCollapse
            _GUICtrlTreeView_Expand($hTree, 0, False)
        Case $bSelect
            $hSelection = _GUICtrlTreeView_GetSelection($hTree)
            $sSelection = _GUICtrlTreeView_GetText($hTree, $hSelection)
            For $i = 1 To $aTreeView[0][0]
                If $hSelection = $aTreeView[$i][2] Then ExitLoop
            Next
            $sOU = $aTreeView[$i][1]
            MsgBox(64, $sTitle & " - Selected OU", "Name: " & $sSelection & @CRLF & "FQDN: " & $sOU)
    EndSwitch
WEnd

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetOUTreeView
; Description ...: Fills a given TreeView control with all OUs starting with a given OU.
; Syntax.........: _AD_GetOUTreeView($sAD_OU, $hAD_TreeView[, $bAD_IsADOpen = False[, $bAD_Count = False[, $bAD_Display = False[, $sAD_Category = ""[, $sAD_Text = " (%)"[, $iAD_SearchScope = 1]]]]]])
; Parameters ....: $sAD_OU        - FQDN of the OU where to start. "" displays all OUs
;                 $hAD_TreeView    - Handle of the TreeView where the OUs will be displayed. Has to be created in advance
;                 $bAD_IsADOpen    - Optional: Pass as True if the connection to the AD has already been made by the calling script (default = False)
;                 $bAD_Count       - Optional: True will display the count of objects in the OU right to the OU name (default = False)
;                                    The LDAP query to determine the count is built from $sAD_Category
;                 $bAD_Display   - Optional: True will display the objects in the OU below the OU itself (default = False)
;                 $sAD_Category    - Optional: Category to search for or the complete LDAP search string e.g. "computer" and "(objectcategory=computer)" are equivalent
;                                    The number of found objects is added to the OUs name (default = "" = don't count objects)
;                 $sAD_Text     - Optional: Text to display the count. Must contain a % which will be replaced by the actual number (default = " (%)")
;                 $iAD_SearchScope - Optional: Search scope for function _AD_GetObjectsInOU: 0 = base, 1 = one-level, 2 = sub-tree (default = 1)
; Return values .: Success - 1
;                 Failure - Returns 0 and sets @error:
;                 |x - Function _AD_Open or _AD_GetAllOUs failed. @error and @extended are set by _AD_Open or _AD_GetAllOUs
; Author ........: water bnased on code by ReFran - http://www.autoitscript.com/forum/topic/84119-treeview-read-to-from-file
; Modified.......: water including ideas by HaeMHuK
; Remarks .......: Use $iAD_SearchScope = 1 to get the number of objects of a single OU.
;                 Use $iAD_SearchScope = 2 to get the number of objects in the OU plus all sub-OUs.
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AD_GetOUTreeView($sAD_OU, $hAD_TreeView, $bAD_IsADOpen = False, $bAD_Count = False, $bAD_Display = False, $sAD_Category = "", $sAD_Text = " (%)", $iAD_SearchScope = 1)

    Local $iAD_Count, $aObjects
    If $bAD_IsADOpen = False Then
        _AD_Open()
        If @error Then Return SetError(@error, @extended, 0)
    EndIf
    $sSeparator = ""
    Local $aAD_OUs = _AD_GetAllOUs($sAD_OU, $sSeparator)
    If @error <> 0 Then Return SetError(@error, @extended, 0)
    Local $aAD_TreeView[$aAD_OUs[0][0] + 1][3] = [[$aAD_OUs[0][0], 3]]
    Local $aAD_Result = $aAD_TreeView
    For $i = 1 To $aAD_OUs[0][0]
        $aAD_Temp = StringSplit($aAD_OUs[$i][0], $sSeparator)
        $aAD_TreeView[$i][0] = StringFormat("%" & $aAD_Temp[0] - 1 & "s", "") & "#" & $aAD_Temp[$aAD_Temp[0]]
        $aAD_TreeView[$i][1] = $aAD_OUs[$i][1]
    Next

    _GUICtrlTreeView_BeginUpdate($hAD_TreeView)
    Local $ahAD_Node[50], $sAD_LDAPString
    If $sAD_Category <> "" And StringIsAlNum($sAD_Category) Then
        $sAD_LDAPString = "(objectcategory=" & $sAD_Category & ")"
    Else
        $sAD_LDAPString = $sAD_Category
    EndIf
    $iAD_OutIndex = 0
    For $iAD_Index = 1 To $aAD_TreeView[0][0]
        $iAD_OutIndex += 1
        $aObjects = ""
        $sAD_Line = StringSplit(StringStripCR($aAD_TreeView[$iAD_Index][0]), @TAB)
        $iAD_Level = StringInStr($sAD_Line[1], "#")
        If $iAD_Level = 0 Then ExitLoop
        If ($bAD_Count Or $bAD_Display) And $sAD_LDAPString <> "" Then
            If $bAD_Display Then
                $aObjects = _AD_GetObjectsInOU($aAD_TreeView[$iAD_Index][1], $sAD_LDAPString, $iAD_SearchScope, "samaccountname,distinguishedname", "", False)
                If @error Then
                    $iAD_Count = 0
                Else
                    $iAD_Count = $aObjects[0][0]
                EndIf
            Else
                $iAD_Count = _AD_GetObjectsInOU($aAD_TreeView[$iAD_Index][1], $sAD_LDAPString, $iAD_SearchScope, "samaccountname,distinguishedname", "", True)
            EndIf
        EndIf
        $sAD_Temp = ""
        If $bAD_Count And $sAD_LDAPString <> "" Then $sAD_Temp = StringReplace($sAD_Text, "%", $iAD_Count)
        If $iAD_Level = 1 Then
            $ahAD_Node[$iAD_Level] = _GUICtrlTreeView_Add($hAD_TreeView, 0, StringMid($sAD_Line[1], $iAD_Level + 1) & $sAD_Temp)
        Else
            $ahAD_Node[$iAD_Level] = _GUICtrlTreeView_AddChild($hAD_TreeView, $ahAD_Node[$iAD_Level - 1], StringMid($sAD_Line[1], $iAD_Level + 1) & $sAD_Temp)
        EndIf
        $aAD_Result[$iAD_OutIndex][0] = $aAD_TreeView[$iAD_Index][0]
        $aAD_Result[$iAD_OutIndex][1] = $aAD_TreeView[$iAD_Index][1]
        $aAD_Result[$iAD_OutIndex][2] = $ahAD_Node[$iAD_Level]
        ; Add the objects of the OU to the TreeView
        If IsArray($aObjects) Then
            ReDim $aAD_Result[UBound($aAD_Result, 1) + $aObjects[0][0]][UBound($aAD_Result, 2)]
            For $iAD_Index2 = 1 To $aObjects[0][0]
                $iAD_OutIndex += 1
                If StringRight($aObjects[$iAD_Index2][0], 1) = "$" Then $aObjects[$iAD_Index2][0] = StringLeft($aObjects[$iAD_Index2][0], StringLen($aObjects[$iAD_Index2][0]) - 1)
                $aAD_Result[$iAD_OutIndex][0] = $aObjects[$iAD_Index2][0]
                $aAD_Result[$iAD_OutIndex][1] = $aObjects[$iAD_Index2][1]
                $aAD_Result[$iAD_OutIndex][2] = _GUICtrlTreeView_AddChild($hAD_TreeView, $ahAD_Node[$iAD_Level], $aObjects[$iAD_Index2][0])
            Next
        EndIf
    Next
    If $bAD_IsADOpen = False Then _AD_Close()
    _GUICtrlTreeView_EndUpdate($hAD_TreeView)
    $aAD_Result[0][0] = UBound($aAD_Result, 1) - 1
    $aAD_Result[0][1] = UBound($aAD_Result, 2)
    Return $aAD_Result

EndFunc   ;==>_AD_GetOUTreeView