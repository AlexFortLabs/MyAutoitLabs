; Testen
; zabbix Agenten verteilen an mehrere Rechner


#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIError.Au3>
#include <Array.au3>
#include <File.au3>
#include <date.au3>
#include <Inet.au3>
Global $password = ""
Global $pass = ""
; range definition
$i = 2  ;this is the store. Each store get's it's own C-Class range.
    For $j = 1 To 255 Step 1
        $remoteIP = "10.2." & $i & "." & $j
        If ping($remoteIP, 500) = 0 Then
            If @error = 1 Then $msg = "Host is offline"
            If @error = 2 Then $msg = "Host is unreachable"
            If @error = 3 Then $msg = "Bad destination"
            If @error = 4 Then $msg = "Other errors"
            ConsoleWrite($remoteIP & ": " & $msg & @CRLF)
        Else
            Remote_Exec($remoteIP, "********")  ; password 1
            Remote_Exec($remoteIP, "$$$$$$$$")  ; password 2
            Remote_Exec($remoteIP, "%%%%%%%%")  ; password 3
        EndIf
    Next


Func Remote_Exec($remoteIP, $password)
    Global $share = "\\" & $remoteIP & "\c$"
    Local $tmp = ""
    $tmp = DriveMapAdd("o:", $share, 0, $remoteIP & "\administrator", $password)
    If $tmp = 1 Then
        FileCopy("zabbix_agent-3.2.10_" & @CPUArch & ".msi", "o:\zabbix", 8)
                $cmd = "msiexec /I zabbix_agent-3.2.11_" & @CPUArch & ".msi HOSTNAME=" & _TCPIpToName($remoteIP) & _
                        " HOSTNAMEFQDN=1 SERVER=152.141.32.212 SERVERACTIVE=152.141.32.212 RMTCMD=1 /qn"

        FileWriteLine("o:\zabbix\run.cmd", $cmd)
        $pid = WMI_Remote_execute_Authenticated($remoteIP, $password, $share & "\zabbix\run.cmd")
        If $pid <> 0 Then ConsoleWrite("MSI started on: " & _TCPIpToName($remoteIP) & "(" & $remoteIP & ") with PID: " & $pid & " and password (" & $password & ")" & @CRLF)
    EndIf
    $pass = ""
    DriveMapDel("o:")

EndFunc

Func WMI_Remote_execute_Authenticated($strComputer, $Password, $strCommand)
    $UserName = $strComputer & "\administrator"
    $SWBemlocator = ObjCreate("WbemScripting.SWbemLocator")
    $objWMIService = $SWBemlocator.ConnectServer($strComputer,"root\CIMV2",$UserName,$Password)
    $objShare = $objWMIService.Get("Win32_Process")
    $objInParam = $objShare.Methods_("Create").inParameters.SpawnInstance_()

    $objInParam.Properties_.Item("CommandLine") =  $strCommand
    $objOutParams = $objWMIService.ExecMethod("Win32_Process", "Create", $objInParam)
    If $objOutParams.ReturnValue = 0 Then Return $objOutParams.ProcessId
EndFunc

Func LogWrite($Msg)
    FileWriteLine(@ScriptDir & "\store.log", _NowDate() & " " & _NowTime() & ": " & $Msg)
EndFunc
