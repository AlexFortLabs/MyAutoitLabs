; Set DNS


MsgBox(32, "Set DNS 1", _DosRun('netsh interface ipv4 add dns "Local Area Connection" 8.8.8.8'))
MsgBox(32, "Set DNS 2", _DosRun('netsh interface ipv4 add dns "Local Area Connection" 208.67.222.222 index=2'))

Func _DosRun($sCommand)
    Local $nResult = Run('"' & @ComSpec & '" /c ' & $sCommand, @SystemDir, @SW_HIDE, 6)
    ProcessWaitClose($nResult)
    Return StdoutRead($nResult)
EndFunc   ;==>_DosRun
