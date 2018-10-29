; Metoden zu Emai Versand: _INetMail() oder _INetSmtpMail()
; optimale Metode von beiden ist _INetSmtpMail()

#include <INet.au3>

$s_SmtpServer = "smtp.funkegruppe.de"
$s_FromName = "Alex Fort"
$s_FromAddress = "deprt042@funkegruppe.de"
$s_ToAddress = "a.fort@funkegruppe.de"
$s_Subject = "My Test UDF"
Dim $as_Body[2]
$as_Body[0] = "Testing the new email udf"
$as_Body[1] = "Second Line"
$Response = _INetSmtpMail ($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body)
$err = @error
If $Response = 1 Then
    MsgBox(0, "Success!", "Mail sent")
Else
    MsgBox(0, "Error!", "Mail failed with error code " & $err)
EndIf