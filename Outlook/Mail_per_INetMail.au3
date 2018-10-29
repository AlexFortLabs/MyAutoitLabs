; Metoden zu Emai Versand: _INetMail() oder _INetSmtpMail()
; optimale Metode von beiden ist _INetSmtpMail()
; hier die zweite Metode, verwendet lokalinstaliertes Mailprogramm:

#include <INet.au3>

$Address = InputBox('Address', 'Enter the E-Mail address to send message to')
$Subject = InputBox('Subject', 'Enter a subject for the E-Mail')
$Body = InputBox('Body', 'Enter the body (message) of the E-Mail')
MsgBox(0,'E-Mail has been opened','The E-Mail has been opened and process identifier for the E-Mail client is ' & _INetMail($address, $subject, $body))