

CreateMailItem()

Func CreateMailItem()
    Local $olMailItem    = 0
    Local $olFormatRichText = 3
    Local $olImportanceLow   = 0
    Local $olImportanceNormal= 1
    Local $olImportanceHigh  = 2

    $oOApp = ObjCreate("Outlook.Application")
    $oOMail = $oOApp.CreateItem($olMailItem)

;~     With $oOMail
;~         .To = ("to@domain.com")
;~         .Subject = "email subject"
;~         .BodyFormat =  $olFormatRichText
;~         .Importance = $olImportanceHigh
;~         .Body = "email message"
;~         .Display
;~         .Send
;~     EndWith
With $oOMail
	.To = ("a.fortowski@funkegruppe.de")
	.Subject = "Test mail"
	.BodyFormat = $olFormatRichText
	.Importance = $olImportanceHigh
	.Body = "Send VIa script.With attachment"
	.attachments.add (""D:\temp\session.txt"",$olByValue,1,"Text im Anhang")
	.Display
	.Send
EndWith

EndFunc

