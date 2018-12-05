; I know how to send one line of text to the clipboard:

ClipPut("Mehrzeiligen Text" & @CRLF &  "hast du jetzt" & @CRLF & "in deiner" & @crlf & "Zwischenablage")
MsgBox(0,"",ClipGet())

