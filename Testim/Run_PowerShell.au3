#include <Constants.au3>
#include <Encoding.au3>



Func _power($comand)

	local $PowerShell = Run(@ComSpec & " /c powershell " & $comand, @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	local $out="", $err=""
	StdinWrite($PowerShell)
	While 1
		$line = StdoutRead($PowerShell)
		If @error Then ExitLoop
		If $line <> "" Then $out&=_Encoding_OEM2ANSI($line)
	Wend


	if $out="" then
		$out=""
		While 1
			$err = StderrRead($PowerShell)
			If @error Then ExitLoop
			If $err <> "" Then $out&=_Encoding_OEM2ANSI($err)

		Wend
	endif

	If $PowerShell Then ProcessClose($PowerShell)
	if $out <> "" then Return $out

EndFunc
