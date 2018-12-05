AutoItSetOption("MustDeclareVars", 1)

ConsoleWrite(Transliteration("Мама мыла раму") & @CRLF)
ConsoleWrite(Transliteration("Рабы не мы, мы не рабы!") & @CRLF)

Exit(0)
;=============================================================================

;=============================================================================
Func Transliteration($sString)
	Static Local $oTranslit = ObjCreate("Scripting.Dictionary")

	With $oTranslit
		If .Count = 0 Then
			.Add("А", "A")
			.Add("Б", "B")
			.Add("В", "V")
			.Add("Г", "G")
			.Add("Д", "D")
			.Add("Е", "E")
			.Add("Ё", "YO")
			.Add("Ж", "ZH")
			.Add("З", "Z")
			.Add("И", "I")
			.Add("Й", "Y'")
			.Add("К", "K")
			.Add("Л", "L")
			.Add("М", "M")
			.Add("Н", "N")
			.Add("О", "O")
			.Add("П", "P")
			.Add("Р", "R")
			.Add("С", "S")
			.Add("Т", "T")
			.Add("У", "U")
			.Add("Ф", "F")
			.Add("Х", "KH")
			.Add("Ц", "TC")
			.Add("Ч", "CH")
			.Add("Ш", "SH")
			.Add("Щ", "SHC")
			.Add("Ъ", "'")
			.Add("Ы", "Y")
			.Add("Ь", "'")
			.Add("Э", "E'")
			.Add("Ю", "YU")
			.Add("Я", "YA")

			.Add("а", "a")
			.Add("б", "b")
			.Add("в", "v")
			.Add("г", "g")
			.Add("д", "d")
			.Add("е", "e")
			.Add("ё", "yo")
			.Add("ж", "zh")
			.Add("з", "z")
			.Add("и", "i")
			.Add("й", "y'")
			.Add("к", "k")
			.Add("л", "l")
			.Add("м", "m")
			.Add("н", "n")
			.Add("о", "o")
			.Add("п", "p")
			.Add("р", "r")
			.Add("с", "s")
			.Add("т", "t")
			.Add("у", "u")
			.Add("ф", "f")
			.Add("х", "kh")
			.Add("ц", "tc")
			.Add("ч", "ch")
			.Add("ш", "sh")
			.Add("щ", "shc")
			.Add("ъ", "'")
			.Add("ы", "y")
			.Add("ь", "'")
			.Add("э", "e'")
			.Add("ю", "yu")
			.Add("я", "ya")
		EndIf
	EndWith

	For $sKey In $oTranslit
		$sString = StringReplace($sString, $sKey, $oTranslit($sKey))
	Next

	Return $sString
EndFunc
;=============================================================================