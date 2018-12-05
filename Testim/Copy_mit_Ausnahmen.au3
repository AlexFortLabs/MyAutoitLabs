; kak iskljuchit' fajly opredelennogo tipa pri kopirovanii (naprimer *.zip) ?

#NoTrayIcon
#include <Array.au3>
#Include <FileOperations.au3>

$sPath = "E:\test\source"
if Not FileExists($sPath) then
MsgBox(64, "Fehler", "Nicht gefunden Ordner: " & @CRLF & $sPath)
Exit
EndIf

$dPath = "C:\test\destination"
if Not FileExists($dPath) then
DirCreate($dPath)
EndIf

$Level = 5                     ; Ordnertiefe
$Exclude = "*.zip|*.7z"        ; alle Erweiterungen die auszunehmen sind aus dem Copiervorgang zip und 7z



$a1FileList = _FO_FileSearch($sPath, $Exclude, False, 0, 1, 1, 0)
For $i=1 To $a1FileList[0]
FileCopy($a1FileList[$i], $dPath)
Next

$aFolderList = _FO_FolderSearch($sPath, "*", True, $Level, 0, 1, 0)
For $i=1 To $aFolderList[0]
DirCreate($dPath & "\" & $aFolderList[$i])

    $a2FileList = _FO_FileSearch($sPath & "\" & $aFolderList[$i], $Exclude, False, 0, 1, 1, 0)
    For $j=1 To $a2FileList[0]
    FileCopy($a2FileList[$j], $dPath & "\" & $aFolderList[$i])
    Next

Next

Exit

; Eshhe mozhno proverjat' nalichie v papke fajlov s opredelennym rasshireniem (esli, konechno, v jetoj papke net vlozhennyh papok).

#cs
Local $sDir = @ScriptDir & '\Google_Translator', $sExclude = '*.zip|*.7z', $aExclude, $fExist

$aExclude = StringSplit($sExclude, '|')

;~ ...
$fExist = False
For $i = 1 To $aExclude[0]
    If FileExists($sDir & '\' & $aExclude[$i]) Then
        $fExist = True
        ExitLoop
    EndIf
Next
ConsoleWrite($sExclude & ' -> ' & $fExist & @LF)
If $fExist Then
;~     ... dein Variant
Else
;~     DirCopy() oder andere Methode
EndIf

#ce