; A set of native PowerShell Cmdlets! This allows you to add the unique features of AutoIt – window manipulation and keystroke simulation
; – to your usual PowerShell scripts. As an additional bonus, the AutoIt PowerShell Cmdlets and Assemblies are digitally signed so they can
; be used with the more strict execution policies. The Cmdlets will also run natively with x86 and x64 versions of PowerShell!

Import the AutoIt PowerShell module
Import-Module .\AutoItX.psd1

# Run notepad.exe
Invoke-AU3Run -Program notepad.exe

# Wait for an untitled notepad window and get the handle
$notepadTitle = "Untitled - Notepad"
Wait-AU3Win -Title $notepadTitle
$winHandle = Get-AU3WinHandle -Title $notepadTitle

# Activate the window
Show-AU3WinActivate -WinHandle $winHandle

# Get the handle of the notepad text control for reliable operations
$controlHandle = Get-AU3ControlHandle -WinHandle $winhandle -Control "Edit1"

# Change the edit control
Set-AU3ControlText -ControlHandle $controlHandle -NewText "Hello! This is being controlled by AutoIt and PowerShell!" -WinHandle $winHandle

# Send some keystrokes to the edit control
Send-AU3ControlKey -ControlHandle $controlHandle -Key "{ENTER}simulate key strokes - line 1" -WinHandle $winHandle
Send-AU3ControlKey -ControlHandle $controlHandle -Key "{ENTER}simulate key strokes - line 2" -WinHandle $winHandle
Send-AU3ControlKey -ControlHandle $controlHandle -Key "{ENTER}{ENTER}" -WinHandle $winHandle