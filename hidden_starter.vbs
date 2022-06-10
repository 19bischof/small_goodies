Set  objShell = CreateObject("WScript.Shell")
Dim commands(1) 'param + 1 = size of array
commands(0) = "Powershell -file %USERPROFILE%\Documents\Scripts\pip_upgrade.ps1"
commands(1) = "Powershell Set-Executionpolicy bypass -Scope CurrentUser"
For Each com in commands
    ' MsgBox(com)
    objShell.Run com, 0, False
Next
' MsgBox("theend")