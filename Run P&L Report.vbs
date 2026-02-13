Set WshShell = CreateObject("WScript.Shell")
' Run the batch file hidden (0 = hidden window)
WshShell.Run chr(34) & "Start P&L Report.bat" & chr(34), 0
Set WshShell = Nothing
