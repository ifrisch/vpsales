' Silent Van Paper Update Launcher
' This VBS script runs the automation without any visible windows

Set objShell = CreateObject("WScript.Shell")

' Run the automation wrapper silently (windowstyle 0 = hidden)
objShell.Run """C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\automation_wrapper.bat""", 0, True

' Exit silently
WScript.Quit
