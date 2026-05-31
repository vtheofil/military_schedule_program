Set oShell = CreateObject("WScript.Shell")
WScript.Sleep 1000
oShell.Run Chr(34) & "C:\Users\theob\Desktop\schedule_project\schedule.exe" & Chr(34) & " --patch", 0,True
oShell.Run Chr(34) & "C:\Users\theob\Desktop\schedule_project\programma_ypiresion_step1.xlsm" & Chr(34), 1, False
