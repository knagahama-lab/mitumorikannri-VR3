' run_hidden.vbs
' Run a PowerShell script without showing any console window.
' Task Scheduler starts: wscript.exe "run_hidden.vbs" "C:\...\script.ps1" [args...]
' (powershell.exe -WindowStyle Hidden still flashes a window for a moment; wscript has no window at all)
Option Explicit
Dim sh, cmd, i
If WScript.Arguments.Count = 0 Then WScript.Quit 1
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & WScript.Arguments(0) & """"
For i = 1 To WScript.Arguments.Count - 1
  cmd = cmd & " """ & WScript.Arguments(i) & """"
Next
Set sh = CreateObject("WScript.Shell")
WScript.Quit sh.Run(cmd, 0, True)
