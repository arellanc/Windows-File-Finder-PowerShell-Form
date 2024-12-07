Set objShell = CreateObject("Shell.Application")
' Construct the path to the PowerShell script
scriptPath = "C:\Users\$env:UserName\Documents\File Finder Project\Launch Find your Files.ps1"
' Execute PowerShell without elevation and hide the window
objShell.ShellExecute "powershell.exe", "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """", "", "", 0
