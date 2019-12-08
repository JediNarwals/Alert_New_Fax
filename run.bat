@echo off
:loop
powershell.exe -window hidden -NoExit "H:\Stuff\Alert_New_Fax.ps1"
timeout /t 2 /nobreak
cmd taskkill /IM powershell.exe /f
timeout /t 2 /nobreak
goto loop