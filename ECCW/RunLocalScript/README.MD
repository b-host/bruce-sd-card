# Rodar Scripts Localmente
## Windows
### Versão Longa
```BASH

Delay 1000
GuiR
WinPrintLine powershell Start-Process powershell -Verb runAs
Press KEY_LEFT_ALT
Press Y
Release
Delay 2000
Press KEY_LEFT_ARROW
PressRelease KEY_ENTER
Release
Delay 1000

## WinPrintLine Set-MpPreference -DisableRealtimeMonitoring $true 
##WinPrintLine Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

## Defina Local do arquivo
WinPrintLine Get-Content Z:ECCW\1.ps1 | PowerShell.exe -noprofile -
```

### Versão Longa

Pode ser que se solicite alteracao na mudanca de politica de execucao, para burlar isso, faca :

```powershell
Delay 1000
GuiR
WinPrintLine powershell -Command "Start-Process powershell -Verb RunAs -ArgumentList '-NoExit -Command Get-Content Z:ECCW\main.ps1 | PowerShell.exe -noprofile -'"
Press KEY_LEFT_ALT
Press Y
Release
Delay 2000
Press KEY_LEFT_ARROW
PressRelease KEY_ENTER
Release
Delay 1000
```