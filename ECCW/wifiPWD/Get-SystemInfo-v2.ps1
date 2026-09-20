#requires -Version 5.1
<#
================================================================================
 INVENTÁRIO TÉCNICO DO WINDOWS - V5
================================================================================
EXECUÇÃO
--------
Execução direta:
    .\system_info_v5_concorrente_sqlite_nativo.ps1

Exemplos:
    .\system_info_v5_concorrente_sqlite_nativo.ps1 -MaxRunspaces 12
    .\system_info_v5_concorrente_sqlite_nativo.ps1 -BrowserItems 100
    .\system_info_v5_concorrente_sqlite_nativo.ps1 -CollectOneDriveShares -IncludeShareLinks

A função principal é Start-SystemInfoCollector.

O arquivo suporta duas formas de uso:
1. Execução direta: .\arquivo.ps1 — executa a coleta automaticamente.
2. Uso como função: . .\arquivo.ps1 — carrega a função sem executar; depois,
   chame Start-SystemInfoCollector com os parâmetros desejados.

PRINCIPAIS PARÂMETROS
---------------------
-MaxRunspaces       Número máximo de tarefas concorrentes.
-TimeoutSeconds     Tempo limite global da coleta.
-RecentFiles        Quantidade máxima de arquivos recentes.
-BrowserItems       Quantidade máxima de registros por navegador.
-DownloadItems      Quantidade máxima de arquivos em Downloads.
-CollectOneDriveShares
                    Habilita a seção de compartilhamentos do OneDrive.
-IncludeShareLinks  Solicita links quando disponíveis por integração autorizada.

OBSERVAÇÕES
-----------
- Utiliza somente PowerShell/.NET e recursos nativos do Windows.
- O SQLite dos navegadores é lido por parser nativo, sem DLL ou programa externo.
- Não coleta senhas, cookies, tokens ou sessões.
- O relatório é gravado em -OutputFile.
================================================================================
#>
#requires -Version 5.1
[CmdletBinding()]
[CmdletBinding()]
param(
# [string]$OutputFile=(Join-Path $env:TEMP 'system_info_v3.txt'),
# Pra salvar em downloads 
[string]$OutputFile=(Join-Path $env:USERPROFILE 'Downloads\system_info_v5.txt'),
 [string]$ExportDir=(Join-Path $env:TEMP 'p'),
 [ValidateRange(1,32)][int]$MaxRunspaces=[Math]::Min([Environment]::ProcessorCount,8),
 [ValidateRange(10,600)][int]$TimeoutSeconds=120,
 [ValidateRange(1,500)][int]$RecentFiles=100,
 [ValidateRange(1,500)][int]$BrowserItems=100,
 [ValidateRange(1,500)][int]$DownloadItems=100,
 [switch]$CollectOneDriveShares,
 [switch]$IncludeShareLinks
)

function Start-SystemInfoCollector {
 [CmdletBinding()]
 param(
  [string]$OutputFile=(Join-Path $env:TEMP 'system_info_v3.txt'),
  [string]$ExportDir=(Join-Path $env:TEMP 'p'),
  [ValidateRange(1,32)][int]$MaxRunspaces=[Math]::Min([Environment]::ProcessorCount,8),
  [ValidateRange(10,600)][int]$TimeoutSeconds=120,
  [ValidateRange(1,500)][int]$RecentFiles=100,
  [ValidateRange(1,500)][int]$BrowserItems=100,
  [ValidateRange(1,500)][int]$DownloadItems=100,
  [switch]$CollectOneDriveShares,
  [switch]$IncludeShareLinks
 )

 $ErrorActionPreference='SilentlyContinue'
 $Started=Get-Date

$ErrorActionPreference='SilentlyContinue'
$Started=Get-Date

# Cada tarefa coleta independentemente e retorna texto; somente o thread principal grava o relatório.
$Tasks=@(
[pscustomobject]@{N='Identificação/OS';O=10;C={
$o=Get-CimInstance Win32_OperatingSystem;$c=Get-CimInstance Win32_ComputerSystem;$r=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion';$p=Get-CimInstance Win32_ComputerSystemProduct
$b=if($r.CurrentBuild){"$($r.CurrentBuild).$($r.UBR)"}else{$o.BuildNumber};$x=@('## 01. IDENTIFICAÇÃO / SISTEMA OPERACIONAL',('-'*78))
$x+="Computador           : $env:COMPUTERNAME";$x+="Usuário              : $env:USERNAME";$x+="Domínio              : $(if($c.PartOfDomain){$c.Domain}else{$c.Workgroup})";$x+="Fabricante           : $($c.Manufacturer)";$x+="Modelo               : $($c.Model)";$x+="UUID                 : $($p.UUID)"
$x+="Produto Windows      : $($r.ProductName)";$x+="Display Version      : $($r.DisplayVersion)";$x+="Release ID           : $($r.ReleaseId)";$x+="Build                : $b";$x+="Versão WMI           : $($o.Version)";$x+="Arquitetura          : $($o.OSArchitecture)";$x+="Idioma               : $($o.MUILanguages -join ', ')";$x+="Instalação           : $($o.InstallDate)";$x+="Último Boot          : $($o.LastBootUpTime)";$x+="Uptime (dias)        : $([math]::Round(((Get-Date)-$o.LastBootUpTime).TotalDays,1))"
$lic=Get-CimInstance SoftwareLicensingProduct|?{$_.PartialProductKey -and $_.LicenseStatus -eq 1}|select -First 1;$x+="Ativação Windows     : $(if($lic){'Licenciado/Ativado'}else{'Não confirmado'})";$x -join "`r`n"
}},
[pscustomobject]@{N='Hardware';O=20;C={
$x=@('## 02. HARDWARE',('-'*78));$bios=Get-CimInstance Win32_BIOS;$cpu=Get-CimInstance Win32_Processor|select -First 1;$ram=Get-CimInstance Win32_PhysicalMemory;$d=Get-CimInstance Win32_DiskDrive;$v=Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3';$pd=Get-PhysicalDisk;$gpu=Get-CimInstance Win32_VideoController;$bad=Get-CimInstance Win32_PnPEntity|?{$_.ConfigManagerErrorCode -and $_.ConfigManagerErrorCode -ne 0};$bat=Get-CimInstance Win32_Battery
$x+="BIOS                : $($bios.Manufacturer) | $($bios.SMBIOSBIOSVersion) | $($bios.ReleaseDate)";$x+="BIOS Serial         : $($bios.SerialNumber)";$x+="CPU                 : $($cpu.Name)";$x+="Cores/Threads       : $($cpu.NumberOfCores)/$($cpu.NumberOfLogicalProcessors)";$x+="Clock               : $($cpu.CurrentClockSpeed)/$($cpu.MaxClockSpeed) MHz"
$x+='';$x+='MEMÓRIA RAM:';$ram|%{$x+="  - $([math]::Round($_.Capacity/1GB,1)) GB | $($_.Speed) MHz | $($_.Manufacturer) | $($_.PartNumber) | $($_.DeviceLocator)"}
$x+='';$x+='DISCOS FÍSICOS:';$d|%{$x+="  - $($_.Model) | $([math]::Round($_.Size/1GB,0)) GB | $($_.InterfaceType) | Serial: $($_.SerialNumber) | FW: $($_.FirmwareRevision)"}
$x+='';$x+='VOLUMES:';$v|%{$u=$_.Size-$_.FreeSpace;$x+="  - $($_.DeviceID) | $($_.FileSystem) | Total: $([math]::Round($_.Size/1GB,2)) GB | Usado: $([math]::Round($u/1GB,2)) GB | Livre: $([math]::Round($_.FreeSpace/$_.Size*100,1))%"}
$x+='';$x+='SAÚDE DOS DISCOS:';$pd|%{$x+="  - $($_.FriendlyName) | $($_.MediaType) | Health: $($_.HealthStatus) | Status: $($_.OperationalStatus)"}
$x+='';$x+='GPU:';$gpu|%{$x+="  - $($_.Name) | Driver: $($_.DriverVersion) | VRAM: $([math]::Round($_.AdapterRAM/1GB,1)) GB | Res: $($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)"}
if($bad){$x+='';$x+='DISPOSITIVOS COM ERRO:';$bad|%{$x+="  - $($_.Name) | Código: $($_.ConfigManagerErrorCode) | $($_.PNPDeviceID)"}}
if($bat){$x+='';$x+='BATERIA:';$bat|%{$x+="  - $($_.Name) | Status: $($_.BatteryStatus) | Design: $($_.DesignCapacity) | Atual: $($_.CurrentCapacity)"}}
$x -join "`r`n"
}},
[pscustomobject]@{N='Rede';O=30;C={
$x=@('## 03. REDE',('-'*78));$a=Get-NetAdapter;$c=Get-NetIPConfiguration;$p=Get-NetConnectionProfile;$pr=Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings';$rt=Get-NetRoute -AddressFamily IPv4|?{$_.NextHop -notin '0.0.0.0','255.255.255.255'}|select -First 30;$n=Get-NetNeighbor|? State -ne Unreachable|select -First 50
$a|%{$x+="  ADP: $($_.Name) | $($_.Status) | MAC: $($_.MacAddress) | Link: $($_.LinkSpeed) | $($_.InterfaceDescription)"}
$c|%{$ip=if($_.IPv4Address){$_.IPv4Address.IPAddress -join ', '}else{'-'};$gw=if($_.IPv4DefaultGateway){$_.IPv4DefaultGateway.NextHop -join ', '}else{'-'};$dns=if($_.DnsServer){$_.DnsServer.ServerAddresses -join ', '}else{'-'};$x+="  IP: $($_.InterfaceAlias) | IPv4: $ip | GW: $gw | DNS: $dns | DHCP: $($_.IPv4DHCPEnabled)"}
$p|%{$x+="  PERFIL: $($_.InterfaceAlias) | $($_.Name) | Categoria: $($_.NetworkCategory) | IPv4: $($_.IPv4Connectivity) | IPv6: $($_.IPv6Connectivity)"};$x+="Proxy               : $(if($pr.ProxyEnable){"Ativo - $($pr.ProxyServer)"}else{'Desativado'})"
$rt|%{$x+="  ROTA: $($_.DestinationPrefix) -> $($_.NextHop) | Métrica: $($_.RouteMetric) | $($_.InterfaceAlias)"};$n|%{$x+="  ARP: $($_.IPAddress) | $($_.MacAddress) | $($_.State) | $($_.InterfaceAlias)"}
$w=netsh wlan show interfaces 2>$null;if($w){$x+='';$x+='WI-FI ATUAL:';$w|sls '^\s*(SSID|BSSID|Radio type|Channel|Receive rate|Transmit rate|Signal|State)\s*:'|%{$x+="  - $($_.Line.Trim())"}}
$wp=netsh wlan show profiles 2>$null|sls '^\s*All User Profile\s*:\s*(.+)'|%{$_.Matches[0].Groups[1].Value.Trim()};if($wp){$x+='';$x+='PERFIS WI-FI SALVOS (somente SSID):';$wp|%{$x+="  - $_"}};$x -join "`r`n"
}},
[pscustomobject]@{N='Segurança';O=40;C={
$x=@('## 04. SEGURANÇA',('-'*78));$fw=Get-NetFirewallProfile;$t=Get-Tpm;$sb=Confirm-SecureBootUEFI;$bl=Get-BitLockerVolume;$av=Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntivirusProduct;$d=Get-MpComputerStatus;$u=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
$x+="Firewall             : $(($fw|%{"$($_.Name): $(if($_.Enabled){'ON'}else{'OFF'})"})-join ', ')";$x+="TPM                 : Presente=$($t.TpmPresent) | Pronto=$($t.TpmReady) | Versão=$($t.ManufacturerVersion)";$x+="Secure Boot          : $(if($sb){'Ativado'}else{'Desativado/Não disponível'})"
$bl|%{$x+="BitLocker            : $($_.MountPoint) | $($_.VolumeStatus) | Proteção: $($_.ProtectionStatus) | $($_.EncryptionMethod)"};$av|%{$x+="Antivírus             : $($_.DisplayName) | Estado: $($_.productState)"};if($d){$x+="Defender              : AV=$($d.AntivirusEnabled) | RT=$($d.RealTimeProtectionEnabled) | Assinatura=$($d.AntivirusSignatureLastUpdated)"};$x+="UAC                  : $(if($u.EnableLUA){'Ativado'}else{'Desativado'})";$x -join "`r`n"
}},
[pscustomobject]@{N='Updates';O=50;C={
$x=@('## 05. WINDOWS UPDATE / PATCHES',('-'*78));$h=Get-HotFix|sort InstalledOn -Desc;$h|select -First 20|%{$x+="  - $($_.HotFixID) | $($_.InstalledOn) | $($_.Description)"}
try{$s=New-Object -ComObject Microsoft.Update.Session;$q=$s.CreateUpdateSearcher();$last=$q.LastSearchSuccessDate;$u=$q.Search('IsInstalled=0').Updates;$x+="Última busca         : $last";$x+="Pendentes            : $($u.Count)";$u|select -First 10|%{$x+="  - $($_.Title)"}}catch{$x+='Windows Update       : consulta indisponível'};$x -join "`r`n"
}},
[pscustomobject]@{N='Software/Drivers';O=60;C={
$x=@('## 06. SOFTWARE / DRIVERS',('-'*78));$a=Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'|? DisplayName|select DisplayName,DisplayVersion,Publisher,InstallDate|sort DisplayName
$x+='PROGRAMAS INSTALADOS:';$a|select -First 100|%{$x+="  - $($_.DisplayName) v$($_.DisplayVersion) | $($_.Publisher) | $($_.InstallDate)"}
$d=Get-CimInstance Win32_PnPSignedDriver|? DeviceName|select DeviceName,DriverVersion,DriverDate,Manufacturer|sort DeviceName;$x+='';$x+='DRIVERS:';$d|select -First 100|%{$x+="  - $($_.DeviceName) | $($_.DriverVersion) | $($_.DriverDate) | $($_.Manufacturer)"};$x -join "`r`n"
}},
[pscustomobject]@{N='Serviços/Startup';O=70;C={
$x=@('## 07. SERVIÇOS / INICIALIZAÇÃO',('-'*78));Get-Service|? Status -eq Running|sort Name|select -First 100|%{$x+="  - $($_.Name) | $($_.DisplayName) | $($_.StartType)"}
$x+='';$x+='INICIALIZAÇÃO (Run):';foreach($p in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run','HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'){$z=Get-ItemProperty $p;if($z){$z.PSObject.Properties|? Name -notmatch '^PS'|%{$x+="  - $($_.Name): $($_.Value)"}}}
$x+='PASTAS STARTUP:';foreach($p in "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup","$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"){Get-ChildItem $p -File|%{$x+="  - $($_.FullName)"}}
$t=Get-ScheduledTask|?{$_.State -ne 'Disabled' -and $_.TaskPath -notlike '\Microsoft\*'};if($t){$x+='';$x+='TAREFAS AGENDADAS (TERCEIROS):';$t|select -First 100|%{$x+="  - $($_.TaskPath)$($_.TaskName) | $($_.State)"}};$x -join "`r`n"
}},
[pscustomobject]@{N='Usuários';O=80;C={
$x=@('## 08. USUÁRIOS / COMPARTILHAMENTOS',('-'*78));$u=Get-LocalUser;$x+='USUÁRIOS LOCAIS:';$u|%{$x+="  - $($_.Name) | Ativo: $($_.Enabled) | Último Login: $($_.LastLogon) | Senha: $($_.PasswordLastSet)"}
$g=Get-LocalGroupMember Administrators;$x+='';$x+='ADMINISTRADORES:';$g|%{$x+="  - $($_.Name) | $($_.ObjectType)"};$s=Get-SmbShare;if($s){$x+='';$x+='COMPARTILHAMENTOS SMB:';$s|%{$x+="  - $($_.Name) | $($_.Path) | $($_.ShareType)"}}
$m=Get-PSDrive -PSProvider FileSystem|? DisplayRoot;if($m){$x+='';$x+='DRIVES MAPEADOS:';$m|%{$x+="  - $($_.Name): -> $($_.DisplayRoot)"}};$x -join "`r`n"
}},
[pscustomobject]@{N='Diagnóstico';O=90;C={
$x=@('## 09. PROCESSOS / DIAGNÓSTICO',('-'*78));Get-Process|sort WorkingSet64 -Desc|select -First 30|%{$cpu=if($_.CPU){[math]::Round($_.CPU,1)}else{0};$x+="  - $($_.ProcessName) | PID: $($_.Id) | Mem: $([math]::Round($_.WorkingSet64/1MB,1)) MB | CPU: ${cpu}s"}
$l=netstat -ano 2>$null|sls LISTENING|select -First 50;if($l){$x+='';$x+='PORTAS EM ESCUTA:';$l|%{$x+="  - $($_.Line.Trim())"}};$e=netstat -ano 2>$null|sls ESTABLISHED|select -First 50;if($e){$x+='';$x+='CONEXÕES ESTABELECIDAS:';$e|%{$x+="  - $($_.Line.Trim())"}}
foreach($log in 'System','Application'){$ev=Get-WinEvent -FilterHashtable @{LogName=$log;Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 50;if($ev){$x+="";$x+="ERROS/CRÍTICOS $log (7 DIAS):";$ev|%{$m=if($_.Message){$_.Message-replace'\s+',' '}else{''};if($m.Length -gt 180){$m=$m.Substring(0,180)};$x+="  - $($_.TimeCreated) | ID $($_.Id) | $($_.ProviderName) | $m"}}}
$d=Get-ChildItem "$env:SystemRoot\Minidump\*.dmp" -File;if($d){$x+='';$x+='MINIDUMPS:';$d|sort LastWriteTime -Desc|select -First 20|%{$x+="  - $($_.FullName) | $($_.LastWriteTime) | $([math]::Round($_.Length/1MB,1)) MB"}};$x -join "`r`n"
}},
[pscustomobject]@{N='Arquivos Recentes';O=100;C={
$x=@('## 10. ARQUIVOS RECENTES',('-'*78));$x+="Limite              : $RecentFiles por categoria";$r=Join-Path $env:APPDATA 'Microsoft\Windows\Recent';$a=Get-ChildItem $r -File|sort LastWriteTime -Desc|select -First $RecentFiles
if($a){$x+='ITENS DA PASTA RECENTES DO WINDOWS:';$a|%{$x+="  - $($_.LastWriteTime) | $($_.Name) | $($_.FullName)"}};$scan=@("$env:USERPROFILE\Desktop","$env:USERPROFILE\Documents","$env:USERPROFILE\Downloads");$f=Get-ChildItem $scan -File -Force -ErrorAction SilentlyContinue|sort LastWriteTime -Desc|select -First $RecentFiles
if($f){$x+='';$x+='ARQUIVOS RECENTEMENTE MODIFICADOS:';$f|%{$x+="  - $($_.LastWriteTime) | $([math]::Round($_.Length/1KB,1)) KB | $($_.FullName)"}};$x -join "`r`n"
}},
[pscustomobject]@{N='Ambiente';O=110;C={
$x=@('## 11. AMBIENTE',('-'*78));foreach($n in 'PATH','APPDATA','PROGRAMDATA','SYSTEMDRIVE','SYSTEMROOT','TEMP','TMP','USERNAME','COMPUTERNAME'){$v=[Environment]::GetEnvironmentVariable($n);if($v){if($v.Length -gt 80){$v=$v.Substring(0,40)+'...'+$v.Substring($v.Length-30)};$x+="  $n = $v"}};$x -join "`r`n"
}}
)

$iss=[initialsessionstate]::CreateDefault()
$pool=[runspacefactory]::CreateRunspacePool(1,$MaxRunspaces,$iss,$Host);$pool.Open()
$pending=@()
foreach($t in $Tasks){
 $ps=[powershell]::Create();$ps.RunspacePool=$pool
 [void]$ps.AddScript({
  param($n,$o,$c);$sw=[Diagnostics.Stopwatch]::StartNew()
  try{$text=&$c;[pscustomobject]@{Name=$n;Order=$o;Status='OK';Seconds=[math]::Round($sw.Elapsed.TotalSeconds,2);Text=$text;Error=''}}
  catch{[pscustomobject]@{Name=$n;Order=$o;Status='ERRO';Seconds=[math]::Round($sw.Elapsed.TotalSeconds,2);Text="[$n] falhou: $($_.Exception.Message)";Error=$_.Exception.ToString()}}
 }).AddArgument($t.N).AddArgument($t.O).AddArgument($t.C)
 $pending+=[pscustomobject]@{T=$t;P=$ps;A=$ps.BeginInvoke()}
}
$results=@();$deadline=(Get-Date).AddSeconds($TimeoutSeconds)
while($pending.Count){
 foreach($j in @($pending)){
  if($j.A.IsCompleted){
   try{$results+=$j.P.EndInvoke($j.A)}catch{$results+=[pscustomobject]@{Name=$j.T.N;Order=$j.T.O;Status='ERRO';Seconds=0;Text="[$($j.T.N)] falhou ao finalizar: $($_.Exception.Message)";Error=$_.Exception.ToString()}}
   $j.P.Dispose();$pending=@($pending|?{$_ -ne $j})
  }
 }
 if($pending.Count -and (Get-Date) -gt $deadline){
  foreach($j in @($pending)){$j.P.Stop();$results+=[pscustomobject]@{Name=$j.T.N;Order=$j.T.O;Status='TIMEOUT';Seconds=$TimeoutSeconds;Text="[$($j.T.N)] timeout após $TimeoutSeconds segundos.";Error='Timeout'};$j.P.Dispose()};$pending=@()
 }else{Start-Sleep -Milliseconds 50}
}
$pool.Close();$pool.Dispose()

$results=$results|sort Order;$sb=[Text.StringBuilder]::new();$sep='='*78
[void]$sb.AppendLine($sep);[void]$sb.AppendLine('INVENTÁRIO TÉCNICO COMPLETO DO SISTEMA - V3 CONCORRENTE');[void]$sb.AppendLine("Coletado em: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')");[void]$sb.AppendLine("Workers: $MaxRunspaces | Timeout: $TimeoutSeconds s");[void]$sb.AppendLine($sep)
foreach($r in $results){[void]$sb.AppendLine($r.Text);[void]$sb.AppendLine('')}
[void]$sb.AppendLine($sep);[void]$sb.AppendLine('## TEMPO / STATUS DAS COLETAS');[void]$sb.AppendLine(('-'*78));$results|%{[void]$sb.AppendLine(('{0,-22} | {1,-8} | {2,7}s' -f $_.Name,$_.Status,$_.Seconds))}
[void]$sb.AppendLine("Tempo total: $([math]::Round(((Get-Date)-$Started).TotalSeconds,2)) s");[void]$sb.AppendLine($sep)
if(!(Test-Path $ExportDir -PathType Container)){New-Item $ExportDir -ItemType Directory -Force|Out-Null}
$sb.ToString()|Out-File $OutputFile -Encoding UTF8 -Force
Write-Host "Relatório salvo em: $OutputFile";Write-Host "Tempo total: $([math]::Round(((Get-Date)-$Started).TotalSeconds,2)) s";Write-Host "Workers: $MaxRunspaces | Timeout: $TimeoutSeconds s"

}

# EXECUÇÃO DIRETA:
# Quando o arquivo é chamado normalmente, a função é executada automaticamente.
# Quando o arquivo é dot-sourced ('. .\arquivo.ps1'), somente a função é definida,
# permitindo chamá-la manualmente com qualquer combinação de parâmetros.
if ($MyInvocation.InvocationName -ne '.') {
 Start-SystemInfoCollector `
  -OutputFile $OutputFile `
  -ExportDir $ExportDir `
  -MaxRunspaces $MaxRunspaces `
  -TimeoutSeconds $TimeoutSeconds `
  -RecentFiles $RecentFiles `
  -BrowserItems $BrowserItems `
  -DownloadItems $DownloadItems `
  -CollectOneDriveShares:$CollectOneDriveShares `
  -IncludeShareLinks:$IncludeShareLinks
}
