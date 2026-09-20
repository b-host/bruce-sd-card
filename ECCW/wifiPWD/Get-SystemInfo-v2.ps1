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

A função principal é Start-SystemInfoCollector. A chamada ao final do arquivo
executa a coleta automaticamente quando o .ps1 é iniciado.

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

param(
 [string]$OutputFile=(Join-Path $env:USERPROFILE 'Downloads\system_info_v5.txt'),
 [ValidateRange(1,32)][int]$MaxRunspaces=[Math]::Min([Environment]::ProcessorCount,8),
 [ValidateRange(10,600)][int]$TimeoutSeconds=120,
 [ValidateRange(1,1000)][int]$RecentFiles=100,
 [ValidateRange(1,1000)][int]$BrowserItems=100,
 [ValidateRange(1,2000)][int]$DownloadItems=500,
 [switch]$CollectOneDriveShares,
 [switch]$IncludeShareLinks,
 [string]$LogFile=(Join-Path $env:USERPROFILE 'Downloads\system_info_v5_debug.log'),
 [switch]$DebugMode
)

function Write-BootstrapLog([string]$Message){
 try{
  $p=Join-Path $env:USERPROFILE 'Downloads\system_info_v5_bootstrap.log'
  Add-Content -LiteralPath $p -Value ("{0} [BOOT] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$Message) -Encoding UTF8
 }catch{}
}
Write-BootstrapLog "Arquivo carregado: $PSCommandPath"
Write-BootstrapLog "PowerShell: $($PSVersionTable.PSVersion)"
function Start-SystemInfoCollector {
 [CmdletBinding()]
 param(
  [string]$OutputFile=(Join-Path $env:USERPROFILE 'Downloads\system_info_v5.txt'),
  [ValidateRange(1,32)][int]$MaxRunspaces=[Math]::Min([Environment]::ProcessorCount,8),
  [ValidateRange(10,600)][int]$TimeoutSeconds=120,
  [ValidateRange(1,1000)][int]$RecentFiles=100,
  [ValidateRange(1,1000)][int]$BrowserItems=100,
  [ValidateRange(1,2000)][int]$DownloadItems=500,
  [switch]$CollectOneDriveShares,
  [switch]$IncludeShareLinks
 )
 $ErrorActionPreference='Continue';$Started=Get-Date
 $script:LogFile=$LogFile
 function Write-CollectorLog([string]$Message,[string]$Level='INFO'){
  $line="{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$Level,$Message
  try{Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop}catch{}
  if($DebugMode){Write-Host $line}
 }
 try{
  $lp=Split-Path -Parent $script:LogFile
  if($lp -and !(Test-Path $lp)){New-Item -ItemType Directory -Path $lp -Force -ErrorAction Stop|Out-Null}
  Set-Content -LiteralPath $script:LogFile -Value '' -Encoding UTF8 -ErrorAction Stop
 }catch{}
 Write-CollectorLog "INICIO | PID=$PID | PS=$($PSVersionTable.PSVersion) | Arquivo=$PSCommandPath"
 Write-CollectorLog "PARAMETROS | Workers=$MaxRunspaces Timeout=$TimeoutSeconds BrowserItems=$BrowserItems RecentFiles=$RecentFiles Downloads=$DownloadItems"

function L([string]$s){$s}
function H([string]$n){@("## $n",('-'*78))}
function Get-SqliteVarInt([byte[]]$b,[ref]$i){
 [UInt64]$v=0
 for($n=0;$n -lt 9;$n++){
  if($i.Value -ge $b.Length){throw 'SQLite varint truncado'}
  $x=$b[$i.Value];$i.Value++
  if($n -eq 8){$v=($v -shl 8)-bor$x;break}
  $v=($v-shl 7)-bor($x -band 0x7f);if(($x -band 0x80) -eq 0){break}
 }
 $v
}
function Get-SqliteRecord([byte[]]$p){
 $i=0;$hs=Get-SqliteVarInt $p ([ref]$i);$ts=@();while($i -lt $hs){$ts+=Get-SqliteVarInt $p ([ref]$i)}
 $d=@();foreach($t in $ts){switch($t){
  0{$d+=,$null};1{$d+=[sbyte]$p[$i];$i++};2{$d+=[int16](($p[$i]-shl 8)-bor$p[$i+1]);$i+=2}
  3{$d+=[int32](($p[$i]-shl 16)-bor($p[$i+1]-shl 8)-bor$p[$i+2]);$i+=3};4{$d+=[int32](($p[$i]-shl 24)-bor($p[$i+1]-shl 16)-bor($p[$i+2]-shl 8)-bor$p[$i+3]);$i+=4}
  5{$v=[int64]0;1..5|ForEach-Object {$v=($v -shl 8)-bor$p[$i];$i++};$d+=$v};6{$v=[int64]0;1..8|ForEach-Object {$v=($v -shl 8)-bor$p[$i];$i++};$d+=$v}
  7{$d+=[BitConverter]::ToDouble([byte[]]$p[$i..($i+7)],0);$i+=8};8{$d+=0};9{$d+=1}
  default{if($t -ge 12){$n=[int](($t-12)/2);$raw=if($n){[byte[]]$p[$i..($i+$n-1)]}else{[byte[]]@()};if(($t % 2) -eq 0){$d+=$raw}else{$d+=[Text.Encoding]::UTF8.GetString($raw)};$i+=$n}}
 }}; $d
}
function Get-SqlitePage([IO.FileStream]$fs,[int]$pg,[int]$ps){
 $b=New-Object byte[] $ps;$fs.Position=([int64]($pg-1)*$ps);if($fs.Read($b,0,$ps)-ne$ps){throw 'Página SQLite incompleta'};$b
}
function Get-SqliteRows([string]$Path,[string]$Table){
 $fs=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 try{
  $h=New-Object byte[] 100;$null=$fs.Read($h,0,100)
  if([Text.Encoding]::ASCII.GetString($h,0,16)-ne('SQLite format 3'+[char]0)){throw 'SQLite inválido'}
  $ps=($h[16]-shl 8)-bor$h[17];if($ps-eq1){$ps=65536}
  $schema=@();$scan=$null
  $scan={param([int]$pg)
   $p=Get-SqlitePage $fs $pg $ps;$o=if($pg -eq 1){100}else{0};$typ=$p[$o];$cnt=($p[$o+3]-shl 8)-bor$p[$o+4];$ptr=$o+8
   if($typ -eq 5){
    for($k=0;$k -lt $cnt;$k++){$co=($p[$ptr]-shl 8)-bor$p[$ptr+1];$ch=($p[$co]-shl 24)-bor($p[$co+1]-shl 16)-bor($p[$co+2]-shl 8)-bor$p[$co+3];&$scan $ch;$ptr+=2}
    $ch=($p[$o+8]-shl 24)-bor($p[$o+9]-shl 16)-bor($p[$o+10]-shl 8)-bor$p[$o+11];&$scan $ch
   }elseif($typ -eq 13){
    for($k=0;$k -lt $cnt;$k++){$co=($p[$ptr]-shl 8)-bor$p[$ptr+1];$q=$co;$len=Get-SqliteVarInt $p ([ref]$q);$null=Get-SqliteVarInt $p ([ref]$q);$n=[int]$len;if($q + $n -gt $ps){$n=$ps-$q};$schema+=,(Get-SqliteRecord $(if($n){[byte[]]$p[$q..($q+$n-1)]}else{[byte[]]@()}));$ptr+=2}
   }
  };&$scan 1
  $m=$schema|Where-Object {$_[0]-eq'table' -and $_[1]-eq$Table}|Select-Object -First 1;if(!$m){return @()};$root=[int]$m[3];$out=@();$read=$null
  $read={param([int]$pg)
   $p=Get-SqlitePage $fs $pg $ps;$o=if($pg -eq 1){100}else{0};$typ=$p[$o];$cnt=($p[$o+3]-shl 8)-bor$p[$o+4];$ptr=$o+8
   if($typ -eq 5){
    for($k=0;$k -lt $cnt;$k++){$co=($p[$ptr]-shl 8)-bor$p[$ptr+1];$ch=($p[$co]-shl 24)-bor($p[$co+1]-shl 16)-bor($p[$co+2]-shl 8)-bor$p[$co+3];&$read $ch;$ptr+=2}
    $ch=($p[$o+8]-shl 24)-bor($p[$o+9]-shl 16)-bor($p[$o+10]-shl 8)-bor$p[$o+11];&$read $ch
   }elseif($typ -eq 13){
    for($k=0;$k -lt $cnt;$k++){$co=($p[$ptr]-shl 8)-bor$p[$ptr+1];$q=$co;$len=Get-SqliteVarInt $p ([ref]$q);$null=Get-SqliteVarInt $p ([ref]$q);$n=[int]$len;if($q + $n -gt $ps){$n=$ps-$q};$out+=,(Get-SqliteRecord $(if($n){[byte[]]$p[$q..($q+$n-1)]}else{[byte[]]@()}));$ptr+=2}
   }
  };&$read $root;$out
 }finally{$fs.Dispose()}
}
$Tasks=@(
@{N='01 OS';C={ $o=Get-CimInstance Win32_OperatingSystem;$r=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion';$c=Get-CimInstance Win32_ComputerSystem;@((H '01. IDENTIFICAÇÃO / SISTEMA OPERACIONAL'),"Computador: $env:COMPUTERNAME","Usuário: $env:USERNAME","Fabricante/Modelo: $($c.Manufacturer) / $($c.Model)","Windows: $($r.ProductName)","Display Version: $($r.DisplayVersion)","Build: $($r.CurrentBuild).$($r.UBR)","Versão: $($o.Version)","Arquitetura: $($o.OSArchitecture)","Instalação: $($o.InstallDate)","Último boot: $($o.LastBootUpTime)","Uptime dias: $([math]::Round(((Get-Date)-$o.LastBootUpTime).TotalDays,1))")-join "`r`n" }}
@{N='02 Hardware';C={ $c=Get-CimInstance Win32_Processor|Select-Object -First 1;$r=Get-CimInstance Win32_PhysicalMemory;$d=Get-CimInstance Win32_DiskDrive;$v=Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3';$g=Get-CimInstance Win32_VideoController;$x=H '02. HARDWARE';$x+="`r`nCPU: $($c.Name) | $($c.NumberOfCores)C/$($c.NumberOfLogicalProcessors)T";$x+="`r`nRAM:";$r|ForEach-Object {$x+="`r`n  - $([math]::Round($_.Capacity/1GB,1)) GB | $($_.Speed) MHz | $($_.Manufacturer) | $($_.PartNumber)"};$x+="`r`nDISCOS:";$d|ForEach-Object {$x+="`r`n  - $($_.Model) | $([math]::Round($_.Size/1GB,0)) GB | $($_.InterfaceType) | $($_.SerialNumber)"};$x+="`r`nVOLUMES:";$v|ForEach-Object {$x+="`r`n  - $($_.DeviceID) | $($_.FileSystem) | $([math]::Round($_.Size/1GB,1)) GB | Livre $([math]::Round($_.FreeSpace/$_.Size*100,1))%"};$x+="`r`nGPU:";$g|ForEach-Object {$x+="`r`n  - $($_.Name) | Driver $($_.DriverVersion)"};$x }}
@{N='03 Rede';C={ $x=H '03. REDE';Get-NetAdapter|ForEach-Object {$x+="`r`nADP: $($_.Name) | $($_.Status) | $($_.MacAddress) | $($_.LinkSpeed)"};Get-NetIPConfiguration|ForEach-Object {$x+="`r`nIP: $($_.InterfaceAlias) | $((($_.IPv4Address).IPAddress)-join ',') | GW: $((($_.IPv4DefaultGateway).NextHop)-join ',') | DNS: $((($_.DNSServer).ServerAddresses)-join ',')"};$x }}
@{N='04 Segurança';C={ $x=H '04. SEGURANÇA';Get-NetFirewallProfile|ForEach-Object {$x+="`r`nFirewall $($_.Name): $($_.Enabled)"};$t=Get-Tpm;$x+="`r`nTPM: Presente=$($t.TpmPresent) Pronto=$($t.TpmReady) Versão=$($t.ManufacturerVersion)";$b=Get-BitLockerVolume;$b|ForEach-Object {$x+="`r`nBitLocker: $($_.MountPoint) | $($_.VolumeStatus) | $($_.ProtectionStatus)"};$x+="`r`nSecure Boot: $(if(Confirm-SecureBootUEFI){'Ativado'}else{'Não confirmado'})";$x }}
@{N='05 Updates';C={ $x=H '05. WINDOWS UPDATE / PATCHES';Get-HotFix|Sort-Object InstalledOn -Descending|Select-Object -First 30|ForEach-Object {$x+="`r`n  - $($_.HotFixID) | $($_.InstalledOn) | $($_.Description)"};$x }}
@{N='06 Software';C={ $x=H '06. SOFTWARE / DRIVERS';Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'|Where-Object { $_.DisplayName }|Sort-Object DisplayName|Select-Object -First 150|ForEach-Object {$x+="`r`n  - $($_.DisplayName) v$($_.DisplayVersion) | $($_.Publisher)"};$x }}
@{N='07 Startup';C={ $x=H '07. SERVIÇOS / STARTUP';Get-Service|Where-Object { $_.Status -eq 'Running' }|Select-Object -First 150|ForEach-Object {$x+="`r`n  - $($_.Name) | $($_.DisplayName)"};foreach($p in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run','HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'){$z=Get-ItemProperty $p;if($z){$z.PSObject.Properties|Where-Object { $_.Name -notmatch '^PS' }|ForEach-Object {$x+="`r`nRUN: $($_.Name) = $($_.Value)"}}};$x }}
@{N='08 Usuários';C={ $x=H '08. USUÁRIOS / COMPARTILHAMENTOS';Get-LocalUser|ForEach-Object {$x+="`r`n  - $($_.Name) | Ativo=$($_.Enabled) | ÚltimoLogin=$($_.LastLogon)"};$x+='`r`nADMINISTRADORES:';Get-LocalGroupMember Administrators|ForEach-Object {$x+="`r`n  - $($_.Name)"};Get-SmbShare|ForEach-Object {$x+="`r`nSMB: $($_.Name) | $($_.Path)"};$x }}
@{N='09 Diagnóstico';C={ $x=H '09. DIAGNÓSTICO';Get-Process|Sort-Object WorkingSet64 -Descending|Select-Object -First 40|ForEach-Object {$x+="`r`n  - $($_.ProcessName) | PID=$($_.Id) | $([math]::Round($_.WorkingSet64/1MB,1)) MB"};$x+='`r`nPORTAS:';netstat -ano 2>$null|Select-String LISTENING|Select-Object -First 80|ForEach-Object {$x+="`r`n  - $($_.Line.Trim())"};$x }}
@{N='10 Recentes';C={ $x=H '10. ARQUIVOS RECENTES';$p=Join-Path $env:APPDATA 'Microsoft\Windows\Recent';Get-ChildItem $p -File|Sort-Object LastWriteTime -Descending|Select-Object -First $RecentFiles|ForEach-Object {$x+="`r`n  - $($_.LastWriteTime) | $($_.Name) | $($_.FullName)"};$x }}
@{N='11 Ambiente';C={ $x=H '11. AMBIENTE';'PATH','APPDATA','PROGRAMDATA','SYSTEMDRIVE','SYSTEMROOT','TEMP','TMP','USERNAME','COMPUTERNAME'|ForEach-Object {$v=[Environment]::GetEnvironmentVariable($_);$x+="`r`n$_ = $v"};$x }}
@{N='12 Downloads';C={ $x=H '12. DOWNLOADS';$f=Get-ChildItem "$env:USERPROFILE\Downloads" -File -Force -ErrorAction SilentlyContinue|Sort-Object LastWriteTime -Descending|Select-Object -First $DownloadItems;$x+="`r`nTotal coletado: $($f.Count)";$f|ForEach-Object {$x+="`r`n  - $($_.LastWriteTime) | $([math]::Round($_.Length/1KB,1)) KB | $($_.Extension) | $($_.FullName)"};$x+='`r`nRESUMO POR EXTENSÃO:';$f|Group-Object Extension|Sort-Object Count -Descending|ForEach-Object {$x+="`r`n  - $($_.Name): $($_.Count)"};$x }}
@{N='13 Navegadores';C={
 $x=H '13. NAVEGADORES / HISTÓRICO / BOOKMARKS / EXTENSÕES'
 $bs=@(@('Chrome',"$env:LOCALAPPDATA\Google\Chrome\User Data"),@('Edge',"$env:LOCALAPPDATA\Microsoft\Edge\User Data"),@('Brave',"$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data"),@('Vivaldi',"$env:LOCALAPPDATA\Vivaldi\User Data"),@('Opera',"$env:APPDATA\Opera Software\Opera Stable"))
 foreach($b in $bs){if(!(Test-Path $b[1])){continue};$x+="`r`n[$($b[0])]"
  $ps=Get-ChildItem $b[1] -Directory|Where-Object {$_.Name -eq 'Default' -or $_.Name -like 'Profile *'}
  foreach($p in $ps){$x+="`r`nPerfil: $($p.Name)";$db=Join-Path $p.FullName 'History'
   if(Test-Path $db){try{$u=Get-SqliteRows $db 'urls';$x+="Histórico: $($u.Count) URLs";$x+='URLs recentes:'
    $u | ForEach-Object { $title=if($_.1){$_.1}else{'(sem título)'}; $url=if($_.0){$_.0}else{'(sem URL)'}; $vc=if($_.2 -ne $null){$_.2}else{0}; $x+="  - $title | Visitas: $vc | $url" } | Select-Object -First $BrowserItems
   }catch{$x+="Histórico: falha SQLite nativo - $($_.Exception.Message)"}}else{$x+='Histórico: banco não encontrado'}
   $bm=Join-Path $p.FullName 'Bookmarks'
   if(Test-Path $bm){try{$j=Get-Content $bm -Raw|ConvertFrom-Json;$fav=@();$walk=$null;$walk={param($n,$path);if($n.children){$n.children|ForEach-Object {&$walk $_ "$path/$($n.name)"}}elseif($n.url){$fav+=[pscustomobject]@{Nome=$n.name;URL=$n.url;Pasta=$path}}};$j.roots.PSObject.Properties|ForEach-Object {&$walk $_.Value $_.Name};$x+="Bookmarks: $($fav.Count)";$fav | Select-Object -First $BrowserItems | ForEach-Object {$x+="  - $($_.Nome) | $($_.URL) | Pasta: $($_.Pasta)"}}catch{$x+='Bookmarks: JSON inválido/não legível'}}else{$x+='Bookmarks: não encontrado'}
   $ep=Join-Path $p.FullName 'Extensions';if(Test-Path $ep){$x+='Extensões:';Get-ChildItem $ep -Directory|Select-Object -First $BrowserItems|ForEach-Object {$x+="  - $($_.Name)"}}
  }
 }
 $ff="$env:APPDATA\Mozilla\Firefox\Profiles";if(Test-Path $ff){$x+="`r`n[Firefox]";Get-ChildItem $ff -Directory|ForEach-Object {$x+="`r`nPerfil: $($_.Name)";$db=Join-Path $_.FullName 'places.sqlite'
  if(Test-Path $db){try{$p=Get-SqliteRows $db 'moz_places';$v=Get-SqliteRows $db 'moz_historyvisits';$x+="Histórico: $($v.Count) visitas | $($p.Count) URLs";$x+='URLs:';$p | Select-Object -First $BrowserItems | ForEach-Object {$x+="  - $($_.2) | $($_.1)"}}catch{$x+="Histórico: falha SQLite nativo - $($_.Exception.Message)"}}else{$x+='Histórico: places.sqlite não encontrado'}
  $e=Join-Path $_.FullName 'extensions.json';$x+="Extensões: $(if(Test-Path $e){'arquivo presente'}else{'não encontrado'})"
 }}
 $x+='';$x+='SQLite: parser nativo PowerShell/.NET; sem SQLite.dll, sqlite3.exe, módulos ou instalações externas.';$x+='Segurança: não coleta senhas, cookies, tokens ou sessões.';$x -join "`r`n" }}
@{N='14 OneDrive';C={
 $x=H '14. ONEDRIVE LOCAL'
 $a=@(Get-ItemProperty 'HKCU:\Software\Microsoft\OneDrive\Accounts\*' -ErrorAction SilentlyContinue | ForEach-Object { $_.UserFolder } | Where-Object { $_ -and (Test-Path $_) })
 if(!$a){$a=@(Join-Path $env:USERPROFILE 'OneDrive' | Where-Object { Test-Path $_ })}
 if(!$a){$x+='`r`nNenhuma pasta OneDrive detectada.'}
 else{foreach($d in $a|Select-Object -Unique){$x+="`r`nPasta: $d";Get-ChildItem $d -File -Recurse -Force -ErrorAction SilentlyContinue|Sort-Object LastWriteTime -Descending|Select-Object -First $RecentFiles|ForEach-Object {$x+="`r`n  - $($_.LastWriteTime) | $([math]::Round($_.Length/1KB,1)) KB | $($_.FullName)"}}}
 $x
 }}
@{N='15 OneDrive Shares';C={
 $x=H '15. ONEDRIVE / LINKS DE COMPARTILHAMENTO'
 if(!$CollectOneDriveShares){
  $x+='`r`nGraph desativado. Use -CollectOneDriveShares.'
 }else{
  $x+='`r`nA consulta de links/permissões requer Microsoft Graph autenticado.'
  $x+=' Esta versão não reutiliza tokens locais nem solicita credenciais.'
  if($IncludeShareLinks){
   $x+='`r`n-IncludeShareLinks solicitado; aguardando integração Graph autenticada.'
  }
 }
 $x
 }}
)
$native=@(
 'function Get-SqliteVarInt '+(Get-Command Get-SqliteVarInt).ScriptBlock.ToString(),
 'function Get-SqliteRecord '+(Get-Command Get-SqliteRecord).ScriptBlock.ToString(),
 'function Get-SqlitePage '+(Get-Command Get-SqlitePage).ScriptBlock.ToString(),
 'function Get-SqliteRows '+(Get-Command Get-SqliteRows).ScriptBlock.ToString()
)-join "`r`n";
Write-CollectorLog "FASE Runspace | Criando pool com $($Tasks.Count) tarefas";$iss=[initialsessionstate]::CreateDefault();$pool=[runspacefactory]::CreateRunspacePool(1,$MaxRunspaces,$iss,$Host);$pool.Open();Write-CollectorLog 'FASE Runspace | Pool aberto';$q=@()
Write-CollectorLog 'FASE Runspace | Iniciando tarefas';foreach($t in $Tasks){Write-CollectorLog "QUEUE | $($t.N)";$ps=[powershell]::Create();$ps.RunspacePool=$pool;[void]$ps.AddScript({param($n,$c,$init,$log);& ([scriptblock]::Create($init));$w=[Diagnostics.Stopwatch]::StartNew();try{Add-Content -LiteralPath $log -Value ("{0} [TASK-START] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$n) -ErrorAction SilentlyContinue;$x=&$c;Add-Content -LiteralPath $log -Value ("{0} [TASK-END] {1} OK {2}s" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$n,[math]::Round($w.Elapsed.TotalSeconds,2)) -ErrorAction SilentlyContinue;[pscustomobject]@{N=$n;S='OK';T=[math]::Round($w.Elapsed.TotalSeconds,2);X=$x}}catch{Add-Content -LiteralPath $log -Value ("{0} [TASK-ERROR] {1} {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$n,$_.Exception.ToString()) -ErrorAction SilentlyContinue;[pscustomobject]@{N=$n;S='ERRO';T=[math]::Round($w.Elapsed.TotalSeconds,2);X="[$n] ERRO: $($_.Exception.Message)"}}}).AddArgument($t.N).AddArgument($t.C).AddArgument($native).AddArgument($LogFile);$q+=[pscustomobject]@{N=$t.N;P=$ps;A=$ps.BeginInvoke()}}
$r=@();Write-CollectorLog "FASE Runspace | Aguardando conclusao (timeout=$TimeoutSeconds s)";$end=(Get-Date).AddSeconds($TimeoutSeconds);while($q.Count){foreach($j in @($q)){if($j.A.IsCompleted){try{$r+=$j.P.EndInvoke($j.A)}catch{$r+=[pscustomobject]@{N=$j.N;S='ERRO';T=0;X="[$($j.N)] falha ao finalizar"}};$j.P.Dispose();$nq=@();foreach($qq in $q){if($qq -ne $j){$nq+=$qq}};$q=$nq}};if($q.Count -and (Get-Date)-gt $end){foreach($j in @($q)){$j.P.Stop();$r+=[pscustomobject]@{N=$j.N;S='TIMEOUT';T=$TimeoutSeconds;X="[$($j.N)] TIMEOUT $TimeoutSeconds s"};$j.P.Dispose()};$q=@()}else{Start-Sleep -Milliseconds 40}}
$pool.Close();$pool.Dispose();Write-CollectorLog "FASE Runspace | Concluido | Resultados=$($r.Count)";$parent=Split-Path $OutputFile -Parent;if($parent -and !(Test-Path $parent)){New-Item $parent -ItemType Directory -Force|Out-Null};$sb=[Text.StringBuilder]::new();$sep='='*78;[void]$sb.AppendLine($sep);[void]$sb.AppendLine('INVENTÁRIO TÉCNICO COMPLETO - V5 CONCORRENTE + SQLITE NATIVO');[void]$sb.AppendLine("Coletado: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Workers: $MaxRunspaces | Timeout: $TimeoutSeconds s");[void]$sb.AppendLine($sep)
foreach($z in $r){[void]$sb.AppendLine($z.X);[void]$sb.AppendLine('')};[void]$sb.AppendLine($sep);[void]$sb.AppendLine('## DESEMPENHO');$r | ForEach-Object { [void]$sb.AppendLine(('{0,-25} | {1,-8} | {2,7}s'-f$_.N,$_.S,$_.T))};[void]$sb.AppendLine("Total: $([math]::Round(((Get-Date)-$Started).TotalSeconds,2)) s");[void]$sb.AppendLine($sep);$sb.ToString()|Out-File $OutputFile -Encoding UTF8 -Force;Write-CollectorLog "RELATORIO | $OutputFile";Write-CollectorLog 'FIM | Coleta concluida';Write-Host "Relatório: $OutputFile";Write-Host "Tempo total: $([math]::Round(((Get-Date)-$Started).TotalSeconds,2)) s";return (Resolve-Path $OutputFile).ProviderPath
}

# -----------------------------------------------------------------------------
# EXECUÇÃO AUTOMÁTICA
# -----------------------------------------------------------------------------
Start-SystemInfoCollector `
 -OutputFile $OutputFile `
 -MaxRunspaces $MaxRunspaces `
 -TimeoutSeconds $TimeoutSeconds `
 -RecentFiles $RecentFiles `
 -BrowserItems $BrowserItems `
 -DownloadItems $DownloadItems `
 -CollectOneDriveShares:$CollectOneDriveShares `
 -IncludeShareLinks:$IncludeShareLinks
