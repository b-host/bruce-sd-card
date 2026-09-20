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

function Invoke-SystemInventory {
    [CmdletBinding()]
    param(
        [string]$OutputFile=(Join-Path $env:TEMP 'system_info_v3.txt'),
        [string]$ExportDir=(Join-Path $env:TEMP 'p'),
        [ValidateRange(1,32)]
        [int]$MaxRunspaces=[Math]::Min([Environment]::ProcessorCount,8),
        [ValidateRange(10,600)]
        [int]$TimeoutSeconds=120,
        [ValidateRange(1,500)]
        [int]$RecentFiles=100
    )

    $ErrorActionPreference='SilentlyContinue'
    $Started=Get-Date

    # ============================================================
    # TODO: TODO O SEU CÓDIGO ATUAL FICA AQUI
    # ============================================================

    $Tasks=@(
        # ...
        # todo o bloco atual de $Tasks
        # ...
    )

    $iss=[initialsessionstate]::CreateDefault()
    $pool=[runspacefactory]::CreateRunspacePool(
        1,
        $MaxRunspaces,
        $iss,
        $Host
    )
    $pool.Open()

    $pending=@()

    foreach($t in $Tasks){
        $ps=[powershell]::Create()
        $ps.RunspacePool=$pool

        [void]$ps.AddScript({
            param($n,$o,$c)

            $sw=[Diagnostics.Stopwatch]::StartNew()

            try {
                $text=&$c

                [pscustomobject]@{
                    Name=$n
                    Order=$o
                    Status='OK'
                    Seconds=[math]::Round($sw.Elapsed.TotalSeconds,2)
                    Text=$text
                    Error=''
                }
            }
            catch {
                [pscustomobject]@{
                    Name=$n
                    Order=$o
                    Status='ERRO'
                    Seconds=[math]::Round($sw.Elapsed.TotalSeconds,2)
                    Text="[$n] falhou: $($_.Exception.Message)"
                    Error=$_.Exception.ToString()
                }
            }
        }).AddArgument($t.N).AddArgument($t.O).AddArgument($t.C)

        $pending += [pscustomobject]@{
            T=$t
            P=$ps
            A=$ps.BeginInvoke()
        }
    }

    $results=@()
    $deadline=(Get-Date).AddSeconds($TimeoutSeconds)

    while($pending.Count){

        foreach($j in @($pending)){
            if($j.A.IsCompleted){

                try {
                    $results += $j.P.EndInvoke($j.A)
                }
                catch {
                    $results += [pscustomobject]@{
                        Name=$j.T.N
                        Order=$j.T.O
                        Status='ERRO'
                        Seconds=0
                        Text="[$($j.T.N)] falhou ao finalizar: $($_.Exception.Message)"
                        Error=$_.Exception.ToString()
                    }
                }

                $j.P.Dispose()
                $pending=@($pending|?{$_ -ne $j})
            }
        }

        if($pending.Count -and (Get-Date) -gt $deadline){

            foreach($j in @($pending)){
                $j.P.Stop()

                $results += [pscustomobject]@{
                    Name=$j.T.N
                    Order=$j.T.O
                    Status='TIMEOUT'
                    Seconds=$TimeoutSeconds
                    Text="[$($j.T.N)] timeout após $TimeoutSeconds segundos."
                    Error='Timeout'
                }

                $j.P.Dispose()
            }

            $pending=@()
        }
        else {
            Start-Sleep -Milliseconds 50
        }
    }

    $pool.Close()
    $pool.Dispose()

    $results=$results|sort Order
    $sb=[Text.StringBuilder]::new()
    $sep='='*78

    [void]$sb.AppendLine($sep)
    [void]$sb.AppendLine(
        'INVENTÁRIO TÉCNICO COMPLETO DO SISTEMA - V3 CONCORRENTE'
    )
    [void]$sb.AppendLine(
        "Coletado em: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    )
    [void]$sb.AppendLine(
        "Workers: $MaxRunspaces | Timeout: $TimeoutSeconds s"
    )
    [void]$sb.AppendLine($sep)

    foreach($r in $results){
        [void]$sb.AppendLine($r.Text)
        [void]$sb.AppendLine('')
    }

    [void]$sb.AppendLine($sep)
    [void]$sb.AppendLine('## TEMPO / STATUS DAS COLETAS')
    [void]$sb.AppendLine(('-'*78))

    $results|%{
        [void]$sb.AppendLine(
            ('{0,-22} | {1,-8} | {2,7}s' -f `
                $_.Name,
                $_.Status,
                $_.Seconds)
        )
    }

    [void]$sb.AppendLine(
        "Tempo total: $([math]::Round(((Get-Date)-$Started).TotalSeconds,2)) s"
    )
    [void]$sb.AppendLine($sep)

    if(!(Test-Path $ExportDir -PathType Container)){
        New-Item $ExportDir -ItemType Directory -Force|Out-Null
    }

    $sb.ToString()|Out-File $OutputFile -Encoding UTF8 -Force

    Write-Host "Relatório salvo em: $OutputFile"
    Write-Host "Tempo total: $([math]::Round(((Get-Date)-$Started).TotalSeconds,2)) s"
    Write-Host "Workers: $MaxRunspaces | Timeout: $TimeoutSeconds s"

    # Retorna o resultado para quem chamou a função
    [pscustomobject]@{
        OutputFile=$OutputFile
        ExportDir=$ExportDir
        TotalSeconds=[math]::Round(((Get-Date)-$Started).TotalSeconds,2)
        Workers=$MaxRunspaces
        TimeoutSeconds=$TimeoutSeconds
        Sections=$results.Count
        Results=$results
    }
}

# ============================================================
# EXECUÇÃO AUTOMÁTICA
# ============================================================

Invoke-SystemInventory