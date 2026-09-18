Function Get-ChromeDump {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$OutFile = 'C:\Get-ChromeDump.txt'
    )

    # Load SQLite assembly
    $assemblyPath = "$env:LOCALAPPDATA\System.Data.SQLite.dll"
    if (Test-Path $assemblyPath) {
        Add-Type -Path $assemblyPath
    } else {
        Write-Warning "[!] SQLite assembly not found at: $assemblyPath"
        return
    }

    # Close Chrome if running
    if (Get-Process -Name chrome -ErrorAction SilentlyContinue) {
        Write-Warning "[!] Stopping Chrome..."
        Stop-Process -Name chrome -Force
    }

    # Locate Chrome user data directory
    $chromepath = if ([environment]::OSVersion.Version.Major -ge 6) {
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default"
    } else {
        "${env:HOMEDRIVE}${env:HOMEPATH}\Local Settings\Application Data\Google\Chrome\User Data\Default"
    }

    if (-not (Test-Path $chromepath)) {
        Throw "Chrome user data directory does not exist"
    }

    # --- Extract Browsing History ---
    Write-Verbose "Parsing history..."
    $history = @()

    $conn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$chromepath\History; Version=3;")
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT url, title, visit_count FROM urls"
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $history += [PSCustomObject]@{
            URL         = $reader.GetString($reader.GetOrdinal('url'))
            Title       = $reader.IsDBNull($reader.GetOrdinal('title')) ? '(null)' : $reader.GetString($reader.GetOrdinal('title'))
            VisitCount  = $reader.GetInt32($reader.GetOrdinal('visit_count'))
        }
    }
    $reader.Close()
    $conn.Close()

    # --- Write output ---
    $null = New-Item -Path (Split-Path $OutFile -Parent) -ItemType Directory -Force

    "[*]HISTORY
" | Out-File $OutFile

    if ($history) {
        $history | Format-Table -AutoSize | Out-File $OutFile -Append
    } else {
        "  (none found)" | Out-File $OutFile -Append
    }

    Write-Host "[+] Dump saved to: $OutFile"
    Write-Warning "[!] Remember to remove SQLite assembly from: $assemblyPath"
}

Get-ChromeDump
