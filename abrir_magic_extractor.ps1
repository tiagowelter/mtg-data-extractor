$scriptPath = Join-Path $PSScriptRoot "magic_extractor.py"

function Set-ValidSslBundle {
    param([string]$VariableName)

    $value = [Environment]::GetEnvironmentVariable($VariableName, "Process")
    if (-not $value) {
        return
    }
    if (-not (Test-Path -LiteralPath $value)) {
        $certifiPath = python -c "import certifi; print(certifi.where())" 2>$null
        if ($certifiPath -and (Test-Path -LiteralPath $certifiPath)) {
            Set-Item -Path "Env:$VariableName" -Value $certifiPath
        } else {
            Remove-Item "Env:$VariableName" -ErrorAction SilentlyContinue
        }
    }
}

Set-ValidSslBundle "SSL_CERT_FILE"
Set-ValidSslBundle "REQUESTS_CA_BUNDLE"
Set-ValidSslBundle "CURL_CA_BUNDLE"

Get-CimInstance Win32_Process |
    Where-Object {
        $_.Name -in @("python.exe", "pythonw.exe") -and
        $_.CommandLine -like "*magic_extractor.py*" -and
        $_.CommandLine -notlike "*Get-CimInstance*"
    } |
    ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force
    }

Start-Process -FilePath "python" -ArgumentList "`"$scriptPath`"" -WorkingDirectory $PSScriptRoot
