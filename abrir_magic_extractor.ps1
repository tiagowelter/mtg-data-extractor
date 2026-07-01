$scriptPath = Join-Path $PSScriptRoot "magic_extractor.py"

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
