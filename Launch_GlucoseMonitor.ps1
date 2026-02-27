# Launcher - uruchamia glowny skrypt w trybie STA (wymagane dla WPF)
$installDir = "C:\Glucose"
$logFile    = Join-Path $installDir "launcher.log"

try {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Launcher START" | Out-File $logFile -Append -Encoding UTF8

    $scriptPath = Join-Path $installDir "GlucoseMonitor.ps1"
    if (-not (Test-Path $scriptPath)) {
        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] BLAD: Nie znaleziono $scriptPath" | Out-File $logFile -Append -Encoding UTF8
        exit 1
    }

    # Sprawdz czy jestesmy w trybie STA (WPF wymaga STA)
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Restart w trybie STA..." | Out-File $logFile -Append -Encoding UTF8
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$scriptPath`"" -WorkingDirectory $installDir
        exit 0
    }

    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Uruchamiam (STA=$([System.Threading.Thread]::CurrentThread.GetApartmentState()))" | Out-File $logFile -Append -Encoding UTF8
    Set-Location $installDir
    & $scriptPath

} catch {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WYJATEK: $($_.Exception.Message)" | Out-File $logFile -Append -Encoding UTF8
}
