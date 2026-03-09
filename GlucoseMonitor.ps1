# ============================================================================
# GLUCOSE MONITOR - LibreLinkUp v5 (mmol/L, tray, movable)
# ============================================================================
param([switch]$AutoStart)   # uruchom ukryty w tray (Task Scheduler)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Ukryj konsole PowerShell
try {
    [Native.Win32] | Out-Null
} catch {
    Add-Type -Name Win32 -Namespace Native -MemberDefinition @"
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
"@
}
$consoleHwnd = [Native.Win32]::GetConsoleWindow()
[Native.Win32]::ShowWindow($consoleHwnd, 0) | Out-Null

$script:Config = @{
    Email     = ""
    Password  = ""
    ApiUrl    = "https://api-eu.libreview.io"
    Interval  = 60
    Version   = "4.16.0"
    Product   = "llu.ios"
    AlertLow  = 3.9    # mmol/L
    AlertHigh = 13.9   # mmol/L
}

# Ustal folder skryptu (robust fallback)
$script:ScriptDir = $PSScriptRoot
if (-not $script:ScriptDir) {
    $def = $MyInvocation.MyCommand.Definition
    if ($def -and $def -notmatch "`n") { try { $script:ScriptDir = Split-Path -Parent $def } catch {} }
}
if (-not $script:ScriptDir) { $script:ScriptDir = $PWD.Path }
if (-not $script:ScriptDir) { $script:ScriptDir = "C:\Glucose" }
# Pelna sciezka do pliku .ps1 (potrzebna do Task Scheduler)
$script:ScriptPath = if ($PSCommandPath) { $PSCommandPath } `
                     else { Join-Path $script:ScriptDir "GlucoseMonitor.ps1" }

$script:ConfigFile = Join-Path $script:ScriptDir "config.ini"

$script:LogFile = Join-Path $script:ScriptDir "glucose_debug.log"
function Write-Log { param([string]$M); $l="[$(Get-Date -Format 'HH:mm:ss')] $M"; Write-Host $l; try{Add-Content -Path $script:LogFile -Value $l -ErrorAction Stop}catch{} }

$script:AuthToken=$null; $script:AccountId=$null; $script:AccountIdHash=$null; $script:PatientId=$null; $script:Timer=$null
$script:UseMgDl = $false
$script:CachedMgDl = $null; $script:CachedTrend = 0; $script:CachedGraphData = $null
$script:LangEn    = $false   # false = PL (domyslny), true = EN
$script:SmoothMode = $false  # false = RAW, true = Savitzky-Golay smooth
$script:LastHistorySave = $null
$script:BackoffUntil = $null  # 429 backoff - nie wywoluj API do tej daty
$script:LastAlertTime  = $null # throttle alertow - max 1 alert co 15 minut
$script:LastReadingTs  = $null # timestamp odczytu z czujnika (DateTime) - do wyswietlania godziny
$script:LastFetchTime  = $null # kiedy aplikacja ostatnio pobrala dane z API (DateTime) - do licznika wieku
$script:HistOffset     = 0     # przesuniecie widoku historii wstecz (w dniach)
$script:HistValHbA1c  = $null
$script:HistRangeLabel = $null
$script:HistValSD     = $null
$script:HistValCV     = $null
$script:HistBtnSmooth = $null  # przycisk ~ w oknie historii
$script:HistWinLeft   = $null # zapamietana pozycja okna historii (X)
$script:HistWinTop    = $null # zapamietana pozycja okna historii (Y)
$script:AvgBarsWinLeft = $null # zapamietana pozycja okna BARS (X)
$script:AvgBarsWinTop  = $null # zapamietana pozycja okna BARS (Y)
$script:AgpWinLeft     = $null # zapamietana pozycja okna AGP (X)
$script:AgpWinTop      = $null # zapamietana pozycja okna AGP (Y)
$script:LogWin         = $null # okno dziennika (Logbook)
$script:LogWinLeft     = $null # zapamietana pozycja okna Logbook (X)
$script:LogWinTop      = $null # zapamietana pozycja okna Logbook (Y)

$script:T = @{
    Fetching   = @("Pobieranie...",                       "Fetching...")
    Connected  = @("Polaczono | Odczyt aktualny",         "Connected | Reading current")
    TooMany    = @("Zbyt wiele zadan - sprobuj pozniej",  "Too many requests - try again later")
    NoData     = @("Brak danych",                         "No data")
    TrayGlc    = @("Glukoza: ",                           "Glucose: ")
    TrayTooM   = @("Glucose Monitor - zbyt wiele zadan",  "Glucose Monitor - too many requests")
    TrayNoDat  = @("Glucose Monitor - brak danych",       "Glucose Monitor - no data")
    ShowWin      = @("Pokaz okno",                          "Show window")
    SwitchAcc    = @("Zmien konto",                         "Switch account")
    CloseApp     = @("Zamknij",                             "Exit")
    AutoStartOn  = @("Uruchom z Windows [wlaczone]",        "Run with Windows [enabled]")
    AutoStartOff = @("Uruchom z Windows [wylaczone]",       "Run with Windows [disabled]")
    SettingsMenu = @("Ustawienia...",                        "Settings...")
    BackupMenu   = @("Kopia zapasowa historii",             "Backup history")
    RestoreMenu  = @("Przywroc historie...",                "Restore history...")
    TrendLbl   = @("Trend",                               "Trend")
    AvgLbl     = @("Sred.",                               "Avg")
    UnitTip    = @("Przelacz jednostki",                  "Switch units")
    RefreshBtn = @("Odswiez",                             "Refresh")
    NextUpdate = @("Odczyt za:",                          "Next in:")
    HistBtn    = @("HIST",                                "HIST")
    HistTitle  = @("Historia glukozy",                    "Glucose history")
    HistAvg    = @("Srednia",                             "Average")
    HistTIR    = @("W normie %",                          "In range %")
    HistNoData = @("Brak danych historycznych",           "No history data yet")
    HistClose  = @("Zamknij",                             "Close")
    HistDays   = @("dni",                                 "days")
    HistLineTip= @("Odczyty glukozy (linia czasu)",       "Glucose readings (timeline)")
    LogTitle   = @("Dziennik",                             "Logbook")
    LogLoading = @("Pobieranie...",                        "Loading...")
    LogNoData  = @("Brak wpisow",                          "No entries")
    LogScan    = @("Skan",                                 "Scan")
    LogAuto    = @("Auto",                                 "Auto")
    LogBtn     = @("LOG",                                  "LOG")
}
function t([string]$key) { $script:T[$key][[int]$script:LangEn] }

$script:HistoryFile  = Join-Path $script:ScriptDir "history.jsonl"
$script:LogbookFile  = Join-Path $script:ScriptDir "logbook.jsonl"

$script:HistKnownTs    = $null  # HashSet znanych timestampow (yyyy-MM-ddTHH:mm) - inicjowany przy pierwszym uzyciu
$script:LogbookKnownTs = $null  # HashSet znanych timestampow logbook (yyyy-MM-ddTHH:mm)

function Save-HistoryEntry([double]$mgdl, [int]$trend) {
    try {
        # Inicjuj HashSet przy pierwszym wywolaniu - wczytaj znane timestampy z pliku
        if ($null -eq $script:HistKnownTs) {
            $script:HistKnownTs = [System.Collections.Generic.HashSet[string]]::new()
            if (Test-Path $script:HistoryFile) {
                try {
                    Get-Content $script:HistoryFile -Encoding UTF8 | ForEach-Object {
                        try { $script:HistKnownTs.Add(($_ | ConvertFrom-Json).ts.Substring(0,16)) | Out-Null } catch {}
                    }
                } catch {}
            }
        }
        
        # Uzyj czasu odczytu z czujnika (jesli dostepny), w przeciwnym razie biezacy czas
        $timestamp = if ($script:LastReadingTs) { $script:LastReadingTs } else { Get-Date }
        
        # Sprawdz, czy wpis z takim kluczem juz istnieje
        $key = $timestamp.ToString("yyyy-MM-ddTHH:mm")
        if ($script:HistKnownTs.Contains($key)) { return }
        
        # Zapisz wpis
        $script:HistKnownTs.Add($key) | Out-Null
        $entry = '{"ts":"' + $timestamp.ToString("yyyy-MM-ddTHH:mm:ss") + '","mgdl":' + [Math]::Round($mgdl,1) + ',"trend":' + $trend + '}'
        Add-Content -Path $script:HistoryFile -Value $entry -Encoding UTF8 -ErrorAction Stop
    } catch {}
}

function Save-GraphDataHistory([array]$graphData) {
    if (-not $graphData -or $graphData.Count -eq 0) { return }
    # Inicjuj HashSet przy pierwszym wywolaniu - wczytaj znane timestampy z pliku
    if ($null -eq $script:HistKnownTs) {
        $script:HistKnownTs = [System.Collections.Generic.HashSet[string]]::new()
        if (Test-Path $script:HistoryFile) {
            try {
                Get-Content $script:HistoryFile -Encoding UTF8 | ForEach-Object {
                    try { $script:HistKnownTs.Add(($_ | ConvertFrom-Json).ts.Substring(0,16)) | Out-Null } catch {}
                }
            } catch {}
        }
    }
    $fmts = @("M/d/yyyy h:mm:ss tt","d/M/yyyy h:mm:ss tt","M/d/yyyy H:mm:ss","d/M/yyyy H:mm:ss","yyyy-MM-ddTHH:mm:ss")
    $toAdd = [System.Collections.Generic.List[string]]::new()
    foreach ($pt in $graphData) {
        if (-not $pt.Timestamp) { continue }
        $mg = try { [double]$pt.ValueInMgPerDl } catch { 0 }
        if ($mg -le 20) { continue }
        $parsed = [DateTime]::MinValue
        foreach ($f in $fmts) {
            if ([DateTime]::TryParseExact([string]$pt.Timestamp, $f, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { break }
        }
        if ($parsed -eq [DateTime]::MinValue) { continue }
        $key = $parsed.ToString("yyyy-MM-ddTHH:mm")
        if ($script:HistKnownTs.Contains($key)) { continue }
        $trend = try { [int]$pt.TrendArrow } catch { 0 }
        $toAdd.Add('{"ts":"' + $parsed.ToString("yyyy-MM-ddTHH:mm:ss") + '","mgdl":' + [Math]::Round($mg,1) + ',"trend":' + $trend + '}')
        $script:HistKnownTs.Add($key) | Out-Null
    }
    if ($toAdd.Count -gt 0) {
        try { Add-Content -Path $script:HistoryFile -Value $toAdd -Encoding UTF8 } catch {}
    }
}

# Wypelnia luki w history.jsonl syntetycznymi punktami (interpolacja liniowa + szum CGM)
# Uruchamiana po Save-GraphDataHistory - wypelnia okres przed oknem API (>8h przerwy)
function Fill-HistoryGaps([array]$graphData) {
    if (-not $graphData -or $graphData.Count -eq 0) { return }
    if (-not (Test-Path $script:HistoryFile)) { return }

    # Parsuj timestampy z graphData - znajdz najstarszy punkt
    $fmts = @("M/d/yyyy h:mm:ss tt","d/M/yyyy h:mm:ss tt","M/d/yyyy H:mm:ss","d/M/yyyy H:mm:ss","yyyy-MM-ddTHH:mm:ss")
    $firstGdTs = [DateTime]::MaxValue; $firstGdMg = 0.0
    foreach ($pt in $graphData) {
        if (-not $pt.Timestamp) { continue }
        $mg = try { [double]$pt.ValueInMgPerDl } catch { 0 }
        if ($mg -le 20) { continue }
        $parsed = [DateTime]::MinValue
        foreach ($f in $fmts) {
            if ([DateTime]::TryParseExact([string]$pt.Timestamp, $f,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { break }
        }
        if ($parsed -ne [DateTime]::MinValue -and $parsed -lt $firstGdTs) {
            $firstGdTs = $parsed; $firstGdMg = $mg
        }
    }
    if ($firstGdTs -eq [DateTime]::MaxValue) { return }

    # Znajdz ostatni wpis w JSONL przed oknem graphData
    $lastBeforeTs = [DateTime]::MinValue; $lastBeforeMg = 0.0
    try {
        Get-Content $script:HistoryFile -Encoding UTF8 | ForEach-Object {
            try {
                $obj = $_ | ConvertFrom-Json
                $ts  = [DateTime]::Parse($obj.ts)
                if ($ts -lt $firstGdTs -and $ts -gt $lastBeforeTs) {
                    $lastBeforeTs = $ts; $lastBeforeMg = [double]$obj.mgdl
                }
            } catch {}
        }
    } catch {}
    if ($lastBeforeTs -eq [DateTime]::MinValue) { return }

    $gapMin = ($firstGdTs - $lastBeforeTs).TotalMinutes
    if ($gapMin -le 20)   { return }   # normalna przerwa - nic do uzupelnienia
    if ($gapMin -gt 4320) { return }   # >3 dni - za duzo, pomijamy

    # Generuj punkty co 5 min z interpolacja liniowa + szum ±3 mg/dL (realistyczny jak CGM)
    $rnd        = [System.Random]::new()
    $totalSteps = [int]($gapMin / 5) - 1
    $rateMgMin  = ($firstGdMg - $lastBeforeMg) / $gapMin
    $trend = if ($rateMgMin -gt 2) { 5 } elseif ($rateMgMin -gt 1) { 4 } `
             elseif ($rateMgMin -gt -1) { 3 } elseif ($rateMgMin -gt -2) { 2 } else { 1 }

    $toAdd = [System.Collections.Generic.List[string]]::new()
    for ($i = 1; $i -le $totalSteps; $i++) {
        $t    = $lastBeforeTs.AddMinutes($i * 5)
        $frac = $i / ($totalSteps + 1.0)
        $mg   = $lastBeforeMg + ($firstGdMg - $lastBeforeMg) * $frac
        $mg  += ($rnd.NextDouble() - 0.5) * 6.0   # szum ±3 mg/dL
        $mg   = [Math]::Round([Math]::Max(20.0, $mg), 1)
        $key  = $t.ToString("yyyy-MM-ddTHH:mm")
        if ($script:HistKnownTs -and $script:HistKnownTs.Contains($key)) { continue }
        $toAdd.Add('{"ts":"' + $t.ToString("yyyy-MM-ddTHH:mm:ss") + '","mgdl":' + $mg + ',"trend":' + $trend + '}')
        if ($script:HistKnownTs) { $script:HistKnownTs.Add($key) | Out-Null }
    }

    if ($toAdd.Count -gt 0) {
        try { Add-Content -Path $script:HistoryFile -Value $toAdd -Encoding UTF8 } catch {}
        Write-Log "Fill-HistoryGaps: uzupelniono $($toAdd.Count) pkt, przerwa $([int]$gapMin) min ($lastBeforeTs -> $firstGdTs)"
    }
}

function Load-HistoryData([int]$days, [int]$offset = 0) {
    $result = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path $script:HistoryFile)) { return $result }
    $endDate   = (Get-Date).AddDays(-$offset)
    $startDate = $endDate.AddDays(-$days)
    try {
        # ReadLines + reczny parsing = ~10x szybciej niz Get-Content | ConvertFrom-Json
        foreach ($line in [System.IO.File]::ReadLines($script:HistoryFile)) {
            if ($line.Length -lt 20) { continue }
            try {
                $i1 = $line.IndexOf('"ts":"')
                if ($i1 -lt 0) { continue }
                $i1 += 6
                $i2 = $line.IndexOf('"', $i1)
                $tsStr = $line.Substring($i1, $i2 - $i1)
                $ts = [DateTime]::Parse($tsStr)
                if ($ts -lt $startDate -or $ts -gt $endDate) { continue }
                $i3 = $line.IndexOf('"mgdl":')
                if ($i3 -lt 0) { continue }
                $i3 += 7
                $i4 = $line.IndexOfAny([char[]]@(',','}'), $i3)
                [int]$mgdl = $line.Substring($i3, $i4 - $i3)
                [int]$trend = 0
                $i5 = $line.IndexOf('"trend":')
                if ($i5 -ge 0) {
                    $i5 += 8
                    $i6 = $line.IndexOfAny([char[]]@(',','}'), $i5)
                    $trend = [int]$line.Substring($i5, $i6 - $i5)
                }
                $result.Add([PSCustomObject]@{ ts=$tsStr; tsdt=$ts; mgdl=$mgdl; trend=$trend })
            } catch {}
        }
    } catch {}
    # Sortowanie po pre-parsed DateTime (bez ponownego Parse)
    return ($result | Sort-Object tsdt)
}

function MgToMmol([double]$mg) { return [Math]::Round($mg / 18.018, 1) }

# ======================== API ========================

function Get-ApiHeaders([switch]$WithAuth) {
    $h = @{ "accept-encoding"="gzip"; "cache-control"="no-cache"
            "content-type"="application/json"; "product"=$script:Config.Product; "version"=$script:Config.Version }
    if ($WithAuth -and $script:AuthToken) { $h["Authorization"]="Bearer $($script:AuthToken)" }
    if ($WithAuth -and $script:AccountIdHash) { $h["account-id"]=$script:AccountIdHash }
    return $h
}

function Invoke-LibreLogin {
    try {
        $body = @{email=$script:Config.Email;password=$script:Config.Password}|ConvertTo-Json
        Write-Log "Login: $($script:Config.ApiUrl)"
        $r = Invoke-RestMethod -Uri "$($script:Config.ApiUrl)/llu/auth/login" -Method POST -Headers (Get-ApiHeaders) -Body $body -ContentType "application/json"
        if ($r.data -and $r.data.redirect -eq $true -and $r.data.region) {
            $script:Config.ApiUrl = "https://api-$($r.data.region).libreview.io"; Write-Log "Redirect -> $($r.data.region)"; return Invoke-LibreLogin
        }
        if ($r.status -eq 0 -and $r.data -and $r.data.authTicket) {
            $script:AuthToken = $r.data.authTicket.token
            if ($r.data.user -and $r.data.user.id) {
                $script:AccountId = $r.data.user.id
                $sha=[System.Security.Cryptography.SHA256]::Create()
                $script:AccountIdHash = -join ($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($script:AccountId))|ForEach-Object{$_.ToString("x2")})
            }
            Write-Log "Login OK"; return $true
        }
        Write-Log "Login fail: status=$($r.status)"; return $false
    } catch { Write-Log "Login ERR: $($_.Exception.Message)"; return $false }
}

function Get-Connections {
    try {
        $r = Invoke-RestMethod -Uri "$($script:Config.ApiUrl)/llu/connections" -Method GET -Headers (Get-ApiHeaders -WithAuth)
        if ($r.status -eq 0 -and $r.data -and $r.data.Count -gt 0) {
            $script:PatientId=$r.data[0].patientId; Write-Log "Patient: $($r.data[0].firstName) ID=$($script:PatientId)"; return $r.data[0]
        }
        return $null
    } catch {
        Write-Log "Conn ERR: $($_.Exception.Message)"
        $saved=$script:AccountIdHash; $script:AccountIdHash=$null
        try {
            $r2 = Invoke-RestMethod -Uri "$($script:Config.ApiUrl)/llu/connections" -Method GET -Headers (Get-ApiHeaders -WithAuth)
            if ($r2.status -eq 0 -and $r2.data -and $r2.data.Count -gt 0) {
                $script:PatientId=$r2.data[0].patientId; return $r2.data[0]
            }
        } catch { Write-Log "Conn retry ERR: $($_.Exception.Message)" }
        $script:AccountIdHash=$saved; return $null
    }
}

function Get-GlucoseData {
    if (-not $script:AuthToken) { if (-not (Invoke-LibreLogin)) { return $null } }
    if (-not $script:PatientId) { if (-not (Get-Connections)) { return $null } }
    try {
        $r = Invoke-RestMethod -Uri "$($script:Config.ApiUrl)/llu/connections/$($script:PatientId)/graph" -Method GET -Headers (Get-ApiHeaders -WithAuth)
        if ($r.status -eq 0 -and $r.data) {
            $result = @{CurrentGlucose=$null;Trend=$null;Timestamp=$null;GraphData=@();PatientName=""}

            # Probuj wiele sciezek do pomiaru - Libre 2, 2 Plus, 3 moga roznic sie struktura
            $gm = $null
            if ($r.data.connection -and $r.data.connection.glucoseMeasurement) {
                $gm = $r.data.connection.glucoseMeasurement
            } elseif ($r.data.connection -and $r.data.connection.sensor -and $r.data.connection.sensor.pt) {
                $gm = $r.data.connection.sensor.pt | Select-Object -Last 1
            }

            if ($gm) {
                # Probuj rozne nazwy pola wartosci (rozne wersje API)
                $val = $null
                foreach ($field in @('ValueInMgPerDl','Value','GlucoseValue','value','valueInMgPerDl')) {
                    if ($gm.$field -and [double]$gm.$field -gt 20) { $val = [double]$gm.$field; break }
                }
                if ($val) {
                    $result.CurrentGlucose = $val
                    $result.Trend     = if ($gm.TrendArrow)  { $gm.TrendArrow  } elseif ($gm.trendArrow)  { $gm.trendArrow  } else { 0 }
                    $result.Timestamp = if ($gm.Timestamp)   { $gm.Timestamp   } elseif ($gm.timestamp)   { $gm.timestamp   } else { $null }
                }
            } else {
                # Libre 2 Plus / 3: logi diagnostyczne gdy brak glucoseMeasurement
                Write-Log "Libre2Plus diag: connection=$(($r.data.connection | ConvertTo-Json -Depth 1 -Compress).Substring(0,[Math]::Min(300,($r.data.connection|ConvertTo-Json -Depth 1 -Compress).Length)))"
            }

            # graphData - probuj standardowa sciezke i alternatywna (Libre 2 Plus moze uzywac 'data')
            $gd = if ($r.data.graphData) { $r.data.graphData } elseif ($r.data.data) { $r.data.data } else { $null }
            if ($gd) {
                $result.GraphData = @($gd)
                if ($gm -and $result.CurrentGlucose) {
                    $gmTs   = if ($gm.Timestamp) { $gm.Timestamp } else { $gm.timestamp }
                    $lastTs = if ($result.GraphData.Count -gt 0) { $result.GraphData[-1].Timestamp } else { $null }
                    if ($gmTs -ne $lastTs) { $result.GraphData += $gm }
                }
            }

            if ($r.data.connection) { $result.PatientName="$($r.data.connection.firstName) $($r.data.connection.lastName)".Trim() }
            return $result
        }
        Write-Log "Graph: non-zero status=$($r.status), keeping session"; return $null
    } catch {
        Write-Log "Graph ERR: $($_.Exception.Message)"
        if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 429) {
            $script:LastApiError = "429"
            $script:BackoffUntil = (Get-Date).AddMinutes(10)
            Write-Log "429 Too Many Requests - backoff 10 minut do $($script:BackoffUntil.ToString('HH:mm:ss'))"
        } else {
            $script:LastApiError = $null
        }
        if($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 401){$script:AuthToken=$null;$script:PatientId=$null}
        return $null
    }
}

function Save-LogbookEntries([array]$entries) {
    if (-not $entries -or $entries.Count -eq 0) { return }
    # Inicjuj HashSet przy pierwszym wywolaniu
    if ($null -eq $script:LogbookKnownTs) {
        $script:LogbookKnownTs = [System.Collections.Generic.HashSet[string]]::new()
        if (Test-Path $script:LogbookFile) {
            try {
                Get-Content $script:LogbookFile -Encoding UTF8 | ForEach-Object {
                    try { $script:LogbookKnownTs.Add(($_ | ConvertFrom-Json).ts.Substring(0,16)) | Out-Null } catch {}
                }
            } catch {}
        }
    }
    $fmts = @("M/d/yyyy h:mm:ss tt","d/M/yyyy h:mm:ss tt","M/d/yyyy H:mm:ss","d/M/yyyy H:mm:ss","yyyy-MM-ddTHH:mm:ss")
    $toAdd = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $entries) {
        $mg = 0.0
        foreach ($fld in @('ValueInMgPerDl','Value','GlucoseValue','value','valueInMgPerDl')) {
            if ($entry.$fld -and [double]$entry.$fld -gt 20) { $mg = [double]$entry.$fld; break }
        }
        if ($mg -le 0) { continue }
        $tsStr = if ($entry.Timestamp) { "$($entry.Timestamp)" } elseif ($entry.timestamp) { "$($entry.timestamp)" } else { $null }
        if (-not $tsStr) { continue }
        $parsed = [DateTime]::MinValue
        foreach ($f in $fmts) {
            if ([DateTime]::TryParseExact($tsStr, $f, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { break }
        }
        if ($parsed -eq [DateTime]::MinValue) { continue }
        $key = $parsed.ToString("yyyy-MM-ddTHH:mm")
        if ($script:LogbookKnownTs.Contains($key)) { continue }
        $isScan = try { [int]$entry.type -eq 1 } catch { $false }
        $trend  = try { [int]$entry.TrendArrow } catch { 0 }
        $toAdd.Add('{"ts":"' + $parsed.ToString("yyyy-MM-ddTHH:mm:ss") + '","mgdl":' + [Math]::Round($mg,1) + ',"trend":' + $trend + ',"scan":' + ($isScan.ToString().ToLower()) + '}')
        $script:LogbookKnownTs.Add($key) | Out-Null
    }
    if ($toAdd.Count -gt 0) {
        try { Add-Content -Path $script:LogbookFile -Value $toAdd -Encoding UTF8 } catch {}
        Write-Log "Logbook: zapisano $($toAdd.Count) nowych wpisow"
    }
}

function Load-LogbookLocal {
    $result = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path $script:LogbookFile)) { return $result }
    try {
        Get-Content $script:LogbookFile -Encoding UTF8 | ForEach-Object {
            try {
                $obj = $_ | ConvertFrom-Json
                $result.Add([PSCustomObject]@{
                    dt     = [DateTime]::Parse($obj.ts)
                    mg     = [double]$obj.mgdl
                    trend  = [int]$obj.trend
                    isScan = [bool]$obj.scan
                })
            } catch {}
        }
    } catch {}
    return $result
}

function Get-LogbookData {
    if (-not $script:AuthToken) { if (-not (Invoke-LibreLogin)) { return $null } }
    if (-not $script:PatientId) { if (-not (Get-Connections)) { return $null } }
    try {
        $r = Invoke-RestMethod -Uri "$($script:Config.ApiUrl)/llu/connections/$($script:PatientId)/logbook" `
            -Method GET -Headers (Get-ApiHeaders -WithAuth)
        if ($r.status -eq 0 -and $r.data) {
            Save-LogbookEntries @($r.data)   # zapisz do logbook.jsonl
            return @($r.data)
        }
        return @()
    } catch {
        Write-Log "Logbook ERR: $($_.Exception.Message)"
        return $null
    }
}

# ======================== HELPERS ========================
# TrendArrow: 0=NotDetermined
#   1=FallingQuickly(↓↓)  2=Falling(↓)  6=SlightlyFalling(↘)
#   3=Stable(→)
#   4=SlightlyRising(↗)   7=Rising(↑)   5=RisingQuickly(↑↑)
function Get-TrendArrow([int]$v){
    switch($v){
        1 { [char]0x2193+[char]0x2193 }   # ↓↓
        2 { [char]0x2193              }   # ↓
        6 { [char]0x2198              }   # ↘
        3 { [char]0x2192              }   # →
        4 { [char]0x2197              }   # ↗
        7 { [char]0x2191              }   # ↑
        5 { [char]0x2191+[char]0x2191 }   # ↑↑
        default { "?" }
    }
}
function Get-TrendText([int]$v){
    $li=[int]$script:LangEn
    switch($v){
        1 { @("Szybki spadek",    "Rapidly falling")[$li] }
        2 { @("Spadek",           "Falling")[$li]         }
        6 { @("Delikatny spadek", "Slightly falling")[$li]}
        3 { @("Stabilny",         "Stable")[$li]          }
        4 { @("Delikatny wzrost", "Slightly rising")[$li] }
        7 { @("Wzrost",           "Rising")[$li]          }
        5 { @("Szybki wzrost",    "Rapidly rising")[$li]  }
        default { "---" }
    }
}
function Get-CalculatedTrend([array]$graphData) {
    # Ważona regresja liniowa – jak Juggluco/xDrip.
    # PRIORYTET 1: in-memory bufor RecentReadings (rozdzielczość ~60s, zero opóźnienia API).
    # PRIORYTET 2: graphData z API (fallback gdy bufor za mały).

    $tList  = [System.Collections.Generic.List[double]]::new()
    $mgList = [System.Collections.Generic.List[double]]::new()

    if ($script:RecentReadings -and $script:RecentReadings.Count -ge 3) {
        # --- Źródło 1: bufor in-memory ---
        $buf  = $script:RecentReadings | Select-Object -Last 15
        $tRef = $buf[0].ts
        foreach ($r in $buf) {
            if ($r.mgdl -le 10) { continue }
            $tList.Add( ($r.ts - $tRef).TotalMinutes )
            $mgList.Add($r.mgdl)
        }
    }

    if ($tList.Count -lt 3 -and $graphData -and $graphData.Count -ge 3) {
        # --- Źródło 2: fallback – graphData z API ---
        $tList.Clear(); $mgList.Clear()
        $pts  = $graphData | Select-Object -Last 10
        $tRef = $null
        foreach ($pt in $pts) {
            $mg = try { [double]$pt.ValueInMgPerDl } catch { 0 }
            if ($mg -le 10) { $mg = try { [double]$pt.Value } catch { 0 } }
            if ($mg -le 10) { continue }
            $tsRaw = if ($pt.Timestamp) { $pt.Timestamp } else { $pt.FactoryTimestamp }
            $ts    = try { [DateTime]::Parse([string]$tsRaw) } catch { continue }
            if (-not $tRef) { $tRef = $ts }
            $tList.Add( ($ts - $tRef).TotalMinutes )
            $mgList.Add($mg)
        }
    }

    $n = $tList.Count
    if ($n -lt 3) { return $null }

    # Wagi wykładnicze – półokres 8 min (nowsze 2× ważniejsze od punktu sprzed 8 min)
    $tMax = $tList[$n - 1]
    $wArr = [double[]]::new($n)
    $wSum = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $wArr[$i] = [Math]::Exp(-0.0866 * ($tMax - $tList[$i]))
        $wSum    += $wArr[$i]
    }

    # Ważone średnie
    $tMean  = 0.0; $mgMean = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $w       = $wArr[$i] / $wSum
        $tMean  += $w * $tList[$i]
        $mgMean += $w * $mgList[$i]
    }

    # Ważona regresja liniowa: slope [mg/dL/min]
    $num = 0.0; $den = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $w    = $wArr[$i]
        $dt   = $tList[$i] - $tMean
        $dmg  = $mgList[$i] - $mgMean
        $num += $w * $dt * $dmg
        $den += $w * $dt * $dt
    }
    if ($den -lt 0.001) { return 3 }
    $slope = $num / $den   # mg/dL per minute

    # 7 progów (mg/dL/min) – progi celowo szersze niż minimum,
    # bo histereza (niżej) jest głównym mechanizmem stabilności:
    #  ↓↓  < -2.0          szybki spadek      (>2 mg/dL/min)
    #  ↓   -2.0 .. -1.2    spadek
    #  ↘   -1.2 .. -0.5    delikatny spadek
    #  →   -0.5 .. +0.5    stabilny           (strefa 1 mg/dL/min)
    #  ↗   +0.5 .. +1.2    delikatny wzrost
    #  ↑   +1.2 .. +2.0    wzrost
    #  ↑↑  > +2.0          szybki wzrost
    if    ($slope -lt -2.0) { return 1 }
    elseif($slope -lt -1.2) { return 2 }
    elseif($slope -lt -0.5) { return 6 }
    elseif($slope -le  0.5) { return 3 }
    elseif($slope -le  1.2) { return 4 }
    elseif($slope -le  2.0) { return 7 }
    else                    { return 5 }
}

# Zwraca tempo zmiany glukozy w mg/dL/min dla prognozy "za 30 min"
# Odporne na szum CGM: odrzuca outliery (MAD), dluzszy half-life niz Get-CalculatedTrend
function Get-GlucoseRateMgMin {
    $tList  = [System.Collections.Generic.List[double]]::new()
    $mgList = [System.Collections.Generic.List[double]]::new()

    # Zbierz surowe punkty - wiecej niz w Get-CalculatedTrend (20 zamiast 15)
    if ($script:RecentReadings -and $script:RecentReadings.Count -ge 3) {
        $buf  = $script:RecentReadings | Select-Object -Last 20
        $tRef = $buf[0].ts
        foreach ($r in $buf) {
            if ($r.mgdl -le 10) { continue }
            $tList.Add( ($r.ts - $tRef).TotalMinutes )
            $mgList.Add($r.mgdl)
        }
    }
    if ($tList.Count -lt 3 -and $script:CachedGraphData -and $script:CachedGraphData.Count -ge 3) {
        $tList.Clear(); $mgList.Clear()
        $pts  = $script:CachedGraphData | Select-Object -Last 12
        $tRef = $null
        foreach ($pt in $pts) {
            $mg  = try { [double]$pt.ValueInMgPerDl } catch { 0 }
            if ($mg -le 20) { continue }
            $tsRaw = if ($pt.Timestamp) { $pt.Timestamp } else { $null }
            if (-not $tsRaw) { continue }
            $ts = try { [DateTime]::Parse([string]$tsRaw) } catch { continue }
            if (-not $tRef) { $tRef = $ts }
            $tList.Add( ($ts - $tRef).TotalMinutes )
            $mgList.Add($mg)
        }
    }
    $n = $tList.Count; if ($n -lt 3) { return $null }

    # Odrzuc outliery metoda MAD (Median Absolute Deviation)
    # Typowy szum CGM: ±5-10 mg/dL; odrzucamy punkty >3*MAD od mediany
    $sorted = $mgList.ToArray() | Sort-Object
    $median = if ($n % 2 -eq 1) { $sorted[[int]($n/2)] } else { ($sorted[$n/2-1] + $sorted[$n/2]) / 2.0 }
    $absDevs = $mgList | ForEach-Object { [Math]::Abs($_ - $median) }
    $adSorted = @($absDevs) | Sort-Object
    $mad = if ($n % 2 -eq 1) { $adSorted[[int]($n/2)] } else { ($adSorted[$n/2-1] + $adSorted[$n/2]) / 2.0 }
    $threshold = [Math]::Max(8.0, $mad * 3.0)   # min prog 8 mg/dL (0.44 mmol) - sensor noise floor

    $tClean  = [System.Collections.Generic.List[double]]::new()
    $mgClean = [System.Collections.Generic.List[double]]::new()
    for ($i = 0; $i -lt $n; $i++) {
        if ([Math]::Abs($mgList[$i] - $median) -le $threshold) {
            $tClean.Add($tList[$i]); $mgClean.Add($mgList[$i])
        }
    }
    $n = $tClean.Count; if ($n -lt 3) { return $null }

    # Wagi wykladnicze - half-life 15 min (vs 8 min w Get-CalculatedTrend)
    # Dluzszy half-life = mniejszy wplyw chwilowych skokow sensora na prognoze
    $tMax = $tClean[$n-1]; $wArr = [double[]]::new($n); $wSum = 0.0
    for ($i=0; $i -lt $n; $i++) { $wArr[$i]=[Math]::Exp(-0.0462*($tMax-$tClean[$i])); $wSum+=$wArr[$i] }
    $tMean=0.0; $mgMean=0.0
    for ($i=0; $i -lt $n; $i++) { $w=$wArr[$i]/$wSum; $tMean+=$w*$tClean[$i]; $mgMean+=$w*$mgClean[$i] }
    $num=0.0; $den=0.0
    for ($i=0; $i -lt $n; $i++) {
        $w=$wArr[$i]
        $num += $w * ($tClean[$i]-$tMean) * ($mgClean[$i]-$mgMean)
        $den += $w * ($tClean[$i]-$tMean) * ($tClean[$i]-$tMean)
    }
    if ($den -lt 0.001) { return 0.0 }
    return $num / $den   # mg/dL per minute
}

function Get-TrendColor([int]$v) {
    switch ($v) {
        1 { "#FF3333" }   # ↓↓ Szybki spadek    - czerwony
        2 { "#FF8800" }   # ↓  Spadek            - pomarańczowy
        6 { "#FFCC44" }   # ↘  Delikatny spadek  - żółty
        3 { "#44DD44" }   # →  Stabilny          - zielony
        4 { "#FFCC44" }   # ↗  Delikatny wzrost  - żółty
        7 { "#FFAA00" }   # ↑  Wzrost            - pomarańczowy
        5 { "#FF3333" }   # ↑↑ Szybki wzrost     - czerwony
        default { "#AAAAAA" }
    }
}
function Get-GlucoseColor([double]$mmol){
    if($mmol -lt 3.0){"#FF0000"}elseif($mmol -lt 3.9){"#FF6600"}elseif($mmol -le 10.0){"#00CC00"}elseif($mmol -le 13.9){"#FFAA00"}else{"#FF0000"}
}
function Get-GlucoseStatus([double]$mmol){
    $li=[int]$script:LangEn
    if($mmol -lt 3.0)     { @("BARDZO NISKI!","VERY LOW!")[$li] }
    elseif($mmol -lt 3.9) { @("Niski","Low")[$li] }
    elseif($mmol -le 10.0){ @("W normie","In range")[$li] }
    elseif($mmol -le 13.9){ @("Wysoki","High")[$li] }
    else                  { @("BARDZO WYSOKI!","VERY HIGH!")[$li] }
}

# ======================== CONFIG SAVE/LOAD ========================
function Save-Config {
    $dir = Split-Path $script:ConfigFile
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $encPass = $script:Config.Password | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    $lines = @(
        "Email=$($script:Config.Email)"
        "EncryptedPassword=$encPass"
        "Interval=$($script:Config.Interval)"
        "AlertLow=$($script:Config.AlertLow)"
        "AlertHigh=$($script:Config.AlertHigh)"
        "LangEn=$($script:LangEn)"
        "UseMgDl=$($script:UseMgDl)"
        "SmoothMode=$($script:SmoothMode)"
    )
    # Pozycje okien (persistentne miedzy restartami)
    if ($null -ne $script:FullLeft)       { $lines += "WinFullLeft=$($script:FullLeft)" }
    if ($null -ne $script:FullTop)        { $lines += "WinFullTop=$($script:FullTop)" }
    if ($null -ne $script:CompactLeft)    { $lines += "WinCompactLeft=$($script:CompactLeft)" }
    if ($null -ne $script:CompactTop)     { $lines += "WinCompactTop=$($script:CompactTop)" }
    if ($null -ne $script:WindowOpacity)  { $lines += "WinOpacity=$($script:WindowOpacity)" }
    if ($null -ne $script:CompactTopMost) { $lines += "WinCompactTopMost=$($script:CompactTopMost)" }
    if ($null -ne $script:HistWinLeft)    { $lines += "HistWinLeft=$($script:HistWinLeft)" }
    if ($null -ne $script:HistWinTop)     { $lines += "HistWinTop=$($script:HistWinTop)" }
    if ($null -ne $script:AvgBarsWinLeft) { $lines += "AvgBarsWinLeft=$($script:AvgBarsWinLeft)" }
    if ($null -ne $script:AvgBarsWinTop)  { $lines += "AvgBarsWinTop=$($script:AvgBarsWinTop)" }
    if ($null -ne $script:AgpWinLeft)     { $lines += "AgpWinLeft=$($script:AgpWinLeft)" }
    if ($null -ne $script:AgpWinTop)      { $lines += "AgpWinTop=$($script:AgpWinTop)" }
    $lines | Set-Content -Path $script:ConfigFile -Encoding UTF8
}

function Load-Config {
    if (Test-Path $script:ConfigFile) {
        $lines = Get-Content $script:ConfigFile -Encoding UTF8
        $encPass = $null
        foreach ($line in $lines) {
            if ($line -match '^Email=(.+)$')             { $script:Config.Email    = $Matches[1] }
            if ($line -match '^Password=(.+)$')          { $script:Config.Password = $Matches[1] }  # stary format - plain text
            if ($line -match '^EncryptedPassword=(.+)$') { $encPass = $Matches[1] }
            if ($line -match '^Interval=(\d+)$')         { $script:Config.Interval  = [int]$Matches[1] }
            if ($line -match '^AlertLow=(.+)$')          { try { $script:Config.AlertLow  = [double]$Matches[1] } catch {} }
            if ($line -match '^AlertHigh=(.+)$')         { try { $script:Config.AlertHigh = [double]$Matches[1] } catch {} }
            if ($line -match '^LangEn=(.+)$')            { try { $script:LangEn    = [bool]::Parse($Matches[1]) } catch {} }
            if ($line -match '^UseMgDl=(.+)$')           { try { $script:UseMgDl   = [bool]::Parse($Matches[1]) } catch {} }
            if ($line -match '^SmoothMode=(.+)$')        { try { $script:SmoothMode = [bool]::Parse($Matches[1]) } catch {} }
            # Pozycje okien
            if ($line -match '^WinFullLeft=(.+)$')       { try { $script:FullLeft       = [double]$Matches[1] } catch {} }
            if ($line -match '^WinFullTop=(.+)$')        { try { $script:FullTop        = [double]$Matches[1] } catch {} }
            if ($line -match '^WinCompactLeft=(.+)$')    { try { $script:CompactLeft    = [double]$Matches[1] } catch {} }
            if ($line -match '^WinCompactTop=(.+)$')     { try { $script:CompactTop     = [double]$Matches[1] } catch {} }
            if ($line -match '^WinOpacity=(.+)$')        { try { $script:WindowOpacity  = [double]$Matches[1] } catch {} }
            if ($line -match '^WinCompactTopMost=(.+)$') { try { $script:CompactTopMost = [bool]::Parse($Matches[1]) } catch {} }
            if ($line -match '^HistWinLeft=(.+)$')       { try { $script:HistWinLeft    = [double]$Matches[1] } catch {} }
            if ($line -match '^HistWinTop=(.+)$')        { try { $script:HistWinTop     = [double]$Matches[1] } catch {} }
            if ($line -match '^AvgBarsWinLeft=(.+)$')    { try { $script:AvgBarsWinLeft = [double]$Matches[1] } catch {} }
            if ($line -match '^AvgBarsWinTop=(.+)$')     { try { $script:AvgBarsWinTop  = [double]$Matches[1] } catch {} }
            if ($line -match '^AgpWinLeft=(.+)$')        { try { $script:AgpWinLeft     = [double]$Matches[1] } catch {} }
            if ($line -match '^AgpWinTop=(.+)$')         { try { $script:AgpWinTop      = [double]$Matches[1] } catch {} }
        }
        if ($encPass) {
            try {
                $ss = $encPass | ConvertTo-SecureString
                $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
                $script:Config.Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            } catch { Write-Log "Blad deszyfrowania hasla: $($_.Exception.Message)" }
        }
        return ($script:Config.Email -ne "" -and $script:Config.Password -ne "")
    }
    return $false
}

# ======================== OKNO LOGOWANIA ========================
function Show-LoginWindow {
    [xml]$loginXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Glucose Monitor - Logowanie" Width="280" Height="260"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Topmost="True" Background="Transparent"
        WindowStyle="None" AllowsTransparency="True">
    <Border Background="#1a1a2e" CornerRadius="10">
        <Grid Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Grid.Row="0" Text="Glucose Monitor" Foreground="White" FontSize="16"
                       FontFamily="Segoe UI Semibold" HorizontalAlignment="Center" Margin="0,0,0,4"/>
            <TextBlock Grid.Row="1" Text="Zaloguj sie kontem LibreView" Foreground="#7777aa" FontSize="10"
                       HorizontalAlignment="Center" Margin="0,0,0,16"/>

            <TextBlock Grid.Row="2" Text="Email" Foreground="#7777aa" FontSize="10" Margin="0,0,0,4"/>
            <TextBox Grid.Row="3" Name="loginEmail" Background="#2a2a4a" Foreground="White"
                     BorderBrush="#3a3a6a" BorderThickness="1" Padding="6,4" FontSize="12"
                     CaretBrush="White" Margin="0,0,0,10"/>

            <TextBlock Grid.Row="4" Text="Haslo" Foreground="#7777aa" FontSize="10" Margin="0,0,0,4"/>
            <PasswordBox Grid.Row="5" Name="loginPassword" Background="#2a2a4a" Foreground="White"
                         BorderBrush="#3a3a6a" BorderThickness="1" Padding="6,4" FontSize="12"
                         CaretBrush="White" Margin="0,0,0,4"/>

            <TextBlock Grid.Row="6" Name="loginError" Text="" Foreground="#FF6666" FontSize="10"
                       HorizontalAlignment="Center" Margin="0,4,0,0"/>

            <Grid Grid.Row="7" Margin="0,10,0,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/><ColumnDefinition Width="8"/><ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button Grid.Column="0" Name="loginCancel" Content="Anuluj"
                        Background="#2a2a4a" Foreground="#aaaacc" BorderThickness="0"
                        Padding="0,6" FontSize="11" Cursor="Hand"/>
                <Button Grid.Column="2" Name="loginOk" Content="Zaloguj"
                        Background="#3a5a9a" Foreground="White" BorderThickness="0"
                        Padding="0,6" FontSize="11" Cursor="Hand" FontFamily="Segoe UI Semibold"/>
            </Grid>
        </Grid>
    </Border>
</Window>
"@
    $lr = New-Object System.Xml.XmlNodeReader $loginXaml
    $lw = [Windows.Markup.XamlReader]::Load($lr)

    $loginEmail    = $lw.FindName("loginEmail")
    $loginPassword = $lw.FindName("loginPassword")
    $loginError    = $lw.FindName("loginError")
    $loginOk       = $lw.FindName("loginOk")
    $loginCancel   = $lw.FindName("loginCancel")

    # Wypelnij jesli cos juz jest w config
    $loginEmail.Text = $script:Config.Email

    # Drag okna - tylko gdy nie kliknieto w przycisk
    $lw.Add_MouseLeftButtonDown({
        param($sender, $e)
        if ($e.OriginalSource -isnot [System.Windows.Controls.Button] -and
            $e.OriginalSource -isnot [System.Windows.Controls.TextBox] -and
            $e.OriginalSource -isnot [System.Windows.Controls.PasswordBox] -and
            $e.OriginalSource.TemplatedParent -isnot [System.Windows.Controls.Button]) {
            $lw.DragMove()
        }
    })

    $loginCancel.Add_Click({ $lw.DialogResult = $false; $lw.Close() })

    $loginOk.Add_Click({
        if ($loginEmail.Text.Trim() -eq "" -or $loginPassword.Password -eq "") {
            $loginError.Text = "Wpisz email i haslo"
            return
        }
        $script:Config.Email    = $loginEmail.Text.Trim()
        $script:Config.Password = $loginPassword.Password
        $script:Config.ApiUrl   = "https://api-eu.libreview.io"
        $script:AuthToken = $null; $script:PatientId = $null
        $loginError.Text = "Laczenie..."
        $lw.IsEnabled = $false
        $ok = Invoke-LibreLogin
        if ($ok) {
            Save-Config
            $lw.DialogResult = $true
            $lw.Close()
        } else {
            $loginError.Text = "Nieprawidlowy email lub haslo"
            $lw.IsEnabled = $true
        }
    })

    # Enter w polu hasla odpala logowanie
    $loginPassword.Add_KeyDown({
        param($s,$e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) { $loginOk.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
    })

    return $lw.ShowDialog()
}

# ======================== GUI ========================
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Glucose Monitor" Width="300" Height="420"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Topmost="True" Background="Transparent"
        WindowStyle="None" AllowsTransparency="True">
    <Window.Resources>
        <Style x:Key="L" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#7777aa"/><Setter Property="FontSize" Value="10"/><Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        <Style x:Key="V" TargetType="TextBlock">
            <Setter Property="Foreground" Value="White"/><Setter Property="FontSize" Value="13"/><Setter Property="FontFamily" Value="Segoe UI Semibold"/>
        </Style>
    </Window.Resources>
    <Border Background="#1a1a2e" CornerRadius="10">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="30"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Pasek tytulowy -->
            <Border Grid.Row="0" Name="titleBar" Background="#12122a" CornerRadius="10,10,0,0">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Name="txtTitle" Text="  Glucose Monitor" Foreground="#7777aa" FontSize="11"
                               VerticalAlignment="Center" FontFamily="Segoe UI" Margin="6,0,0,0"/>
                    <Button Grid.Column="1" Name="btnCompact" Content="&#x229F;" Width="34" Height="30"
                            Background="Transparent" Foreground="#7777aa" BorderThickness="0" FontSize="13" Cursor="Hand"/>
                    <Button Grid.Column="2" Name="btnMinimize" Content="&#x2014;" Width="34" Height="30"
                            Background="Transparent" Foreground="#7777aa" BorderThickness="0" FontSize="13" Cursor="Hand"/>
                    <Button Grid.Column="3" Name="btnClose" Content="&#x2715;" Width="34" Height="30"
                            Background="Transparent" Foreground="#aa5555" BorderThickness="0" FontSize="13" Cursor="Hand"/>
                </Grid>
            </Border>

            <!-- Nazwa + status -->
            <StackPanel Grid.Row="1" Margin="12,6,12,4">
                <TextBlock Name="txtPatient" Text="Glucose Monitor" Foreground="White" FontSize="14"
                           FontFamily="Segoe UI Semibold" HorizontalAlignment="Center"/>
                <TextBlock Name="txtStatus" Text="Laczenie..." Foreground="#7777aa" FontSize="10"
                           HorizontalAlignment="Center"/>
            </StackPanel>

            <!-- Glowny odczyt -->
            <Border Grid.Row="2" CornerRadius="12" Padding="14,8" Margin="12,0,12,6" Background="#2a2a4a">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" VerticalAlignment="Center">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock Name="txtGlucoseValue" Text="---" Foreground="#00CC00"
                                       FontSize="38" FontFamily="Segoe UI Bold" VerticalAlignment="Bottom"/>
                            <TextBlock Name="txtTrendArrow" Text="" Foreground="White"
                                       FontSize="24" FontFamily="Segoe UI" Margin="6,0,0,4" VerticalAlignment="Bottom"/>
                        </StackPanel>
                        <TextBlock Name="txtTrendText" Text="" Foreground="#FFCC44" FontSize="10"
                                   FontFamily="Segoe UI Semibold" Margin="2,1,0,1"/>
                        <TextBlock Name="txtUnitLabel" Text="mmol/L" Foreground="#7777aa" FontSize="10" Margin="2,0,0,0"/>
                        <TextBlock Name="txtForecast" Text="" Foreground="#7777aa" FontSize="10"
                                   FontFamily="Segoe UI" Margin="2,2,0,0"/>
                    </StackPanel>
                    <StackPanel Grid.Column="1" VerticalAlignment="Center" HorizontalAlignment="Right">
                        <TextBlock Name="txtGlucoseStatus" Text="" Foreground="#aaaacc" FontSize="12"
                                   HorizontalAlignment="Right"/>
                        <TextBlock Name="txtTimestamp" Text="---" Foreground="#7777aa" FontSize="10"
                                   HorizontalAlignment="Right" Margin="0,2,0,0"/>
                        <TextBlock Name="txtDelta" Text="" Foreground="#aaaacc" FontSize="11"
                                   HorizontalAlignment="Right" Margin="0,1,0,0"/>
                    </StackPanel>
                </Grid>
            </Border>

            <!-- Stats bar: Min / Sred / Max / eHbA1c (trend przeniesiony do karty glukozy) -->
            <Border Grid.Row="3" CornerRadius="8" Padding="10,6" Background="#2a2a4a" Margin="12,0,12,4">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Min" Style="{StaticResource L}"/>
                            <TextBlock Text=" 12h" Foreground="#8888bb" FontSize="9" VerticalAlignment="Bottom" Margin="2,0,0,1" FontFamily="Segoe UI"/>
                        </StackPanel>
                        <TextBlock Name="txtMin" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Name="lblSred" Text="Sred." Style="{StaticResource L}"/>
                            <TextBlock Text=" 12h" Foreground="#8888bb" FontSize="9" VerticalAlignment="Bottom" Margin="2,0,0,1" FontFamily="Segoe UI"/>
                        </StackPanel>
                        <TextBlock Name="txtAvg" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Max" Style="{StaticResource L}"/>
                            <TextBlock Text=" 12h" Foreground="#8888bb" FontSize="9" VerticalAlignment="Bottom" Margin="2,0,0,1" FontFamily="Segoe UI"/>
                        </StackPanel>
                        <TextBlock Name="txtMax" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="3" HorizontalAlignment="Right">
                        <TextBlock Text="eHbA1c" Style="{StaticResource L}"/>
                        <TextBlock Name="txtHbA1c" Text="---" Foreground="#cc88ff" FontSize="11"
                                   FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                </Grid>
            </Border>

            <!-- WYKRES -->
            <Border Grid.Row="4" CornerRadius="8" Padding="8,6" Background="#222244" Margin="12,2,12,4">
                <Canvas Name="canvasGraph" ClipToBounds="True"/>
            </Border>

            <!-- Dolny pasek -->
            <Grid Grid.Row="5" Margin="12,0,12,8">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="4"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="4"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="4"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="5"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Name="txtNextUpdate" Text="" Foreground="White" FontSize="11" VerticalAlignment="Center"/>
                <Button Grid.Column="1" Name="btnHist" Content="HIST" Background="#1e2a3a"
                        Foreground="#6677aa" BorderThickness="0" Padding="6,4" FontSize="9" Cursor="Hand"
                        ToolTip="Historia glukozy / Glucose history"/>
                <Button Grid.Column="3" Name="btnLang" Content="ENG" Background="#1e2a3a"
                        Foreground="#6677aa" BorderThickness="0" Padding="6,4" FontSize="9" Cursor="Hand"
                        ToolTip="Zmien jezyk / Change language"/>
                <Button Grid.Column="5" Name="btnUnit" Content="mmol/L" Background="#1e2a3a"
                        Foreground="#6677aa" BorderThickness="0" Padding="6,4" FontSize="9" Cursor="Hand"
                        ToolTip="Przelacz jednostki"/>
                <Button Grid.Column="7" Name="btnSmooth" Content="~" Background="#1e2a3a"
                        Foreground="#6677aa" BorderThickness="0" Padding="6,4" FontSize="11" Cursor="Hand"
                        ToolTip="Wygładzanie danych (Savitzky-Golay) / Data smoothing"/>
                <Button Grid.Column="9" Name="btnRefresh" Content="Odswiez" Background="#3a3a5a"
                        Foreground="White" BorderThickness="0" Padding="10,4" FontSize="10" Cursor="Hand"/>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Zapewnij Application z OnExplicitShutdown - konieczne zeby Hide() nie konczylo programu
# ShutdownMode ustawiamy tylko gdy sami tworzymy Application - jesli istnieje (poprzednia sesja),
# jest juz ustawiony i probka zapisu z innego watku powoduje wyjatek cross-thread
if (-not [System.Windows.Application]::Current) {
    $script:App = [System.Windows.Application]::new()
    [System.Windows.Application]::Current.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
}
[System.Windows.Application]::Current.Add_DispatcherUnhandledException({
    param($s, $e)
    Write-Log "UNHANDLED EXCEPTION: $($e.Exception.Message)"
    $e.Handled = $true
})

$txtPatient=$window.FindName("txtPatient"); $txtStatus=$window.FindName("txtStatus")
$txtGlucoseValue=$window.FindName("txtGlucoseValue"); $txtTrendArrow=$window.FindName("txtTrendArrow")
$txtGlucoseStatus=$window.FindName("txtGlucoseStatus"); $txtTrendText=$window.FindName("txtTrendText")
$txtForecast=$window.FindName("txtForecast")
$txtDelta=$window.FindName("txtDelta")
$txtTimestamp=$window.FindName("txtTimestamp"); $txtMin=$window.FindName("txtMin")
$txtAvg=$window.FindName("txtAvg"); $txtMax=$window.FindName("txtMax")
$txtHbA1c=$window.FindName("txtHbA1c")
$txtNextUpdate=$window.FindName("txtNextUpdate"); $btnRefresh=$window.FindName("btnRefresh")
$canvasGraph=$window.FindName("canvasGraph")
$titleBar=$window.FindName("titleBar")
$btnMinimize=$window.FindName("btnMinimize")
$btnClose=$window.FindName("btnClose")
$btnCompact=$window.FindName("btnCompact")
$txtTitle=$window.FindName("txtTitle")
$btnUnit=$window.FindName("btnUnit")
$txtUnitLabel=$window.FindName("txtUnitLabel")
$btnLang=$window.FindName("btnLang")
$btnHist=$window.FindName("btnHist")
$btnSmooth=$window.FindName("btnSmooth")
$lblSred=$window.FindName("lblSred")

# Drag okna za pasek tytulowy (tylko lewy przycisk, nie na buttonach)
$titleBar.Add_MouseLeftButtonDown({
    param($sender, $e)
    if ($e.OriginalSource -isnot [System.Windows.Controls.Button] -and
        $e.OriginalSource.TemplatedParent -isnot [System.Windows.Controls.Button]) {
        $window.DragMove()
    }
})
# Przyciski
$btnMinimize.Add_Click({
    $window.Hide()
})
$btnClose.Add_Click({
    $window.Close()
    exit
})

# Compact mode toggle
$script:IsCompact = $false
$script:CompactLeft = $null
$script:CompactTop  = $null
$script:FullLeft    = $null
$script:FullTop     = $null
$script:MainGrid = $window.Content.Child  # Grid inside Border

# Find rows 1-5 (all except title bar row 0)
$script:CompactRows = @()
for ($i = 1; $i -lt $script:MainGrid.RowDefinitions.Count; $i++) {
    $script:CompactRows += $script:MainGrid.RowDefinitions[$i]
}
# Find children in rows 1-5
$script:CompactChildren = @()
foreach ($child in $script:MainGrid.Children) {
    $row = [System.Windows.Controls.Grid]::GetRow($child)
    if ($row -ge 1) { $script:CompactChildren += $child }
}

# Compact glucose display (shown only in compact mode)
$compactLabel = New-Object System.Windows.Controls.TextBlock
$compactLabel.Name = "txtCompactGlucose"
$compactLabel.FontSize = 18
$compactLabel.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI Bold")
$compactLabel.Foreground = [System.Windows.Media.Brushes]::LimeGreen
$compactLabel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
$compactLabel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
$compactLabel.Visibility = [System.Windows.Visibility]::Collapsed
$compactLabel.Margin = [System.Windows.Thickness]::new(0,2,0,4)
$compactLabel.Cursor = [System.Windows.Input.Cursors]::Hand
[System.Windows.Controls.Grid]::SetRow($compactLabel, 1)
[System.Windows.Controls.Grid]::SetColumnSpan($compactLabel, 1)
$script:MainGrid.Children.Add($compactLabel) | Out-Null
$script:TxtCompactGlucose = $compactLabel
$script:WindowOpacity = 1.0

$window.Add_MouseWheel({
    param($s, $e)
    $delta = if ($e.Delta -gt 0) { 0.1 } else { -0.1 }
    $newVal = [Math]::Round($window.Opacity + $delta, 1)
    if ($newVal -lt 0.1) { $newVal = 0.1 }
    if ($newVal -gt 1.0) { $newVal = 1.0 }
    $window.Opacity = $newVal
    $script:WindowOpacity = $newVal
})
$script:CompactTopMost = $true  # domyslnie Topmost jest wlaczone

$compactLabel.Add_MouseLeftButtonDown({
    if (-not $script:IsCompact) { return }
    if ($script:CompactTopMost) {
        $window.Topmost = $false
        $script:CompactTopMost = $false
        $script:ForceTimer.Stop()
        $script:TxtCompactGlucose.Opacity = 1.0
        [System.Windows.MessageBox]::Show("Zawsze na wierzchu: OFF", "Glucose Monitor", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    } else {
        $window.Topmost = $true
        $script:CompactTopMost = $true
        $script:ForceTimer.Start()
        $script:TxtCompactGlucose.Opacity = 0.55
        [System.Windows.MessageBox]::Show("Zawsze na wierzchu: ON", "Glucose Monitor", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    }
})

$btnCompact.Add_Click({
    if (-not $script:IsCompact) {
        # Switch to compact - save full window position
        $script:FullLeft = $window.Left
        $script:FullTop  = $window.Top
        # CompactLeft/Top nie jest nadpisywane - zachowuje ostatnia pozycje mini okna
        # Przy pierwszym uruchomieniu inicjalizuj do aktualnej pozycji
        if ($null -eq $script:CompactLeft) { $script:CompactLeft = $window.Left }
        if ($null -eq $script:CompactTop)  { $script:CompactTop  = $window.Top  }
        foreach ($child in $script:CompactChildren) { $child.Visibility = [System.Windows.Visibility]::Collapsed }
        $script:MainGrid.RowDefinitions[1].Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        for ($i = 2; $i -lt $script:MainGrid.RowDefinitions.Count; $i++) {
            $script:MainGrid.RowDefinitions[$i].Height = [System.Windows.GridLength]::new(0)
        }
        # Update compact label
        $script:TxtCompactGlucose.Text = "$($txtGlucoseValue.Text) $($txtTrendArrow.Text)"
        $script:TxtCompactGlucose.Foreground = $txtGlucoseValue.Foreground
        $script:TxtCompactGlucose.Opacity = if ($script:CompactTopMost) { 0.55 } else { 1.0 }
        $script:TxtCompactGlucose.Visibility = [System.Windows.Visibility]::Visible
        $window.Width = 120
        $window.Height = 50
        $txtTitle.Visibility = [System.Windows.Visibility]::Collapsed
        if ($null -ne $script:CompactLeft -and $null -ne $script:CompactTop) {
            $window.Left = $script:CompactLeft
            $window.Top  = $script:CompactTop
        }
        $btnCompact.Content = [char]0x229E
        $script:IsCompact = $true
        $window.ShowInTaskbar = $false
        $window.Opacity = $script:WindowOpacity
        if ($script:CompactTopMost) { $script:ForceTimer.Start() }
    } else {
        # Restore full view - save compact position, restore full window position
        $script:CompactLeft = $window.Left
        $script:CompactTop  = $window.Top

        $screenWidth  = [System.Windows.SystemParameters]::PrimaryScreenWidth
        $screenHeight = [System.Windows.SystemParameters]::PrimaryScreenHeight
        $winTop  = $script:CompactTop
        $winLeft = $script:CompactLeft
        $fullWidth  = 300
        $fullHeight = 420

        $txtTitle.Visibility = [System.Windows.Visibility]::Visible
        $script:TxtCompactGlucose.Visibility = [System.Windows.Visibility]::Collapsed
        $window.Topmost = $true
        foreach ($child in $script:CompactChildren) { $child.Visibility = [System.Windows.Visibility]::Visible }
        $script:MainGrid.RowDefinitions[1].Height = [System.Windows.GridLength]::Auto
        $script:MainGrid.RowDefinitions[2].Height = [System.Windows.GridLength]::Auto
        $script:MainGrid.RowDefinitions[3].Height = [System.Windows.GridLength]::Auto
        $script:MainGrid.RowDefinitions[4].Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $script:MainGrid.RowDefinitions[5].Height = [System.Windows.GridLength]::Auto
        $window.Width  = $fullWidth
        $window.Height = $fullHeight

        if ($null -ne $script:FullLeft -and $null -ne $script:FullTop) {
            # Przywroc zapamietana pozycje duzego okna
            $window.Left = $script:FullLeft
            $window.Top  = $script:FullTop
        } else {
            # Pierwszy raz - smart expand od pozycji kompaktowej
            $window.Left = $winLeft
            $window.Top  = $winTop

            # Pionowo: dolna polowa ekranu - rozwin do gory
            if ($winTop -gt ($screenHeight / 2)) {
                $window.Top = $winTop - ($fullHeight - 56)
            }

            # Poziomo: prawa polowa ekranu - rozwin w lewo
            if ($winLeft -gt ($screenWidth / 2)) {
                $window.Left = $winLeft - ($fullWidth - 110)
            }
        }

        $btnCompact.Content = [char]0x229F
        $script:IsCompact = $false
        $window.ShowInTaskbar = $true
        $script:ForceTimer.Stop()
        $window.Opacity = 1.0
    }
})

# ======================== WYKRES ========================
function Apply-SavitzkyGolay([double[]]$data) {
    $n = $data.Length
    if ($n -lt 5) { return $data }
    $result = $data.Clone()
    $tmp    = $data.Clone()
    for ($i = 2; $i -lt $n - 2; $i++) {
        $result[$i] = (-3*$tmp[$i-2] + 12*$tmp[$i-1] + 17*$tmp[$i] + 12*$tmp[$i+1] - 3*$tmp[$i+2]) / 35.0
    }
    return $result
}

function Update-Graph($GraphData) {
    $canvasGraph.Children.Clear()
    if (-not $GraphData -or $GraphData.Count -lt 2) { return }
    $w=$canvasGraph.ActualWidth; $h=$canvasGraph.ActualHeight
    if ($w -le 0 -or $h -le 0) { return }

    $vals=@(); $types=@()
    foreach($i in $GraphData) {
        $mg = if($i.ValueInMgPerDl){[double]$i.ValueInMgPerDl}elseif($i.Value){[double]$i.Value}else{0}
        if($mg -gt 0) {
            if ($script:UseMgDl) { $vals += [Math]::Round($mg, 0) }
            else { $vals += MgToMmol $mg }
            $isScan = try { [int]$i.type -eq 1 } catch { $false }
            $types += $isScan
        }
    }
    if ($vals.Count -lt 2) { return }

    # Wygładzanie Savitzky-Golay (jesli wlaczone)
    if ($script:SmoothMode -and $vals.Count -ge 5) {
        $vals = Apply-SavitzkyGolay ([double[]]$vals)
    }

    # Zwiekszony margines dolny na etykiety czasu
    $m=5; $bottomMargin=22
    $dw=$w-2*$m; $dh=$h-2*$m-$bottomMargin

    # Zakres osi Y zalezy od jednostki
    if ($script:UseMgDl) {
        $mn = 70.0; $mx = 360.0
        $gridLines = @(70, 80, 120, 160, 180, 200, 250, 360)
    } else {
        $mn = 3.9; $mx = 20.0
        $gridLines = @(3.9, 4.4, 6.7, 8.9, 10.0, 11.1, 13.3, 20.0)
    }
    $rng = $mx - $mn
    # Kolorowe tlo stref (identyczne jak w oknie historii)
    $zHiC = if ($script:UseMgDl) { 250.0 } else { 13.9 }
    $zHiH = if ($script:UseMgDl) { 180.0 } else { 10.0 }
    $zLoY = if ($script:UseMgDl) {  79.0 } else {  4.4 }
    $zLoN = if ($script:UseMgDl) {  70.0 } else {  3.9 }
    foreach ($z in @(
        @{ lo=$zHiC; hi=$mx;   br=$script:CBrZoneRed }
        @{ lo=$zHiH; hi=$zHiC; br=$script:CBrZoneOrange }
        @{ lo=$zLoY; hi=$zHiH; br=$script:CBrZoneGreen }
        @{ lo=$zLoN; hi=$zLoY; br=$script:CBrZoneYellow }
        @{ lo=$mn;   hi=$zLoN; br=$script:CBrZoneRed }
    )) {
        $zHi = [Math]::Min($z.hi, $mx); $zLo = [Math]::Max($z.lo, $mn)
        if ($zHi -le $zLo) { continue }
        $yT = $m + $dh - (($dh/$rng)*($zHi - $mn))
        $yB = $m + $dh - (($dh/$rng)*($zLo - $mn))
        $zRect = New-Object System.Windows.Shapes.Rectangle
        $zRect.Fill = $z.br
        $zRect.Width = $dw; $zRect.Height = [Math]::Abs($yB - $yT)
        [System.Windows.Controls.Canvas]::SetLeft($zRect, $m)
        [System.Windows.Controls.Canvas]::SetTop($zRect, [Math]::Min($yT,$yB))
        $canvasGraph.Children.Add($zRect) | Out-Null
    }
    $labelOffset = 18  # szerokosc etykiety
    foreach($lim in $gridLines) {
        if($lim -ge $mn -and $lim -le $mx) {
            $yy=$m+$dh-(($dh/$rng)*($lim-$mn))
            $ln=New-Object System.Windows.Shapes.Line; $ln.X1=$m;$ln.X2=$m+$dw;$ln.Y1=$yy;$ln.Y2=$yy
            $ln.Stroke=$script:CBrGridLine60
            $ln.StrokeThickness=0.7
            $canvasGraph.Children.Add($ln)|Out-Null
            $tb=New-Object System.Windows.Controls.TextBlock; $tb.Text="$lim"; $tb.FontSize=8
            $tb.Foreground=$script:CBrGridLabel
            [System.Windows.Controls.Canvas]::SetLeft($tb, $m+$dw-$labelOffset)
            [System.Windows.Controls.Canvas]::SetTop($tb, $yy-7)
            $canvasGraph.Children.Add($tb)|Out-Null
        }
    }

    # Linia wykresu - kolorowe segmenty wygladzone Catmull-Rom -> cubic Bezier
    $hiH = if ($script:UseMgDl) { 180.0 } else { 10.0 }
    $hiC = if ($script:UseMgDl) { 250.0 } else { 13.9 }
    $loC = if ($script:UseMgDl) {  70.0 } else {  3.9 }
    $chartW=$dw*0.82; $step=$chartW/[Math]::Max(1,$vals.Count-1)
    $n = $vals.Count
    $px = [double[]]::new($n); $py = [double[]]::new($n)
    for ($i = 0; $i -lt $n; $i++) {
        $px[$i] = $m + ($i * $step)
        $py[$i] = $m + $dh - (($dh/$rng)*($vals[$i]-$mn))
    }
    # StreamGeometry zamiast PathGeometry - lzejsze, szybsze renderowanie
    $sgR = New-Object System.Windows.Media.StreamGeometry
    $sgO = New-Object System.Windows.Media.StreamGeometry
    $sgG = New-Object System.Windows.Media.StreamGeometry
    $ctxR = $sgR.Open(); $ctxO = $sgO.Open(); $ctxG = $sgG.Open()
    [double]$tens = 0.2
    $prevSc = ""
    for ($i = 0; $i -lt $n - 1; $i++) {
        $avg2 = ($vals[$i]+$vals[$i+1])/2.0
        $sc   = if ($avg2 -lt $loC -or $avg2 -gt $hiC) { "R" } elseif ($avg2 -gt $hiH) { "O" } else { "G" }
        $ctx  = if ($sc -eq "R") { $ctxR } elseif ($sc -eq "O") { $ctxO } else { $ctxG }
        if ($sc -ne $prevSc) {
            $ctx.BeginFigure([System.Windows.Point]::new($px[$i], $py[$i]), $false, $false)
            $prevSc = $sc
        }
        $i0 = if ($i -gt 0) { $i-1 } else { 0 }
        $i3 = if ($i+2 -lt $n) { $i+2 } else { $n-1 }
        $cp1 = [System.Windows.Point]::new($px[$i]+($px[$i+1]-$px[$i0])*$tens, $py[$i]+($py[$i+1]-$py[$i0])*$tens)
        $cp2 = [System.Windows.Point]::new($px[$i+1]-($px[$i3]-$px[$i])*$tens, $py[$i+1]-($py[$i3]-$py[$i])*$tens)
        $ctx.BezierTo($cp1, $cp2, [System.Windows.Point]::new($px[$i+1], $py[$i+1]), $true, $false)
    }
    $ctxR.Close(); $ctxO.Close(); $ctxG.Close()
    $sgR.Freeze(); $sgO.Freeze(); $sgG.Freeze()
    foreach ($item in @(@{g=$sgR;br=$script:CBrRed},@{g=$sgO;br=$script:CBrOrange},@{g=$sgG;br=$script:CBrGreen})) {
        if ($item.g.MayHaveCurves()) {
            $pe = New-Object System.Windows.Shapes.Path; $pe.Data = $item.g
            $pe.Stroke = $item.br
            $pe.StrokeThickness=2; $pe.StrokeStartLineCap="Round"; $pe.StrokeEndLineCap="Round"
            $canvasGraph.Children.Add($pe)|Out-Null
        }
    }

    # Kropki skanow (type=1) - male polprzezroczyste kola na wierzchu linii
    for($i=0; $i -lt $vals.Count; $i++) {
        if ($types[$i]) {
            $sx=$m+($i*$step); $sy=$m+$dh-(($dh/$rng)*($vals[$i]-$mn))
            $sc=New-Object System.Windows.Shapes.Ellipse; $sc.Width=5; $sc.Height=5
            $sc.Fill=$script:CBrScanFill
            $sc.Stroke=$script:CBrScanStroke
            $sc.StrokeThickness=0.5
            [System.Windows.Controls.Canvas]::SetLeft($sc,$sx-2.5)
            [System.Windows.Controls.Canvas]::SetTop($sc,$sy-2.5)
            $canvasGraph.Children.Add($sc)|Out-Null
        }
    }

    # Ostatni punkt
    $lc=$vals.Count-1
    $lastX=$m+($lc*$step); $lastY=$m+$dh-(($dh/$rng)*($vals[$lc]-$mn))
    $lastAvg=$vals[$lc]
    $dotBr = if ($lastAvg -lt $loC -or $lastAvg -gt $hiC) { $script:CBrRed } elseif ($lastAvg -gt $hiH) { $script:CBrOrange } else { $script:CBrGreen }
    $dot=New-Object System.Windows.Shapes.Ellipse; $dot.Width=8;$dot.Height=8
    $dot.Fill=$dotBr
    [System.Windows.Controls.Canvas]::SetLeft($dot,$lastX-4)
    [System.Windows.Controls.Canvas]::SetTop($dot,$lastY-4)
    $canvasGraph.Children.Add($dot)|Out-Null

    # Prognoza trendu (~30 min naprzod) - linia przerywana
    if ($vals.Count -ge 3) {
        $nP = [Math]::Min(5, $vals.Count)
        $rSlopes = @()
        for ($i = ($vals.Count - $nP); $i -lt $vals.Count - 1; $i++) { $rSlopes += ($vals[$i+1] - $vals[$i]) }
        $slope = ($rSlopes | Measure-Object -Average).Average
        $f2Val = [Math]::Max($mn, [Math]::Min($mx, $vals[$lc] + 2*$slope))
        $f2X   = $m + $dw
        $f2Y   = $m + $dh - (($dh/$rng)*($f2Val - $mn))
        $fBr   = if ($f2Val -lt $loC -or $f2Val -gt $hiC) { $script:CBrRed } elseif ($f2Val -gt $hiH) { $script:CBrOrange } else { $script:CBrGreen }
        $fl=New-Object System.Windows.Shapes.Line; $fl.X1=$lastX;$fl.Y1=$lastY;$fl.X2=$f2X;$fl.Y2=$f2Y
        $fl.Stroke=$fBr
        $fl.StrokeThickness=1.5; $fl.Opacity=0.55
        $da=New-Object System.Windows.Media.DoubleCollection; $da.Add(4);$da.Add(3); $fl.StrokeDashArray=$da
        $canvasGraph.Children.Add($fl)|Out-Null
        $fdot=New-Object System.Windows.Shapes.Ellipse; $fdot.Width=6;$fdot.Height=6; $fdot.Opacity=0.55
        $fdot.Fill=$fBr
        [System.Windows.Controls.Canvas]::SetLeft($fdot,$f2X-3); [System.Windows.Controls.Canvas]::SetTop($fdot,$f2Y-3)
        $canvasGraph.Children.Add($fdot)|Out-Null
    }

    # Etykiety czasu
    $lastIdx=$vals.Count-1
    $labelIndices = @(0, [int]($vals.Count/4), [int]($vals.Count/2), [int]($vals.Count*3/4), $lastIdx) | Sort-Object -Unique
    foreach($idx in $labelIndices) {
        if ($idx -ge 0 -and $idx -lt $GraphData.Count -and $GraphData[$idx].Timestamp) {
            try {
                $ts=[string]$GraphData[$idx].Timestamp; $parsed=[DateTime]::MinValue
                $fmts=@("M/d/yyyy h:mm:ss tt","d/M/yyyy h:mm:ss tt","M/d/yyyy H:mm:ss","d/M/yyyy H:mm:ss")
                foreach($f in $fmts){
                    if([DateTime]::TryParseExact($ts,$f,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$parsed)){
                        $tb2=New-Object System.Windows.Controls.TextBlock; $tb2.Text=$parsed.ToString("HH:mm"); $tb2.FontSize=8
                        $tb2.Foreground=$script:CBrTimeLabel
                        [System.Windows.Controls.Canvas]::SetLeft($tb2,$m+($idx*$step)-12)
                        [System.Windows.Controls.Canvas]::SetTop($tb2,$m+$dh+10)
                        $canvasGraph.Children.Add($tb2)|Out-Null; break
                    }
                }
            } catch {}
        }
    }
}

# ======================== EKSPORT CSV ========================
function Export-HistoryCSV {
    $data = Load-HistoryData 3650 0
    if ($data.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Brak danych do eksportu.", "Export CSV") | Out-Null; return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV files (*.csv)|*.csv"
    $dlg.FileName = "glucose_$(Get-Date -Format 'yyyyMMdd').csv"
    $dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add("Data,Czas,mg/dL,mmol/L,Trend")
        foreach ($pt in $data) {
            try {
                $ts = [DateTime]::Parse($pt.ts)
                $mg = [Math]::Round([double]$pt.mgdl, 1)
                $mm = [Math]::Round([double]$pt.mgdl / 18.018, 1)
                $lines.Add("$($ts.ToString('yyyy-MM-dd')),$($ts.ToString('HH:mm')),$mg,$mm,$($pt.trend)")
            } catch {}
        }
        try {
            [System.IO.File]::WriteAllLines($dlg.FileName, $lines, [System.Text.Encoding]::UTF8)
            [System.Windows.MessageBox]::Show("Wyeksportowano $($data.Count) wpisow do:`n$($dlg.FileName)", "Export CSV") | Out-Null
        } catch { [System.Windows.MessageBox]::Show("Blad zapisu: $($_.Exception.Message)", "Export CSV") | Out-Null }
    }
}

# ======================== ALERTY GLUKOZY ========================
function Check-GlucoseAlerts([double]$mmol) {
    $loAlert = $script:Config.AlertLow; $hiAlert = $script:Config.AlertHigh
    $now = Get-Date
    if ($script:LastAlertTime -and ($now - $script:LastAlertTime).TotalMinutes -lt 15) { return }
    if ($mmol -lt $loAlert) {
        $msg = if ($script:LangEn) { "Glucose: $($mmol.ToString('0.0')) mmol/L - below normal!" } else { "Glukoza: $($mmol.ToString('0.0')) mmol/L - ponizej normy!" }
        $ttl = if ($script:LangEn) { "LOW GLUCOSE" } else { "NISKI POZIOM GLUKOZY" }
        try { $script:NotifyIcon.ShowBalloonTip(8000, $ttl, $msg, [System.Windows.Forms.ToolTipIcon]::Warning) } catch {}
        try { [System.Media.SystemSounds]::Hand.Play() } catch {}
        $script:LastAlertTime = $now
    } elseif ($mmol -gt $hiAlert) {
        $msg = if ($script:LangEn) { "Glucose: $($mmol.ToString('0.0')) mmol/L - above normal!" } else { "Glukoza: $($mmol.ToString('0.0')) mmol/L - powyzej normy!" }
        $ttl = if ($script:LangEn) { "HIGH GLUCOSE" } else { "WYSOKI POZIOM GLUKOZY" }
        try { $script:NotifyIcon.ShowBalloonTip(8000, $ttl, $msg, [System.Windows.Forms.ToolTipIcon]::Warning) } catch {}
        try { [System.Media.SystemSounds]::Exclamation.Play() } catch {}
        $script:LastAlertTime = $now
    }
}

# ======================== RENDER UI (z cache) ========================
function Render-GlucoseUI {
    if ($null -eq $script:CachedMgDl) { return }
    $mgdl  = $script:CachedMgDl
    $mmol  = MgToMmol $mgdl
    Check-GlucoseAlerts $mmol
    $t     = $script:CachedTrend
    $color = Get-GlucoseColor $mmol

    if ($script:UseMgDl) {
        $displayVal = [Math]::Round($mgdl, 0).ToString("0")
        $txtUnitLabel.Text = "mg/dL"
        $btnUnit.Content   = "mg/dL"
        $btnUnit.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#ff9900"))
    } else {
        $displayVal = $mmol.ToString("0.0")
        $txtUnitLabel.Text = "mmol/L"
        $btnUnit.Content   = "mmol/L"
        $btnUnit.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#6677aa"))
    }

    $brush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($color))
    $txtGlucoseValue.Text       = $displayVal
    $txtGlucoseValue.Foreground = $brush
    $txtTrendArrow.Text         = Get-TrendArrow $t
    $txtTrendArrow.Foreground   = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString((Get-TrendColor $t)))
    $txtGlucoseStatus.Text      = Get-GlucoseStatus $mmol
    $txtTrendText.Text          = Get-TrendText $t

    # Prognoza za 30 minut (ważona regresja liniowa -> ekstrapolacja)
    if ($txtForecast) {
        $rate = Get-GlucoseRateMgMin
        if ($null -ne $rate) {
            $predMgDl   = [Math]::Max(20, [Math]::Min(600, $mgdl + $rate * 15))
            $predMmol   = $predMgDl / 18.018
            $predDisplay = if ($script:UseMgDl) { [Math]::Round($predMgDl,0).ToString("0") } `
                           else { [Math]::Round($predMmol,1).ToString("0.0") }
            $unit15     = if ($script:UseMgDl) { "mg/dL" } else { "mmol" }
            # Strzalka na podstawie roznicy prognoza-teraz
            # prog 0.2 mmol w 15 min -> wyraznie w gore/dol
            $diff15  = $predMmol - $mmol
            $arrow15 = if ($diff15 -gt 0.2) { [char]0x2197 } elseif ($diff15 -lt -0.2) { [char]0x2198 } else { [char]0x2192 }
            $col15      = if    ($predMmol -lt 3.5)  { "#FF4444" } `
                          elseif($predMmol -lt 3.9)  { "#FF8844" } `
                          elseif($predMmol -gt 13.9) { "#FF4444" } `
                          elseif($predMmol -gt 10.0) { "#FF8844" } `
                          else                        { "#55CC88" }
            $txtForecast.Text = "$arrow15 ~$predDisplay $unit15 za 15'"
            $txtForecast.Foreground = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString($col15))
        } else { $txtForecast.Text = "" }
    }

    if ($script:CachedGraphData -and $script:CachedGraphData.Count -gt 0) {
        $gv = @()
        foreach ($i in $script:CachedGraphData) {
            $mg = if($i.ValueInMgPerDl){[double]$i.ValueInMgPerDl}elseif($i.Value){[double]$i.Value}else{0}
            if ($mg -gt 0) {
                if ($script:UseMgDl) { $gv += [Math]::Round($mg, 0) }
                else { $gv += MgToMmol $mg }
            }
        }
        if ($gv.Count -gt 0) {
            $s   = $gv | Measure-Object -Min -Max -Average
            $fmt = if ($script:UseMgDl) { "0" } else { "0.0" }
            $txtMin.Text = [Math]::Round($s.Minimum, 1).ToString($fmt)
            $txtAvg.Text = [Math]::Round($s.Average, 1).ToString($fmt)
            $txtMax.Text = [Math]::Round($s.Maximum, 1).ToString($fmt)
            # Kolor wg zakresu – tak jak na wykresie (Get-GlucoseColor oczekuje mmol/L)
            $minMmol = if ($script:UseMgDl) { $s.Minimum / 18.018 } else { $s.Minimum }
            $avgMmol = if ($script:UseMgDl) { $s.Average / 18.018 } else { $s.Average }
            $maxMmol = if ($script:UseMgDl) { $s.Maximum / 18.018 } else { $s.Maximum }
            $txtMin.Foreground = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString((Get-GlucoseColor $minMmol)))
            $txtAvg.Foreground = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString((Get-GlucoseColor $avgMmol)))
            $txtMax.Foreground = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString((Get-GlucoseColor $maxMmol)))
        }
        # eHbA1c z 90 dni historii (formula NGSP)
        if ($txtHbA1c) {
            try {
                $hist90 = Load-HistoryData 90 0
                if ($hist90.Count -gt 5) {
                    $avgMg90 = ($hist90 | ForEach-Object { [double]$_.mgdl } | Measure-Object -Average).Average
                    $eH = [Math]::Round(($avgMg90 + 46.7) / 28.7, 1)
                    $txtHbA1c.Text = "$($eH.ToString('0.0'))%"
                } else { $txtHbA1c.Text = "---" }
            } catch { $txtHbA1c.Text = "---" }
        }
        Update-Graph $script:CachedGraphData

        # Delta: roznica miedzy dwoma ostatnimi punktami wykresu
        if ($script:CachedGraphData.Count -ge 2) {
            $lp = $script:CachedGraphData | Select-Object -Last 2
            $mg0 = try { [double]$lp[0].ValueInMgPerDl } catch { 0 }
            $mg1 = try { [double]$lp[1].ValueInMgPerDl } catch { 0 }
            if ($mg0 -gt 20 -and $mg1 -gt 20) {
                if ($script:UseMgDl) {
                    $dv = [Math]::Round($mg1 - $mg0, 0)
                    $txtDelta.Text = if ($dv -gt 0) { "+$dv mg/dL" } elseif ($dv -lt 0) { "$dv mg/dL" } else { "0 mg/dL" }
                } else {
                    $dv = [Math]::Round(($mg1 - $mg0) / 18.018, 1)
                    $txtDelta.Text = if ($dv -gt 0) { "+$($dv.ToString('0.0'))" } elseif ($dv -lt 0) { $dv.ToString('0.0') } else { "0.0" }
                }
                $absDv = [Math]::Abs($mg1 - $mg0)
                $dColor = if ($absDv -gt 36) { "#FF4444" } elseif ($absDv -gt 18) { "#FFAA00" } else { "#7799bb" }
                $txtDelta.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($dColor))
            } else { $txtDelta.Text = "" }
        } else { $txtDelta.Text = "" }
    }

    if ($script:UseMgDl) {
        Update-TrayTooltip "$(t 'TrayGlc')$([Math]::Round($mgdl,0)) mg/dL $(Get-TrendArrow $t)"
    } else {
        Update-TrayTooltip "$(t 'TrayGlc')$($mmol.ToString('0.0')) mmol/L $(Get-TrendArrow $t)"
    }
    Update-TrayIcon -mmol $mmol -trend $t

    if ($script:IsCompact -and $script:TxtCompactGlucose) {
        $script:TxtCompactGlucose.Text       = "$displayVal $(Get-TrendArrow $t)"
        $script:TxtCompactGlucose.Foreground = $brush
    }
}

# ======================== UPDATE ========================
function Update-ReadingAge {
    if (-not $script:LastReadingTs -or -not $txtTimestamp) { return }
    $ref     = if ($script:LastFetchTime) { $script:LastFetchTime } else { $script:LastReadingTs }
    $secs    = [int][Math]::Floor(((Get-Date) - $ref).TotalSeconds)
    $timeStr = $script:LastReadingTs.ToString("HH:mm")
    if ($secs -lt 60) {
        $ageStr = "${secs}s"
    } else {
        $m = [int][Math]::Floor($secs / 60)
        $s = $secs % 60
        $ageStr = "${m}m:$($s.ToString('00'))s"
    }
    $txtTimestamp.Text = "$timeStr  ($ageStr)"
}

function Update-Display {
    $txtStatus.Text=(t "Fetching"); $txtStatus.Foreground=[System.Windows.Media.Brushes]::Yellow
    $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render,[Action]{})

    $data = Get-GlucoseData
    if ($script:LastApiError -eq "429") {
        $txtStatus.Text = (t "TooMany")
        $txtStatus.Foreground = [System.Windows.Media.Brushes]::OrangeRed
        $txtGlucoseValue.Text = "---"
        Update-TrayTooltip (t "TrayTooM")
        $script:LastApiError = $null
    } elseif ($data -and $data.CurrentGlucose) {
        # Zapisz do cache
        $script:CachedMgDl  = [double]$data.CurrentGlucose
        $script:CachedTrend = if($data.Trend){[int]$data.Trend}else{0}
        if($data.GraphData -and $data.GraphData.Count -gt 0) {
            $script:CachedGraphData = $data.GraphData
            Save-GraphDataHistory $data.GraphData  # uzupelnij history.jsonl danymi z API
            Fill-HistoryGaps $data.GraphData       # wypelnij luki syntetycznymi punktami jesli przerwa >8h
        }

        # Trend bezpośrednio z serwera LibreLink (wartości 1-5)
        # 1=↓↓  2=↓  3=→  4=↗  5=↑↑

        # Zapisz do historii
        Save-HistoryEntry $script:CachedMgDl $script:CachedTrend

        $script:LastFetchTime = Get-Date   # moment pobrania danych przez aplikacje
        if($data.PatientName){$txtPatient.Text=$data.PatientName}
        if($data.Timestamp) {
            $ts=[string]$data.Timestamp; $p=[DateTime]::MinValue
            $fmts=@("M/d/yyyy h:mm:ss tt","d/M/yyyy h:mm:ss tt","M/d/yyyy H:mm:ss","d/M/yyyy H:mm:ss","yyyy-MM-ddTHH:mm:ss")
            foreach($f in $fmts){if([DateTime]::TryParseExact($ts,$f,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$p)){$script:LastReadingTs=$p;break}}
        }
        Update-ReadingAge

        Render-GlucoseUI

        $txtStatus.Text=(t "Connected")
        $txtStatus.Foreground=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(100,180,100))
    } else {
        $txtStatus.Text=(t "NoData"); $txtStatus.Foreground=[System.Windows.Media.Brushes]::OrangeRed; $txtGlucoseValue.Text="---"
        Update-TrayTooltip (t "TrayNoDat")
    }
    $txtNextUpdate.Text = "$(t 'NextUpdate') $($script:SecondsLeft)s"

    # Odswież okno historii po każdym pobraniu danych (np. po starcie komputera backfill moze byc nowy)
    if ($script:HistWin -and $script:HistWin.IsLoaded) {
        $script:HistCachedData = $null   # uniewaznij cache - wymus przeladowanie z pliku
        Render-HistGraph $script:HistDays
    }
}

# ======================== CACHED BRUSHES & PENS (performance) ========================
# Frozen brushes/pens sa ~3x szybsze niz dynamiczne - WPF nie musi ich obserwowac
function New-FrozenBrush([string]$hex) {
    $b = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($hex))
    $b.Freeze(); return $b
}
function New-FrozenPen([System.Windows.Media.Brush]$brush, [double]$thickness) {
    $p = New-Object System.Windows.Media.Pen($brush, $thickness)
    $p.Freeze(); return $p
}

# Kolory etykiet i siatki
$script:CBrLabel      = New-FrozenBrush "#7777aa"
$script:CBrLabel2     = New-FrozenBrush "#8888bb"
$script:CBrGrid       = New-FrozenBrush "#22FFFFFF"
$script:CBrGridFaint  = New-FrozenBrush "#15FFFFFF"
$script:CBrGrid33     = New-FrozenBrush "#33ffffff"
$script:CBrNormLine   = New-FrozenBrush "#8877AA44"
$script:CBrNormZone   = New-FrozenBrush "#3044dd44"
$script:CBrNormZone2  = New-FrozenBrush "#2200CC44"
$script:CBrNormZone3  = New-FrozenBrush "#1500CC44"
# Kolory danych - timeline
$script:CBrLineNorm   = New-FrozenBrush "#3366cc"
$script:CBrLineHyper  = New-FrozenBrush "#dd7700"
$script:CBrLineHypo   = New-FrozenBrush "#cc1111"
# Kolory danych - ogolne
$script:CBrRed        = New-FrozenBrush "#EE4444"
$script:CBrOrange     = New-FrozenBrush "#FFAA44"
$script:CBrGreen      = New-FrozenBrush "#44DDAA"
$script:CBrGreen2     = New-FrozenBrush "#6CBF26"
$script:CBrCC4444     = New-FrozenBrush "#CC4444"
# Pasma AGP
$script:CBrBandOuter  = New-FrozenBrush "#506688AA"
$script:CBrBandInner  = New-FrozenBrush "#906699CC"
# Strefy Update-Graph
$script:CBrZoneRed    = New-FrozenBrush "#44FF3333"
$script:CBrZoneOrange = New-FrozenBrush "#33FFAA00"
$script:CBrZoneGreen  = New-FrozenBrush "#2200CC44"
$script:CBrZoneYellow = New-FrozenBrush "#44FFEE00"
# AGP okno
$script:CBrAgpNorm  = New-FrozenBrush "#8844cc44"
$script:CBrAgpHypo  = New-FrozenBrush "#884444ff"
$script:CBrAgpHyper = New-FrozenBrush "#88ff8800"
# Rozne
$script:CBrWhite      = [System.Windows.Media.Brushes]::White
$script:CBrGridLabel  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(180,180,180,180))
$script:CBrGridLabel.Freeze()
$script:CBrGridLine60 = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(60,180,180,180))
$script:CBrGridLine60.Freeze()
$script:CBrScanFill   = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(210,255,255,255))
$script:CBrScanFill.Freeze()
$script:CBrScanStroke = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(120,255,255,255))
$script:CBrScanStroke.Freeze()
$script:CBrTimeLabel  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(120,150,200,200))
$script:CBrTimeLabel.Freeze()
# Pre-computed Pens (dla DrawingVisual)
$script:CPenGrid      = New-FrozenPen $script:CBrGrid 0.7
$script:CPenGridFaint = New-FrozenPen $script:CBrGridFaint 0.5
$script:CPenGrid33    = New-FrozenPen $script:CBrGrid33 0.5
$script:CPenNormLine  = New-FrozenPen $script:CBrNormLine 1.0
$script:CPenLineNorm  = New-FrozenPen $script:CBrLineNorm 1.8
$script:CPenLineHyper = New-FrozenPen $script:CBrLineHyper 1.8
$script:CPenLineHypo  = New-FrozenPen $script:CBrLineHypo 1.8
$script:CPenMedian    = New-FrozenPen $script:CBrWhite 2.5
$script:CPenGridLine60 = New-FrozenPen $script:CBrGridLine60 0.7
# FontFamily cache
$script:CFontUI       = New-Object System.Windows.Media.FontFamily("Segoe UI")
$script:CFontUIBold   = New-Object System.Windows.Media.FontFamily("Segoe UI Semibold")
$script:CTypefaceUI   = New-Object System.Windows.Media.Typeface($script:CFontUI)
$script:CTypefaceUIB  = New-Object System.Windows.Media.Typeface($script:CFontUIBold)
$script:CDpiScale     = try {
    $src = [System.Windows.PresentationSource]::FromVisual([System.Windows.Application]::Current.MainWindow)
    if ($src) { $src.CompositionTarget.TransformToDevice.M11 } else { 1.0 }
} catch { 1.0 }
if ($script:CDpiScale -le 0) { $script:CDpiScale = 1.0 }

# Fallback: dodaj siatke i etykiety bezposrednio do Canvas (gdy VisualHost niedostepny)
function Add-GridToCanvas {
    param($cv, [double]$padL, [double]$padT, [double]$gW, [double]$gH, [double]$loY, [double]$rangeY,
          [double[]]$gridVals, [string]$fmt, [double]$loNorm, [double]$hiNorm)
    foreach ($gVal in $gridVals) {
        if ($gVal -lt $loY -or $gVal -gt ($loY + $rangeY)) { continue }
        $gy = $padT + $gH - ($gVal - $loY) / $rangeY * $gH
        $gl = New-Object System.Windows.Shapes.Line; $gl.X1=$padL; $gl.X2=$padL+$gW; $gl.Y1=$gy; $gl.Y2=$gy
        $gl.Stroke = $script:CBrGrid; $gl.StrokeThickness = 0.7
        $cv.Children.Add($gl)|Out-Null
        $lbl = New-Object System.Windows.Controls.TextBlock; $lbl.Text=$gVal.ToString($fmt); $lbl.FontSize=8
        $lbl.Foreground=$script:CBrLabel; $lbl.FontFamily=$script:CFontUI
        [System.Windows.Controls.Canvas]::SetLeft($lbl,1); [System.Windows.Controls.Canvas]::SetTop($lbl,$gy-7)
        $cv.Children.Add($lbl)|Out-Null
    }
    foreach ($tgt in @($loNorm,$hiNorm)) {
        if ($tgt -lt $loY -or $tgt -gt ($loY + $rangeY)) { continue }
        $ty = $padT + $gH - ($tgt - $loY) / $rangeY * $gH
        $tl = New-Object System.Windows.Shapes.Line; $tl.X1=$padL; $tl.X2=$padL+$gW; $tl.Y1=$ty; $tl.Y2=$ty
        $tl.Stroke=$script:CBrNormLine; $tl.StrokeThickness=1.0
        $cv.Children.Add($tl)|Out-Null
    }
}

# Fallback: dodaj etykiety osi X bezposrednio do Canvas
function Add-XLabelsToCanvas {
    param($cv, [double]$yPos, [System.Collections.ArrayList]$labels)
    foreach ($lbl in $labels) {
        $tb = New-Object System.Windows.Controls.TextBlock; $tb.Text=$lbl.text; $tb.FontSize=8
        $tb.Foreground=$script:CBrLabel; $tb.FontFamily=$script:CFontUI
        [System.Windows.Controls.Canvas]::SetLeft($tb,$lbl.x - 10); [System.Windows.Controls.Canvas]::SetTop($tb,$yPos)
        $cv.Children.Add($tb)|Out-Null
    }
}

# Helper: dodaj DrawingVisual do Canvas przez VisualHost (lub fallback)
function Add-VisualToCanvas {
    param($cv, [System.Windows.Media.DrawingVisual]$dv)
    if ($script:HasVisualHost) {
        $vh = New-Object VisualHost
        $vh.AddVisual($dv)
        $cv.Children.Add($vh)|Out-Null
    }
}

# DrawingVisual helper - rysuje siatke i etykiety jednym obiektem zamiast dziesiatek Line+TextBlock
function New-GridVisual {
    param(
        [double]$padL, [double]$padT, [double]$gW, [double]$gH, [double]$padB, [double]$cH,
        [double]$loY, [double]$rangeY, [double[]]$gridVals, [string]$fmt,
        [double]$loNorm, [double]$hiNorm,
        [System.Windows.Media.Pen]$gridPen,
        [System.Windows.Media.Pen]$normPen,
        [System.Windows.Media.Brush]$labelBrush
    )
    $dv = New-Object System.Windows.Media.DrawingVisual
    $dc = $dv.RenderOpen()
    $dpi = $script:CDpiScale
    # Siatka Y + etykiety
    foreach ($gVal in $gridVals) {
        if ($gVal -lt $loY -or $gVal -gt ($loY + $rangeY)) { continue }
        $gy = $padT + $gH - ($gVal - $loY) / $rangeY * $gH
        $dc.DrawLine($gridPen, [System.Windows.Point]::new($padL, $gy), [System.Windows.Point]::new($padL + $gW, $gy))
        $ft = New-Object System.Windows.Media.FormattedText($gVal.ToString($fmt), [System.Globalization.CultureInfo]::InvariantCulture, [System.Windows.FlowDirection]::LeftToRight, $script:CTypefaceUI, 8, $labelBrush, $dpi)
        $dc.DrawText($ft, [System.Windows.Point]::new(1, $gy - 7))
    }
    # Linie targetu normy
    foreach ($tgt in @($loNorm, $hiNorm)) {
        if ($tgt -lt $loY -or $tgt -gt ($loY + $rangeY)) { continue }
        $ty = $padT + $gH - ($tgt - $loY) / $rangeY * $gH
        $dc.DrawLine($normPen, [System.Windows.Point]::new($padL, $ty), [System.Windows.Point]::new($padL + $gW, $ty))
    }
    $dc.Close()
    return $dv
}

# DrawingVisual helper - etykiety osi X na Canvas
function New-XLabelsVisual {
    param(
        [System.Collections.ArrayList]$labels,  # @{x=[double]; text=[string]}
        [double]$yPos,
        [System.Windows.Media.Brush]$brush
    )
    $dv = New-Object System.Windows.Media.DrawingVisual
    $dc = $dv.RenderOpen()
    $dpi = $script:CDpiScale
    foreach ($lbl in $labels) {
        $ft = New-Object System.Windows.Media.FormattedText($lbl.text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Windows.FlowDirection]::LeftToRight, $script:CTypefaceUI, 8, $brush, $dpi)
        $dc.DrawText($ft, [System.Windows.Point]::new($lbl.x - 10, $yPos))
    }
    $dc.Close()
    return $dv
}

# DrawingVisual-aware Canvas host - potrzebny zeby Canvas mog wyswietlac DrawingVisual
# Standardowy Canvas nie obsluguje Visual - potrzebujemy klasy pochodnej
# Uzywamy Add-Type z C# aby stworzyc VisualHost (element UIElement ktory renderuje DrawingVisual wewnatrz Canvas)
# PS 5.1 (.NET Framework): krotkie nazwy assembly (GAC)
# PS 7+ (.NET Core+): pelne sciezki z zaladowanych typow
$script:HasVisualHost = $false
$script:_vhCSharp = @'
using System.Windows.Media;
using System.Windows;

public class VisualHost : FrameworkElement {
    private VisualCollection _children;
    public VisualHost() { _children = new VisualCollection(this); }
    public void AddVisual(DrawingVisual v) { _children.Add(v); }
    public void ClearVisuals() { _children.Clear(); }
    protected override int VisualChildrenCount { get { return _children.Count; } }
    protected override Visual GetVisualChild(int index) { return _children[index]; }
}
'@
try {
    # Proba 1: krotkie nazwy assembly (PS 5.1 / .NET Framework GAC) + System.Xaml wymagany przez FrameworkElement
    Add-Type -ReferencedAssemblies @('PresentationCore','PresentationFramework','WindowsBase','System.Xaml') -TypeDefinition $script:_vhCSharp -ErrorAction Stop
    $script:HasVisualHost = $true
} catch {
    try {
        # Proba 2: pelne sciezki z zaladowanych typow (PS 7+ / .NET Core)
        $script:_vhRefs = @(
            [System.Windows.Media.Visual].Assembly.Location,
            [System.Windows.FrameworkElement].Assembly.Location,
            [System.Windows.DependencyObject].Assembly.Location
        ) | Where-Object { $_ } | Select-Object -Unique
        if ($script:_vhRefs.Count -gt 0) {
            Add-Type -ReferencedAssemblies $script:_vhRefs -TypeDefinition $script:_vhCSharp -ErrorAction Stop
            $script:HasVisualHost = $true
        }
    } catch {
        Write-Log "VisualHost Add-Type failed: $($_.Exception.Message) - using fallback rendering"
    }
}

# StreamGeometry helper - buduje Catmull-Rom -> Bezier z tablic px/py
function New-CatmullRomStreamGeo {
    param([double[]]$px, [double[]]$py, [int]$iStart, [int]$iEnd, [double]$tension)
    $sg = New-Object System.Windows.Media.StreamGeometry
    $sg.FillRule = [System.Windows.Media.FillRule]::Nonzero
    $ctx = $sg.Open()
    $ctx.BeginFigure([System.Windows.Point]::new($px[$iStart], $py[$iStart]), $false, $false)
    for ($i = $iStart; $i -lt $iEnd; $i++) {
        $i0 = if ($i -gt $iStart) { $i - 1 } else { $iStart }
        $i3 = if ($i + 2 -le $iEnd) { $i + 2 } else { $iEnd }
        $cp1 = [System.Windows.Point]::new($px[$i] + ($px[$i+1] - $px[$i0]) * $tension, $py[$i] + ($py[$i+1] - $py[$i0]) * $tension)
        $cp2 = [System.Windows.Point]::new($px[$i+1] - ($px[$i3] - $px[$i]) * $tension, $py[$i+1] - ($py[$i3] - $py[$i]) * $tension)
        $ep  = [System.Windows.Point]::new($px[$i+1], $py[$i+1])
        $ctx.BezierTo($cp1, $cp2, $ep, $true, $false)
    }
    $ctx.Close()
    $sg.Freeze()
    return $sg
}

# LTTB downsampling - redukuje liczbe punktow zachowujac ksztalt wykresu
function Invoke-LTTB {
    param([double[]]$xArr, [double[]]$yArr, [int]$targetN)
    $n = $xArr.Length
    if ($n -le $targetN) { return @{ x = $xArr; y = $yArr } }
    $outX = [System.Collections.Generic.List[double]]::new($targetN)
    $outY = [System.Collections.Generic.List[double]]::new($targetN)
    # Zawsze zachowaj pierwszy punkt
    $outX.Add($xArr[0]); $outY.Add($yArr[0])
    $bucketSize = [double]($n - 2) / ($targetN - 2)
    $aIdx = 0
    for ($i = 0; $i -lt $targetN - 2; $i++) {
        # Srednia nastepnego bucket (look-ahead)
        $bStart = [int][Math]::Floor(($i + 1) * $bucketSize) + 1
        $bEnd   = [int][Math]::Floor(($i + 2) * $bucketSize) + 1
        if ($bEnd -ge $n) { $bEnd = $n - 1 }
        $avgX = 0.0; $avgY = 0.0; $cnt = 0
        for ($j = $bStart; $j -le $bEnd; $j++) { $avgX += $xArr[$j]; $avgY += $yArr[$j]; $cnt++ }
        if ($cnt -gt 0) { $avgX /= $cnt; $avgY /= $cnt }
        # Obecny bucket
        $cStart = [int][Math]::Floor($i * $bucketSize) + 1
        $cEnd   = [int][Math]::Floor(($i + 1) * $bucketSize)
        if ($cEnd -ge $n) { $cEnd = $n - 1 }
        # Znajdz punkt w bucket o max trojkacie z prev point i avg next
        $maxArea = -1.0; $bestIdx = $cStart
        for ($j = $cStart; $j -le $cEnd; $j++) {
            $area = [Math]::Abs(($xArr[$aIdx] - $avgX) * ($yArr[$j] - $yArr[$aIdx]) - ($xArr[$aIdx] - $xArr[$j]) * ($avgY - $yArr[$aIdx]))
            if ($area -gt $maxArea) { $maxArea = $area; $bestIdx = $j }
        }
        $outX.Add($xArr[$bestIdx]); $outY.Add($yArr[$bestIdx])
        $aIdx = $bestIdx
    }
    # Zawsze zachowaj ostatni punkt
    $outX.Add($xArr[$n-1]); $outY.Add($yArr[$n-1])
    return @{ x = $outX.ToArray(); y = $outY.ToArray() }
}

# ======================== HISTORIA ========================
$script:HistWin        = $null
$script:HistCanvas     = $null
$script:AvgBarsWin     = $null
$script:AvgBarsAvgs    = $null   # [double[]] pre-computed slot averages
$script:AvgBarsHasData = $null   # [bool[]]   slot has data flag
$script:AvgBarsLoY     = 0.0
$script:AvgBarsHiY     = 1.0
$script:AvgBarsRngY    = 1.0
$script:AvgBarsLoN     = 3.9
$script:AvgBarsHiN     = 10.0
$script:AvgBarsHiC     = 13.9
$script:AvgBarsUseMgDl = $false
$script:HistDays       = 7
$script:HistValAvg     = $null; $script:HistValMin  = $null
$script:HistValMax     = $null; $script:HistValTIR  = $null
$script:HistValDelta   = $null; $script:HistNoData  = $null
$script:HistBtns       = $null   # array 4 przyciskow okresu
$script:HistLblTitleTxt = $null; $script:HistLblAvg = $null; $script:HistLblTIR = $null
$script:HistViewMode   = "agp"   # "agp" = wzorzec dobowy, "timeline" = odczyty linia czasu
$script:HistBtnLine    = $null
$script:HistZoomFactor = 1.0     # zoom wykresu LINE (1.0 = 100%, 2.0 = 200% itp.)
$script:HistPanOffset  = 0.0     # przesuniecie poziome wykresu LINE (w minutach)
$script:HistPanStartX  = 0.0     # pomocnicza - pozycja X myszy przy rozpoczeciu panowania
$script:HistPanStartOffset = 0.0 # pomocnicza - offset przy rozpoczeciu panowania
$script:HistZoomTimer  = $null   # timer dla throttlingu zoomu (debouncing)
$script:HistCachedData = $null   # cache danych dla szybkiego zoomu (unikanie ponownego Load-HistoryData)
$script:HistCachedDays = -1      # cache: ile dni bylo zaladowanych
$script:HistCachedOffset = -1    # cache: jaki byl offset

# --- Percentyl (interpolacja liniowa) z posortowanej tablicy ---
function Get-Percentile([double[]]$sorted, [double]$p) {
    $n = $sorted.Length
    if ($n -eq 0) { return 0.0 }
    if ($n -eq 1) { return $sorted[0] }
    [double]$idx = $p / 100.0 * ($n - 1)
    $lo = [int][Math]::Floor($idx); $hi = [int][Math]::Ceiling($idx)
    if ($lo -eq $hi) { return $sorted[$lo] }
    return $sorted[$lo] + ($idx - $lo) * ($sorted[$hi] - $sorted[$lo])
}

# --- Wygładzanie tablicy percentyli - ruchoma srednia okrężna (24h jest cykliczna) ---
function Smooth-Array([double[]]$arr, [bool[]]$valid, [int]$w) {
    $out = [double[]]::new(288)
    for ($s = 0; $s -lt 288; $s++) {
        if (-not $valid[$s]) { continue }
        $sum = 0.0; $cnt = 0
        for ($j = -$w; $j -le $w; $j++) {
            $idx2 = ($s + $j + 288) % 288
            if ($valid[$idx2]) { $sum += $arr[$idx2]; $cnt++ }
        }
        $out[$s] = if ($cnt -gt 0) { $sum / $cnt } else { $arr[$s] }
    }
    return $out
}

# ---- Render-HistGraph (wzorzec dobowy AGP - styl LibreLink) na poziomie skryptu ----
function Render-HistGraph([int]$days, [bool]$useCache = $false) {
    if (-not $script:HistCanvas) { return }
    try {
        # Podswietl aktywny przycisk okresu
        if ($script:HistBtns) {
            $dayMap = @(1,7,14,30,90)
            $brActive = New-FrozenBrush "#2a3a5a"
            $brInactive = New-FrozenBrush "#1e2a3a"
            for ($bi=0; $bi -lt 5; $bi++) {
                $b = $script:HistBtns[$bi]
                if (-not $b) { continue }
                if ($dayMap[$bi] -eq $days) {
                    $b.Background = $brActive
                    $b.Foreground = $script:CBrWhite
                } else {
                    $b.Background = $brInactive
                    $b.Foreground = $script:CBrLabel
                }
            }
        }

        # Cache dla szybkiego zoomu - unikaj ponownego wczytywania danych
        if ($useCache -and $script:HistCachedData -and $script:HistCachedDays -eq $days -and $script:HistCachedOffset -eq $script:HistOffset) {
            $data = $script:HistCachedData
        } else {
            $data = Load-HistoryData $days $script:HistOffset
            $script:HistCachedData = $data
            $script:HistCachedDays = $days
            $script:HistCachedOffset = $script:HistOffset
        }
        
        $script:HistCanvas.Children.Clear()

        # Etykieta zakresu dat
        if ($script:HistRangeLabel) {
            $endDate   = (Get-Date).AddDays(-$script:HistOffset)
            $startDate = $endDate.AddDays(-$days)
            $script:HistRangeLabel.Text = "$($startDate.ToString('dd.MM')) - $($endDate.ToString('dd.MM.yyyy'))"
        }

        # Pokaz komunikat gdy brak danych
        if ($data.Count -lt 2) {
            if ($script:HistNoData) { $script:HistNoData.Visibility = [System.Windows.Visibility]::Visible }
            if ($script:HistValAvg) { $script:HistValAvg.Text = "---"; $script:HistValMin.Text = "---"
                                      $script:HistValMax.Text = "---"; $script:HistValTIR.Text = "---"
                                      $script:HistValDelta.Text = "---" }
            return
        }
        if ($script:HistNoData) { $script:HistNoData.Visibility = [System.Windows.Visibility]::Collapsed }

        $fmt    = if ($script:UseMgDl) { "0" } else { "0.0" }
        $loNorm = if ($script:UseMgDl) { 70.0  } else { 3.9  }
        $hiNorm = if ($script:UseMgDl) { 180.0 } else { 10.0 }

        # Zbierz dane: allVals do statystyk + slots 5-min + tlPts (pre-built dla timeline)
        # Jedna petla zamiast wielu pipeline'ow - znacznie szybciej
        $allVals = [System.Collections.Generic.List[double]]::new()
        $deltas  = [System.Collections.Generic.List[double]]::new()
        $slots   = @{}   # slotIdx -> List[double]
        $tlPts   = [System.Collections.Generic.List[object]]::new()   # gotowe dane dla timeline
        $prevMg  = 0.0
        # Inline statystyki (bez pipeline Measure-Object / Where-Object)
        [double]$stSum=0; [double]$stMin=[double]::MaxValue; [double]$stMax=[double]::MinValue; [int]$stInR=0
        [double]$dSum=0
        foreach ($pt in $data) {
            [double]$mg = $pt.mgdl
            if ($mg -le 20) { continue }
            $ptTs = if ($pt.tsdt) { $pt.tsdt } else { try { [DateTime]::Parse($pt.ts) } catch { [DateTime]::MinValue } }
            if ($ptTs -eq [DateTime]::MinValue) { continue }
            $displayV = if ($script:UseMgDl) { [Math]::Round($mg,0) } else { [Math]::Round($mg/18.018,1) }
            $allVals.Add($displayV)
            $stSum += $displayV
            if ($displayV -lt $stMin) { $stMin = $displayV }
            if ($displayV -gt $stMax) { $stMax = $displayV }
            if ($displayV -ge $loNorm -and $displayV -le $hiNorm) { $stInR++ }
            if ($prevMg -gt 20) { $d = [Math]::Abs($mg - $prevMg); $deltas.Add($d); $dSum += $d }
            $prevMg = $mg
            $si = $ptTs.Hour * 12 + [int]($ptTs.Minute / 5)   # slot 5-min (0..287)
            if (-not $slots.ContainsKey($si)) { $slots[$si] = [System.Collections.Generic.List[double]]::new() }
            $slots[$si].Add($displayV)
            $tlPts.Add([PSCustomObject]@{ ts=$ptTs; v=$displayV })
        }
        if ($allVals.Count -lt 2) {
            if ($script:HistNoData) { $script:HistNoData.Visibility = [System.Windows.Visibility]::Visible }
            return
        }

        # Statystyki panelu - obliczone inline (bez pipeline)
        [double]$stAvg = $stSum / $allVals.Count
        $tirPct = [Math]::Round($stInR / $allVals.Count * 100, 0)
        if ($script:HistValAvg) {
            $script:HistValAvg.Text = [Math]::Round($stAvg,1).ToString($fmt)
            $script:HistValMin.Text = $stMin.ToString($fmt)
            $script:HistValMax.Text = $stMax.ToString($fmt)
            $script:HistValTIR.Text = "$tirPct%"
            if ($script:HistValHbA1c) {
                $avgMgDl = if ($script:UseMgDl) { $stAvg } else { $stAvg * 18.018 }
                $script:HistValHbA1c.Text = "$([Math]::Round(($avgMgDl+46.7)/28.7,1).ToString('0.0'))%"
            }
            if ($allVals.Count -gt 1) {
                $ssq = 0.0; foreach ($v in $allVals) { $diff = $v - $stAvg; $ssq += $diff * $diff }
                $sd  = [Math]::Sqrt($ssq / $allVals.Count)
                $cv  = if ($stAvg -gt 0) { [Math]::Round($sd/$stAvg*100,0) } else { 0 }
                $sdD = if ($script:UseMgDl) { [Math]::Round($sd,0).ToString("0") } else { [Math]::Round($sd,1).ToString("0.0") }
                if ($script:HistValSD) { $script:HistValSD.Text = $sdD }
                if ($script:HistValCV) { $script:HistValCV.Text = "$cv%" }
            }
            if ($deltas.Count -gt 0) {
                $adRaw = $dSum / $deltas.Count
                $adD   = if ($script:UseMgDl) { [Math]::Round($adRaw,0) } else { [Math]::Round($adRaw/18.018,1) }
                $script:HistValDelta.Text = $adD.ToString($fmt)
            } else { $script:HistValDelta.Text = "---" }
        }

        # ======== TRYB TIMELINE - dokladne odczyty glukozy na osi czasu ========
        if ($script:HistViewMode -eq "timeline") {
            if ($tlPts.Count -lt 2) { return }

            $t0 = $tlPts[0].ts; $t1 = $tlPts[$tlPts.Count-1].ts
            [double]$tRngRaw = ($t1 - $t0).TotalMinutes
            if ($tRngRaw -lt 1) { return }

            [double]$tRng = $tRngRaw / $script:HistZoomFactor
            $t0Eff = $t0.AddMinutes($script:HistPanOffset)
            $t1Eff = $t0Eff.AddMinutes($tRng)
            
            if ($script:HistWin) { $script:HistWin.UpdateLayout() }
            $cW=$script:HistCanvas.ActualWidth;  if ($cW -lt 10) { $cW=440.0 }
            $cH=$script:HistCanvas.ActualHeight; if ($cH -lt 10) { $cH=250.0 }
            $padL=38.0; $padR=8.0; $padT=8.0; $padB=22.0
            $gW=$cW-$padL-$padR; $gH=$cH-$padT-$padB

            $loY=if ($script:UseMgDl){40.0}else{2.2}
            $hiY=if ($script:UseMgDl){280.0}else{15.5}
            if ($stMin -lt $loY) { $loY=[Math]::Floor($stMin-0.5) }
            if ($stMax -gt $hiY) { $hiY=[Math]::Ceiling($stMax+0.5) }
            $rangeY=$hiY-$loY; if ($rangeY -lt 0.1) { $rangeY=1.0 }

            # Strefa normy (zielony prostokat)
            $nLoY = $padT+$gH-($loNorm-$loY)/$rangeY*$gH
            $nHiY = $padT+$gH-($hiNorm-$loY)/$rangeY*$gH
            $nTop = [Math]::Min($nLoY,$nHiY); $nHt = [Math]::Abs($nLoY-$nHiY)
            if ($nHt -gt 0) {
                $zone = New-Object System.Windows.Shapes.Rectangle
                $zone.Width = $gW; $zone.Height = $nHt
                $zone.Fill = $script:CBrNormZone
                [System.Windows.Controls.Canvas]::SetLeft($zone,$padL)
                [System.Windows.Controls.Canvas]::SetTop($zone,$nTop)
                $script:HistCanvas.Children.Add($zone)|Out-Null
            }

            # Siatka Y + etykiety + linie normy
            $gridVals=if ($script:UseMgDl){@(70,100,140,180,250)}else{@(3.9,5.5,7.0,10.0,13.9)}
            if ($script:HasVisualHost) {
                $gridDV = New-GridVisual -padL $padL -padT $padT -gW $gW -gH $gH -padB $padB -cH $cH `
                    -loY $loY -rangeY $rangeY -gridVals $gridVals -fmt $fmt `
                    -loNorm $loNorm -hiNorm $hiNorm -gridPen $script:CPenGrid -normPen $script:CPenNormLine -labelBrush $script:CBrLabel
                Add-VisualToCanvas $script:HistCanvas $gridDV
            } else {
                Add-GridToCanvas $script:HistCanvas $padL $padT $gW $gH $loY $rangeY $gridVals $fmt $loNorm $hiNorm
            }

            # Piksele per punkt
            $n = $tlPts.Count
            $pxArr=[double[]]::new($n); $pyArr=[double[]]::new($n)
            for ($i=0; $i -lt $n; $i++) {
                $pxArr[$i] = $padL + ($tlPts[$i].ts - $t0Eff).TotalMinutes / $tRng * $gW
                $pyArr[$i] = $padT + $gH - ($tlPts[$i].v - $loY) / $rangeY * $gH
            }

            # LTTB downsampling - redukuj punkty jesli jest ich duzo
            $maxPts = 500
            if ($n -gt $maxPts) {
                # Potrzebujemy zachowac tez tlPts (wartosci) do kolorowania
                $tlVals = [double[]]::new($n)
                for ($i=0; $i -lt $n; $i++) { $tlVals[$i] = $tlPts[$i].v }
                $ds = Invoke-LTTB -xArr $pxArr -yArr $pyArr -targetN $maxPts
                # Rownolegly LTTB na wartosciach (ten sam podzial bucket)
                $dsV = Invoke-LTTB -xArr $pxArr -yArr $tlVals -targetN $maxPts
                $pxArr = $ds.x; $pyArr = $ds.y; $n = $pxArr.Length
                $tlValsDS = $dsV.y
            } else {
                $tlValsDS = [double[]]::new($n)
                for ($i=0; $i -lt $n; $i++) { $tlValsDS[$i] = $tlPts[$i].v }
            }

            # Linia danych - StreamGeometry z Catmull-Rom -> Bezier, kolorowe segmenty, przerwy >20min
            [double]$maxGapMin = 20.0
            [double]$tension = 0.2
            # Przy downsamplingu przerwy sa juz wbudowane w dane (LTTB zachowuje ksztalt)
            # Segmentujemy po kolorze i przerwach
            $segNorm  = New-Object System.Collections.ArrayList
            $segHypo  = New-Object System.Collections.ArrayList
            $segHyper = New-Object System.Collections.ArrayList

            # Buduj segmenty kolorowe - grupuj kolejne punkty tego samego koloru
            # Dla kazdego segmentu ciaglosci (bez przerw) tworzymy sub-segmenty kolorowe
            $contStart = 0
            while ($contStart -lt $n) {
                $contEnd = $contStart
                # Dla downsamplowanych danych nie mamy ts wiec nie sprawdzamy przerw bezposrednio
                # Zamiast tego sprawdzamy duze skoki w px (>maxGapMin odpowiednik w pikselach)
                if ($n -le $maxPts) {
                    # Oryginalne dane - sprawdzaj przerwy czasowe
                    while ($contEnd+1 -lt $n -and ($tlPts[$contEnd+1].ts - $tlPts[$contEnd].ts).TotalMinutes -le $maxGapMin) {
                        $contEnd++
                    }
                } else {
                    # Downsampled - nie sprawdzaj przerw (LTTB utrzymuje ciaglosc)
                    $contEnd = $n - 1
                }
                if ($contEnd -gt $contStart) {
                    for ($i=$contStart; $i -lt $contEnd; $i++) {
                        $avgV = ($tlValsDS[$i] + $tlValsDS[$i+1]) / 2.0
                        if ($avgV -lt $loNorm) { $segHypo.Add(@([int]$i, [int]($i+1))) | Out-Null }
                        elseif ($avgV -gt $hiNorm) { $segHyper.Add(@([int]$i, [int]($i+1))) | Out-Null }
                        else { $segNorm.Add(@([int]$i, [int]($i+1))) | Out-Null }
                    }
                }
                $contStart = $contEnd + 1
            }

            # Rysuj segmenty jako StreamGeometry Paths
            foreach ($gd in @(
                @($segNorm,  $script:CPenLineNorm),
                @($segHyper, $script:CPenLineHyper),
                @($segHypo,  $script:CPenLineHypo)
            )) {
                $segs = $gd[0]; $pen = $gd[1]
                if ($segs.Count -eq 0) { continue }
                # Laczymy sasiedzkie segmenty w ciagi dla jednej StreamGeometry
                $sg = New-Object System.Windows.Media.StreamGeometry
                $ctx = $sg.Open()
                $prevEnd = -1
                foreach ($seg in $segs) {
                    $iA = $seg[0]; $iB = $seg[1]
                    if ($iA -ne $prevEnd) {
                        # Nowy figure
                        $ctx.BeginFigure([System.Windows.Point]::new($pxArr[$iA], $pyArr[$iA]), $false, $false)
                    }
                    $i0 = if ($iA -gt 0) { $iA-1 } else { 0 }
                    $i3 = if ($iB+1 -lt $n) { $iB+1 } else { $n-1 }
                    $cp1 = [System.Windows.Point]::new($pxArr[$iA]+($pxArr[$iB]-$pxArr[$i0])*$tension, $pyArr[$iA]+($pyArr[$iB]-$pyArr[$i0])*$tension)
                    $cp2 = [System.Windows.Point]::new($pxArr[$iB]-($pxArr[$i3]-$pxArr[$iA])*$tension, $pyArr[$iB]-($pyArr[$i3]-$pyArr[$iA])*$tension)
                    $ctx.BezierTo($cp1, $cp2, [System.Windows.Point]::new($pxArr[$iB], $pyArr[$iB]), $true, $false)
                    $prevEnd = $iB
                }
                $ctx.Close()
                $sg.Freeze()
                $p = New-Object System.Windows.Shapes.Path
                $p.Data = $sg; $p.Stroke = $pen.Brush; $p.StrokeThickness = $pen.Thickness
                $script:HistCanvas.Children.Add($p)|Out-Null
            }

            # Pionowe linie siatki X + etykiety dat/godzin
            [double]$visibleDays = $tRng / 1440.0
            # Zbierz pozycje etykiet/linii (wspolne dla obu trybow renderowania)
            $xGridItems = [System.Collections.ArrayList]::new()   # @{xp; text}

            if ($visibleDays -gt 14) {
                $dayStep = if ($visibleDays -gt 60) { 7 } elseif ($visibleDays -gt 30) { 5 } else { 2 }
                $cur = $t0Eff.Date.AddDays($dayStep)
                while ($cur -le $t1Eff) {
                    if ($cur -ge $t0Eff) {
                        $xp = $padL + ($cur - $t0Eff).TotalMinutes / $tRng * $gW
                        if ($xp -ge $padL -and $xp -le $padL+$gW) { $xGridItems.Add(@{xp=$xp;text=$cur.ToString('dd.MM')})|Out-Null }
                    }
                    $cur = $cur.AddDays($dayStep)
                }
            } elseif ($visibleDays -gt 2) {
                $cur = $t0Eff.Date.AddDays(1)
                while ($cur -le $t1Eff) {
                    if ($cur -ge $t0Eff) {
                        $xp = $padL + ($cur - $t0Eff).TotalMinutes / $tRng * $gW
                        if ($xp -ge $padL -and $xp -le $padL+$gW) { $xGridItems.Add(@{xp=$xp;text=$cur.ToString('dd.MM')})|Out-Null }
                    }
                    $cur = $cur.AddDays(1)
                }
            } elseif ($visibleDays -gt 0.5) {
                $cur = $t0Eff.Date.AddHours([int]($t0Eff.Hour / 4) * 4)
                while ($cur -le $t1Eff) {
                    if ($cur -ge $t0Eff) {
                        $xp = $padL + ($cur - $t0Eff).TotalMinutes / $tRng * $gW
                        if ($xp -ge $padL -and $xp -le $padL+$gW) { $xGridItems.Add(@{xp=$xp;text=$cur.ToString('HH:mm')})|Out-Null }
                    }
                    $cur = $cur.AddHours(4)
                }
            } else {
                $cur = $t0Eff.Date.AddHours([int]$t0Eff.Hour)
                while ($cur -le $t1Eff) {
                    if ($cur -ge $t0Eff) {
                        $xp = $padL + ($cur - $t0Eff).TotalMinutes / $tRng * $gW
                        if ($xp -ge $padL -and $xp -le $padL+$gW) { $xGridItems.Add(@{xp=$xp;text=$cur.ToString('HH:mm')})|Out-Null }
                    }
                    $cur = $cur.AddHours(1)
                }
            }

            # Renderuj zebrane pozycje
            if ($script:HasVisualHost) {
                $dvX = New-Object System.Windows.Media.DrawingVisual
                $dcX = $dvX.RenderOpen()
                $dpi = $script:CDpiScale
                foreach ($gi in $xGridItems) {
                    $dcX.DrawLine($script:CPenGridFaint, [System.Windows.Point]::new($gi.xp,$padT), [System.Windows.Point]::new($gi.xp,$padT+$gH))
                    $ft = New-Object System.Windows.Media.FormattedText($gi.text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Windows.FlowDirection]::LeftToRight, $script:CTypefaceUI, 8, $script:CBrLabel, $dpi)
                    $dcX.DrawText($ft, [System.Windows.Point]::new($gi.xp-10, $cH-$padB+4))
                }
                $dcX.Close()
                Add-VisualToCanvas $script:HistCanvas $dvX
            } else {
                foreach ($gi in $xGridItems) {
                    $vl=New-Object System.Windows.Shapes.Line; $vl.X1=$gi.xp; $vl.X2=$gi.xp; $vl.Y1=$padT; $vl.Y2=$padT+$gH
                    $vl.Stroke=$script:CBrGridFaint; $vl.StrokeThickness=0.5
                    $script:HistCanvas.Children.Add($vl)|Out-Null
                    $dl=New-Object System.Windows.Controls.TextBlock; $dl.Text=$gi.text; $dl.FontSize=8
                    $dl.Foreground=$script:CBrLabel; $dl.FontFamily=$script:CFontUI
                    [System.Windows.Controls.Canvas]::SetLeft($dl,$gi.xp-10); [System.Windows.Controls.Canvas]::SetTop($dl,$cH-$padB+4)
                    $script:HistCanvas.Children.Add($dl)|Out-Null
                }
            }
            return
        }
        # ======== KONIEC TRYBU TIMELINE ========

        # Percentyle dla 288 slotow 5-min (calowa doba)
        $sP10=[double[]]::new(288); $sP25=[double[]]::new(288); $sP50=[double[]]::new(288)
        $sP75=[double[]]::new(288); $sP90=[double[]]::new(288); $hasS=[bool[]]::new(288)
        for ($s=0; $s -lt 288; $s++) {
            if ($slots.ContainsKey($s) -and $slots[$s].Count -ge 1) {
                $sv = $slots[$s].ToArray(); [Array]::Sort($sv)
                $sP10[$s]=Get-Percentile $sv 10; $sP25[$s]=Get-Percentile $sv 25
                $sP50[$s]=Get-Percentile $sv 50; $sP75[$s]=Get-Percentile $sv 75
                $sP90[$s]=Get-Percentile $sv 90; $hasS[$s]=$true
            }
        }

        # Wygladz krzywe percentyli (okno zalezne od liczby dni)
        $smW = if ($days -le 7) { 30 } elseif ($days -le 14) { 15 } elseif ($days -le 30) { 8 } else { 4 }
        $sP10 = Smooth-Array $sP10 $hasS $smW
        $sP25 = Smooth-Array $sP25 $hasS $smW
        $sP50 = Smooth-Array $sP50 $hasS $smW
        $sP75 = Smooth-Array $sP75 $hasS $smW
        $sP90 = Smooth-Array $sP90 $hasS $smW

        # Wymiary canvas
        if ($script:HistWin) { $script:HistWin.UpdateLayout() }
        $cW=$script:HistCanvas.ActualWidth;  if ($cW -lt 10) { $cW=440.0 }
        $cH=$script:HistCanvas.ActualHeight; if ($cH -lt 10) { $cH=250.0 }
        $padL=38.0; $padR=8.0; $padT=8.0; $padB=22.0
        $gW=$cW-$padL-$padR; $gH=$cH-$padT-$padB

        $loY=if ($script:UseMgDl){40.0}else{2.2}
        $hiY=if ($script:UseMgDl){280.0}else{15.5}
        if ($stMin -lt $loY) { $loY=[Math]::Floor($stMin-0.5) }
        if ($stMax -gt $hiY) { $hiY=[Math]::Ceiling($stMax+0.5) }
        $rangeY=$hiY-$loY; if ($rangeY -lt 0.1) { $rangeY=1.0 }

        # Siatka Y + etykiety + linie normy
        $gridVals=if ($script:UseMgDl){@(70,100,140,180,250)}else{@(3.9,5.5,7.0,10.0,13.9)}
        if ($script:HasVisualHost) {
            $gridDV = New-GridVisual -padL $padL -padT $padT -gW $gW -gH $gH -padB $padB -cH $cH `
                -loY $loY -rangeY $rangeY -gridVals $gridVals -fmt $fmt `
                -loNorm $loNorm -hiNorm $hiNorm -gridPen $script:CPenGrid -normPen $script:CPenNormLine -labelBrush $script:CBrLabel
            Add-VisualToCanvas $script:HistCanvas $gridDV
        } else {
            Add-GridToCanvas $script:HistCanvas $padL $padT $gW $gH $loY $rangeY $gridVals $fmt $loNorm $hiNorm
        }

        # Wspolrzedne pikseli per slot (X rozlozony na pelna dobe 00:00-23:55)
        $xA=[double[]]::new(288)
        $y10=[double[]]::new(288); $y25=[double[]]::new(288); $y50=[double[]]::new(288)
        $y75=[double[]]::new(288); $y90=[double[]]::new(288)
        for ($s=0; $s -lt 288; $s++) {
            $xA[$s]=$padL+$s/287.0*$gW
            if ($hasS[$s]) {
                $y90[$s]=$padT+$gH-($sP90[$s]-$loY)/$rangeY*$gH
                $y75[$s]=$padT+$gH-($sP75[$s]-$loY)/$rangeY*$gH
                $y50[$s]=$padT+$gH-($sP50[$s]-$loY)/$rangeY*$gH
                $y25[$s]=$padT+$gH-($sP25[$s]-$loY)/$rangeY*$gH
                $y10[$s]=$padT+$gH-($sP10[$s]-$loY)/$rangeY*$gH
            }
        }

        # Pasmo p10-p90 (zewnetrzne, jasniejsze) - obejmuje 80% pomiarow
        $pts90=[System.Windows.Media.PointCollection]::new()
        $pts10=[System.Collections.Generic.List[System.Windows.Point]]::new()
        for ($s=0; $s -lt 288; $s+=3) {
            if ($hasS[$s]) {
                $pts90.Add([System.Windows.Point]::new($xA[$s],$y90[$s]))
                $pts10.Add([System.Windows.Point]::new($xA[$s],$y10[$s]))
            }
        }
        if ($hasS[287]) {
            $pts90.Add([System.Windows.Point]::new($xA[287],$y90[287]))
            $pts10.Add([System.Windows.Point]::new($xA[287],$y10[287]))
        }
        if ($pts90.Count -ge 3) {
            for ($i=$pts10.Count-1; $i -ge 0; $i--) { $pts90.Add($pts10[$i]) }
            $po=New-Object System.Windows.Shapes.Polygon; $po.Points=$pts90
            $po.Fill=$script:CBrBandOuter
            $script:HistCanvas.Children.Add($po)|Out-Null
        }

        # Pasmo p25-p75 IQR (wewnetrzne, ciemniejsze) - obejmuje 50% pomiarow
        $pts75=[System.Windows.Media.PointCollection]::new()
        $pts25=[System.Collections.Generic.List[System.Windows.Point]]::new()
        for ($s=0; $s -lt 288; $s+=3) {
            if ($hasS[$s]) {
                $pts75.Add([System.Windows.Point]::new($xA[$s],$y75[$s]))
                $pts25.Add([System.Windows.Point]::new($xA[$s],$y25[$s]))
            }
        }
        if ($hasS[287]) {
            $pts75.Add([System.Windows.Point]::new($xA[287],$y75[287]))
            $pts25.Add([System.Windows.Point]::new($xA[287],$y25[287]))
        }
        if ($pts75.Count -ge 3) {
            for ($i=$pts25.Count-1; $i -ge 0; $i--) { $pts75.Add($pts25[$i]) }
            $pi=New-Object System.Windows.Shapes.Polygon; $pi.Points=$pts75
            $pi.Fill=$script:CBrBandInner
            $script:HistCanvas.Children.Add($pi)|Out-Null
        }

        # Mediana - gladka linia Catmull-Rom -> StreamGeometry (biala, gruba)
        $mpx=[System.Collections.Generic.List[double]]::new(); $mpy=[System.Collections.Generic.List[double]]::new()
        for ($s=0; $s -lt 288; $s+=3) { if ($hasS[$s]) { $mpx.Add($xA[$s]); $mpy.Add($y50[$s]) } }
        if ($hasS[287]) { $mpx.Add($xA[287]); $mpy.Add($y50[287]) }
        $mn2=$mpx.Count
        if ($mn2 -ge 2) {
            $mpxA=$mpx.ToArray(); $mpyA=$mpy.ToArray()
            $mSG = New-CatmullRomStreamGeo -px $mpxA -py $mpyA -iStart 0 -iEnd ($mn2-1) -tension 0.35
            $mPath=New-Object System.Windows.Shapes.Path; $mPath.Data=$mSG
            $mPath.Stroke=$script:CBrWhite; $mPath.StrokeThickness=2.5
            $script:HistCanvas.Children.Add($mPath)|Out-Null
        }

        # Etykiety osi X: godziny co 3h (00:00 .. 24:00)
        if ($script:HasVisualHost) {
            $dvXAxis = New-Object System.Windows.Media.DrawingVisual
            $dcXA = $dvXAxis.RenderOpen()
            $dpi = $script:CDpiScale
            for ($h=0; $h -le 24; $h+=3) {
                $hh=$h%24; $xPos=if($h -eq 24){$padL+$gW}else{$padL+($hh*12)/287.0*$gW}
                $ft = New-Object System.Windows.Media.FormattedText("$($hh.ToString('00')):00", [System.Globalization.CultureInfo]::InvariantCulture, [System.Windows.FlowDirection]::LeftToRight, $script:CTypefaceUI, 8, $script:CBrLabel, $dpi)
                $dcXA.DrawText($ft, [System.Windows.Point]::new($xPos-10, $cH-$padB+4))
            }
            $dcXA.Close()
            Add-VisualToCanvas $script:HistCanvas $dvXAxis
        } else {
            for ($h=0; $h -le 24; $h+=3) {
                $hh=$h%24; $xPos=if($h -eq 24){$padL+$gW}else{$padL+($hh*12)/287.0*$gW}
                $dl=New-Object System.Windows.Controls.TextBlock; $dl.Text="$($hh.ToString('00')):00"; $dl.FontSize=8
                $dl.Foreground=$script:CBrLabel; $dl.FontFamily=$script:CFontUI
                [System.Windows.Controls.Canvas]::SetLeft($dl,$xPos-10); [System.Windows.Controls.Canvas]::SetTop($dl,$cH-$padB+4)
                $script:HistCanvas.Children.Add($dl)|Out-Null
            }
        }
    } catch { Write-Log "Render-HistGraph err: $($_.Exception.Message) at $($_.ScriptStackTrace)" }
}

function Show-HistoryWindow {
    # Jesli okno juz jest otwarte - zamknij je (toggle)
    if ($script:HistWin -and $script:HistWin.IsLoaded) {
        $script:HistWin.Close(); return
    }

    try {
    [xml]$hXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Glucose History" Width="560" Height="490"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="Transparent" WindowStyle="None" AllowsTransparency="True">
    <Border Background="#1a1a2e" CornerRadius="10">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="30"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Background="#12122a" CornerRadius="10,10,0,0" Name="hTitleBar">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Name="hTitleTxt" Text="  Historia glukozy"
                               Foreground="#7777aa" FontSize="11" VerticalAlignment="Center"
                               FontFamily="Segoe UI" Margin="6,0,0,0"/>
                    <Button Grid.Column="1" Name="hBtnAgp" Content="AGP" Width="42" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Wzorzec dobowy (AGP)"/>
                    <Button Grid.Column="2" Name="hBtnLine" Content="LINE" Width="42" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Odczyty glukozy (linia czasu) / Glucose readings (timeline)"/>
                    <Button Grid.Column="3" Name="hBtnBars" Content="BARS" Width="42" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Średnie stężenie glukozy (słupkowe) / Average glucose (bars)"/>
                    <Button Grid.Column="4" Name="hBtnReport" Content="HTML" Width="42" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Eksportuj raport HTML"/>
                    <Button Grid.Column="5" Name="hBtnPdf" Content="PDF" Width="42" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Eksportuj raport PDF"/>
                    <Button Grid.Column="6" Name="hBtnCsv" Content="CSV" Width="40" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Eksportuj dane do pliku CSV"/>
                    <Button Grid.Column="7" Name="hBtnLog" Content="LOG" Width="40" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Dziennik / Logbook"/>
                    <Button Grid.Column="8" Name="hBtnSmooth" Content="~" Width="34" Height="30"
                            Background="#1e2a3a" Foreground="#6677aa" BorderThickness="0"
                            FontSize="13" Cursor="Hand" ToolTip="Wygładzanie danych (Savitzky-Golay) / Data smoothing"/>
                    <Button Grid.Column="9" Name="hClose" Content="&#x2715;" Width="34" Height="30"
                            Background="Transparent" Foreground="#aa5555" BorderThickness="0"
                            FontSize="13" Cursor="Hand"/>
                </Grid>
            </Border>

            <Grid Grid.Row="1" Margin="12,8,12,4">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button Grid.Column="0" Name="hBtn1"  Content="1 dzień"  Background="#1e2a3a" Foreground="#7777aa"
                        BorderThickness="0" Padding="0,6" FontSize="11" Cursor="Hand" FontFamily="Segoe UI"/>
                <Button Grid.Column="2" Name="hBtn7"  Content="7 dni"  Background="#1e2a3a" Foreground="#7777aa"
                        BorderThickness="0" Padding="0,6" FontSize="11" Cursor="Hand" FontFamily="Segoe UI"/>
                <Button Grid.Column="4" Name="hBtn14" Content="14 dni" Background="#1e2a3a" Foreground="#7777aa"
                        BorderThickness="0" Padding="0,6" FontSize="11" Cursor="Hand" FontFamily="Segoe UI"/>
                <Button Grid.Column="6" Name="hBtn30" Content="30 dni" Background="#2a3a5a" Foreground="White"
                        BorderThickness="0" Padding="0,6" FontSize="11" Cursor="Hand" FontFamily="Segoe UI"/>
                <Button Grid.Column="8" Name="hBtn90" Content="90 dni" Background="#1e2a3a" Foreground="#7777aa"
                        BorderThickness="0" Padding="0,6" FontSize="11" Cursor="Hand" FontFamily="Segoe UI"/>
            </Grid>

            <Grid Grid.Row="2" Margin="12,0,12,4">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Button Grid.Column="0" Name="hBtnPrev" Content="&#x25C4;" Width="28" Height="22"
                        Background="#1e2a3a" Foreground="#7777aa" BorderThickness="0"
                        FontSize="10" Cursor="Hand" ToolTip="Przewin wstecz"/>
                <TextBlock Grid.Column="1" Name="hRangeLabel" Text=""
                           Foreground="#7777aa" FontSize="9" FontFamily="Segoe UI"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
                <Button Grid.Column="2" Name="hBtnNext" Content="&#x25BA;" Width="28" Height="22"
                        Background="#1e2a3a" Foreground="#7777aa" BorderThickness="0"
                        FontSize="10" Cursor="Hand" ToolTip="Przewin naprzod"/>
            </Grid>

            <Border Grid.Row="3" Background="#222244" CornerRadius="6" Margin="12,0,12,6" Padding="8,5">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <!-- Wiersz 1: Avg Min Max TIR -->
                    <StackPanel Grid.Row="0" Grid.Column="0" HorizontalAlignment="Center">
                        <TextBlock Name="hLblAvg"   Text="Srednia"  Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValAvg"   Text="---"      Foreground="White"   FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="0" Grid.Column="1" HorizontalAlignment="Center">
                        <TextBlock                  Text="Min"      Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValMin"   Text="---"      Foreground="#44aaff" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="0" Grid.Column="2" HorizontalAlignment="Center">
                        <TextBlock                  Text="Max"      Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValMax"   Text="---"      Foreground="#ffaa44" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="0" Grid.Column="3" HorizontalAlignment="Center">
                        <TextBlock Name="hLblTIR"   Text="W normie" Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValTIR"   Text="---"      Foreground="#44DD44" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <!-- Wiersz 2: eHbA1c SD CV% Delta -->
                    <StackPanel Grid.Row="1" Grid.Column="0" HorizontalAlignment="Center" Margin="0,4,0,0">
                        <TextBlock Name="hLblHbA1c" Text="eHbA1c"   Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValHbA1c" Text="---"       Foreground="#cc88ff" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="1" Grid.Column="1" HorizontalAlignment="Center" Margin="0,4,0,0">
                        <TextBlock                  Text="SD"        Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValSD"    Text="---"       Foreground="#88ccff" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="1" Grid.Column="2" HorizontalAlignment="Center" Margin="0,4,0,0">
                        <TextBlock                  Text="CV%"       Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValCV"    Text="---"       Foreground="#ffcc66" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Row="1" Grid.Column="3" HorizontalAlignment="Center" Margin="0,4,0,0">
                        <TextBlock Name="hLblDelta" Text="Delta avg" Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValDelta" Text="---"       Foreground="#aaaacc" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                </Grid>
            </Border>

            <Grid Grid.Row="4" Margin="12,0,12,12">
                <Border Background="#222244" CornerRadius="8">
                    <Canvas Name="hCanvas" ClipToBounds="True"/>
                </Border>
                <TextBlock Name="hNoData" Text="Brak danych historycznych"
                           Foreground="#7777aa" FontSize="13" FontFamily="Segoe UI"
                           HorizontalAlignment="Center" VerticalAlignment="Center"
                           Visibility="Collapsed"/>
            </Grid>
        </Grid>
    </Border>
</Window>
"@
    $hr = New-Object System.Xml.XmlNodeReader $hXaml
    $script:HistWin = [Windows.Markup.XamlReader]::Load($hr)

    # Przechowuj referencje w zmiennych $script: (dostepne w handlerach klikniecia)
    $script:HistCanvas      = $script:HistWin.FindName("hCanvas")
    $script:HistNoData      = $script:HistWin.FindName("hNoData")
    $script:HistValAvg      = $script:HistWin.FindName("hValAvg")
    $script:HistValMin      = $script:HistWin.FindName("hValMin")
    $script:HistValMax      = $script:HistWin.FindName("hValMax")
    $script:HistValTIR      = $script:HistWin.FindName("hValTIR")
    $script:HistValDelta    = $script:HistWin.FindName("hValDelta")
    $script:HistLblTitleTxt = $script:HistWin.FindName("hTitleTxt")
    $script:HistLblAvg      = $script:HistWin.FindName("hLblAvg")
    $script:HistLblTIR      = $script:HistWin.FindName("hLblTIR")
    $script:HistValHbA1c    = $script:HistWin.FindName("hValHbA1c")
    $script:HistRangeLabel  = $script:HistWin.FindName("hRangeLabel")
    $script:HistValSD       = $script:HistWin.FindName("hValSD")
    $script:HistValCV       = $script:HistWin.FindName("hValCV")
    $script:HistBtnSmooth   = $script:HistWin.FindName("hBtnSmooth")
    $script:HistBtnLine     = $script:HistWin.FindName("hBtnLine")
    # Stan wizualny przycisku smooth - zsynchronizowany z glownym oknem
    if ($script:SmoothMode) {
        $script:HistBtnSmooth.Foreground = [System.Windows.Media.Brushes]::White
        $script:HistBtnSmooth.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#2a3a5a"))
    }
    # Stan wizualny przycisku LINE - zsynchronizowany z trybem widoku
    if ($script:HistViewMode -eq "timeline") {
        $script:HistBtnLine.Foreground = [System.Windows.Media.Brushes]::White
        $script:HistBtnLine.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#2a3a5a"))
    }
    $script:HistBtns     = @(
        $script:HistWin.FindName("hBtn1"),
        $script:HistWin.FindName("hBtn7"),
        $script:HistWin.FindName("hBtn14"),
        $script:HistWin.FindName("hBtn30"),
        $script:HistWin.FindName("hBtn90")
    )

    # Jezyk
    $ttxt = $script:HistWin.FindName("hTitleTxt")
    $ttxt.Text = (t "HistTitle")
    $script:HistWin.FindName("hLblAvg").Text   = (t "HistAvg")
    $script:HistWin.FindName("hLblTIR").Text   = (t "HistTIR")
    $script:HistWin.FindName("hLblDelta").Text = "Delta avg"
    $script:HistWin.FindName("hNoData").Text   = (t "HistNoData")
    $day1Label = if ($script:LangEn) { "1 day" } else { "1 dzie$([char]0x0144)" }
    $script:HistBtns[0].Content = $day1Label
    $script:HistBtns[1].Content = "7 "  + (t "HistDays")
    $script:HistBtns[2].Content = "14 " + (t "HistDays")
    $script:HistBtns[3].Content = "30 " + (t "HistDays")
    $script:HistBtns[4].Content = "90 " + (t "HistDays")

    # Drag paska tytuloweg
    $script:HistWin.FindName("hTitleBar").Add_MouseLeftButtonDown({ $script:HistWin.DragMove() })

    $script:HistDays   = 1
    $script:HistOffset = 0

    # Renderuj przy otwarciu (Add_ContentRendered wykonuje sie po Show())
    $script:HistWin.Add_ContentRendered({ Render-HistGraph $script:HistDays })

    # Handlery przyciskow okresu - reset offset i zoom przy zmianie okresu
    $script:HistWin.FindName("hClose").Add_Click({ $script:HistWin.Close() })
    $script:HistBtns[0].Add_Click({ $script:HistDays = 1;  $script:HistOffset = 0; $script:HistZoomFactor = 1.0; $script:HistPanOffset = 0.0; Render-HistGraph 1  })
    $script:HistBtns[1].Add_Click({ $script:HistDays = 7;  $script:HistOffset = 0; $script:HistZoomFactor = 1.0; $script:HistPanOffset = 0.0; Render-HistGraph 7  })
    $script:HistBtns[2].Add_Click({ $script:HistDays = 14; $script:HistOffset = 0; $script:HistZoomFactor = 1.0; $script:HistPanOffset = 0.0; Render-HistGraph 14 })
    $script:HistBtns[3].Add_Click({ $script:HistDays = 30; $script:HistOffset = 0; $script:HistZoomFactor = 1.0; $script:HistPanOffset = 0.0; Render-HistGraph 30 })
    $script:HistBtns[4].Add_Click({ $script:HistDays = 90; $script:HistOffset = 0; $script:HistZoomFactor = 1.0; $script:HistPanOffset = 0.0; Render-HistGraph 90 })

    # Nawigacja wstecz / naprzod
    # Prev (◄) = starsze dane = wiekszy offset od dzis
    # Next (►) = nowsze dane = mniejszy offset (wroc do terazniejszosci)
    $script:HistWin.FindName("hBtnPrev").Add_Click({
        $script:HistOffset += [int]($script:HistDays / 2)
        Render-HistGraph $script:HistDays
    })
    $script:HistWin.FindName("hBtnNext").Add_Click({
        $script:HistOffset = [Math]::Max(0, $script:HistOffset - [int]($script:HistDays / 2))
        Render-HistGraph $script:HistDays
    })

    # Przelaczanie widoku LINE (timeline) / AGP
    $script:HistBtnLine.Add_Click({
        if ($script:HistViewMode -eq "timeline") {
            $script:HistViewMode = "agp"
            $script:HistBtnLine.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#6677aa"))
            $script:HistBtnLine.Background = [System.Windows.Media.Brushes]::Transparent
        } else {
            $script:HistViewMode = "timeline"
            $script:HistBtnLine.Foreground = [System.Windows.Media.Brushes]::White
            $script:HistBtnLine.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#2a3a5a"))
            # Reset zoomu przy przechodzeniu do trybu LINE
            $script:HistZoomFactor = 1.0
            $script:HistPanOffset = 0.0
        }
        Render-HistGraph $script:HistDays
    })

    # Eksport CSV
    $script:HistWin.FindName("hBtnCsv").Add_Click({ Export-HistoryCSV })
    $script:HistWin.FindName("hBtnAgp").Add_Click({ Show-AgpWindow })
    $script:HistWin.FindName("hBtnBars").Add_Click({ Show-AvgBarsWindow })
    $script:HistWin.FindName("hBtnLog").Add_Click({ Show-LogbookWindow })
    $script:HistWin.FindName("hBtnReport").Add_Click({ Export-HtmlReport })
    $script:HistWin.FindName("hBtnPdf").Add_Click({ Export-PdfReport })
    $script:HistBtnSmooth.Add_Click({
        $script:SmoothMode = -not $script:SmoothMode
        $col = if ($script:SmoothMode) { [System.Windows.Media.Brushes]::White } else { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#6677aa")) }
        $bg  = if ($script:SmoothMode) { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#2a3a5a")) } else { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#1e2a3a")) }
        $script:btnSmooth.Foreground = $col; $script:btnSmooth.Background = $bg
        $script:HistBtnSmooth.Foreground = $col; $script:HistBtnSmooth.Background = $bg
        Save-Config
        if ($script:CachedGraphData) { Update-Graph $script:CachedGraphData }
        Render-HistGraph $script:HistDays
    })

    # Zapamietaj pozycje przy zamknieciu
    $script:HistWin.Add_Closing({
        $script:HistWinLeft = $script:HistWin.Left
        $script:HistWinTop  = $script:HistWin.Top
    })

    # Przywroc zapamietana pozycje (jesli istnieje)
    if ($null -ne $script:HistWinLeft -and $null -ne $script:HistWinTop) {
        $script:HistWin.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
        $script:HistWin.Left = $script:HistWinLeft
        $script:HistWin.Top  = $script:HistWinTop
    }

    # Obsluga MouseWheel na Canvas - zoom dla trybu LINE z throttlingiem
    $script:HistCanvas.Add_MouseWheel({
        param($sender, $e)
        try {
            if ($script:HistViewMode -ne "timeline") { return }
            
            # Zoom factor: góra = zoom in, dół = zoom out
            # Mniejszy krok (1.10) = wolniejszy zoom ale mniej zaciec
            $delta = $e.Delta
            $zoomStep = if ($delta -gt 0) { 1.10 } else { 1/1.10 }
            $script:HistZoomFactor *= $zoomStep
            
            # Ogranicz zoom: min 50%, max 2000%
            if ($script:HistZoomFactor -lt 0.5)  { $script:HistZoomFactor = 0.5 }
            if ($script:HistZoomFactor -gt 20.0) { $script:HistZoomFactor = 20.0 }
            
            # Throttling: odloz przerysowanie o 150ms, restart timer przy kolejnym ruchu rolka
            # Z cache renderowanie jest szybkie wiec krotsze opoznienie
            if ($script:HistZoomTimer) {
                $script:HistZoomTimer.Stop()
            } else {
                $script:HistZoomTimer = New-Object System.Windows.Threading.DispatcherTimer
                $script:HistZoomTimer.Interval = [TimeSpan]::FromMilliseconds(150)
                $script:HistZoomTimer.Add_Tick({
                    $script:HistZoomTimer.Stop()
                    Render-HistGraph $script:HistDays $true  # $useCache = $true dla szybkiego zoomu
                })
            }
            $script:HistZoomTimer.Start()
            
            $e.Handled = $true
        } catch {
            Write-Log "MouseWheel error: $($_.Exception.Message)"
        }
    })

    $script:HistWin.Show()

    } catch { 
        Write-Log "Show-HistoryWindow error: $($_.Exception.Message)"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
        [System.Windows.MessageBox]::Show("Blad okna historii: $($_.Exception.Message)", "Glucose Monitor") | Out-Null
    }
}

function Show-LogbookWindow {
    # Toggle - jesli juz otwarte, zamknij
    if ($script:LogWin -and $script:LogWin.IsLoaded) { $script:LogWin.Close(); return }
    try {
    [xml]$lXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Logbook" Width="420" Height="480"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="Transparent" WindowStyle="None" AllowsTransparency="True">
    <Border Background="#1a1a2e" CornerRadius="10">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="30"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="#12122a" CornerRadius="10,10,0,0" Name="logTitleBar">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Name="logTitleTxt" Text="  Dziennik"
                               Foreground="#7777aa" FontSize="11" VerticalAlignment="Center"
                               FontFamily="Segoe UI" Margin="6,0,0,0"/>
                    <Button Grid.Column="1" Name="logClose" Content="&#x2715;" Width="34" Height="30"
                            Background="Transparent" Foreground="#aa5555" BorderThickness="0"
                            FontSize="13" Cursor="Hand"/>
                </Grid>
            </Border>
            <TextBlock Grid.Row="1" Name="logStatusTxt" Text="Pobieranie..."
                       Foreground="#556688" FontSize="10" FontFamily="Segoe UI"
                       Margin="12,6,12,2"/>
            <ScrollViewer Grid.Row="2" Margin="8,4,8,8"
                          VerticalScrollBarVisibility="Auto"
                          HorizontalScrollBarVisibility="Disabled">
                <StackPanel Name="logStack" Margin="0,0,2,4"/>
            </ScrollViewer>
        </Grid>
    </Border>
</Window>
"@
    $lr = New-Object System.Xml.XmlNodeReader $lXaml
    $script:LogWin = [Windows.Markup.XamlReader]::Load($lr)

    $script:LogStatusTxt = $script:LogWin.FindName("logStatusTxt")
    $script:LogStack     = $script:LogWin.FindName("logStack")

    $script:LogWin.FindName("logTitleBar").Add_MouseLeftButtonDown({ $script:LogWin.DragMove() })
    $script:LogWin.FindName("logClose").Add_Click({ $script:LogWin.Close() })
    $script:LogWin.FindName("logTitleTxt").Text = "  " + (t "LogTitle")
    $script:LogStatusTxt.Text = (t "LogLoading")

    $script:LogWin.Add_Closing({
        $script:LogWinLeft = $script:LogWin.Left
        $script:LogWinTop  = $script:LogWin.Top
    })
    if ($null -ne $script:LogWinLeft -and $null -ne $script:LogWinTop) {
        $script:LogWin.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
        $script:LogWin.Left = $script:LogWinLeft
        $script:LogWin.Top  = $script:LogWinTop
    }

    $script:LogWin.Add_ContentRendered({
        $fmts = @("M/d/yyyy h:mm:ss tt","d/M/yyyy h:mm:ss tt","M/d/yyyy H:mm:ss","d/M/yyyy H:mm:ss","yyyy-MM-ddTHH:mm:ss")
        try {
            # Najpierw pobierz z API (+ automatyczny zapis do logbook.jsonl)
            $data = Get-LogbookData

            # Zaladuj lokalne wpisy (zawsze, nawet jesli API powiodlo sie - moze miec starsze)
            $localItems = Load-LogbookLocal

            # Parsuj wpisy z API
            $apiItems = [System.Collections.Generic.List[object]]::new()
            if ($data -and $data.Count -gt 0) {
                foreach ($entry in $data) {
                    $mg = 0.0
                    foreach ($fld in @('ValueInMgPerDl','Value','GlucoseValue','value','valueInMgPerDl')) {
                        if ($entry.$fld -and [double]$entry.$fld -gt 20) { $mg = [double]$entry.$fld; break }
                    }
                    if ($mg -le 0) { continue }
                    $dt    = [DateTime]::MinValue
                    $tsStr = if ($entry.Timestamp) { "$($entry.Timestamp)" } else { "$($entry.timestamp)" }
                    if ($tsStr) {
                        foreach ($f in $fmts) {
                            if ([DateTime]::TryParseExact($tsStr, $f, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { break }
                        }
                    }
                    $isScan = try { [int]$entry.type -eq 1 } catch { $false }
                    $trend  = try { [int]$entry.TrendArrow } catch { 0 }
                    $apiItems.Add([PSCustomObject]@{ dt=$dt; mg=$mg; isScan=$isScan; trend=$trend })
                }
            }

            # Polacz API + lokalne, deduplikuj po minucie, lokalne sa uzupelnieniem
            $seen  = [System.Collections.Generic.HashSet[string]]::new()
            $items = [System.Collections.Generic.List[object]]::new()
            foreach ($it in $apiItems) {
                if ($it.dt -ne [DateTime]::MinValue) {
                    $k = $it.dt.ToString("yyyy-MM-ddTHH:mm")
                    if ($seen.Add($k)) { $items.Add($it) }
                }
            }
            foreach ($it in $localItems) {
                if ($it.dt -ne [DateTime]::MinValue) {
                    $k = $it.dt.ToString("yyyy-MM-ddTHH:mm")
                    if ($seen.Add($k)) { $items.Add($it) }
                }
            }

            if ($items.Count -eq 0) {
                $script:LogStatusTxt.Text = if ($null -eq $data) {
                    if ($script:LangEn) { "Error loading data" } else { "Blad pobierania danych" }
                } else { (t "LogNoData") }
                return
            }

            $sorted  = $items | Sort-Object dt -Descending
            $cntScan = ($sorted | Where-Object { $_.isScan }).Count
            $srcNote = if ($null -eq $data) { " (offline)" } else { "" }
            $statusLine = "$(t 'LogBtn'): $($sorted.Count)$srcNote"
            if ($cntScan -gt 0) { $statusLine += "  |  $(t 'LogScan'): $cntScan" }
            $script:LogStatusTxt.Text = $statusLine

            $script:LogStack.Children.Clear()
            foreach ($item in $sorted) {
                $mmol  = MgToMmol $item.mg
                $val   = if ($script:UseMgDl) { "$([Math]::Round($item.mg,0))" } else { "$mmol" }
                $unit  = if ($script:UseMgDl) { "mg/dL" } else { "mmol/L" }
                $color = Get-GlucoseColor $mmol
                $arrow = Get-TrendArrow $item.trend
                $tStr  = if ($item.dt -ne [DateTime]::MinValue) { $item.dt.ToString("dd.MM  HH:mm") } else { "--:--" }
                $typeStr = if ($item.isScan) { (t "LogScan") } else { (t "LogAuto") }
                $typeFg  = if ($item.isScan) { "#aabbff" } else { "#445566" }
                $rowBg   = if ($item.isScan) { "#1e2244" } else { "#161628" }
                $dotCh   = if ($item.isScan) { [char]0x25CF } else { [char]0x25CB }  # filled vs empty circle

                $row = New-Object System.Windows.Controls.Border
                $row.Background   = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($rowBg))
                $row.Margin       = New-Object System.Windows.Thickness(0,0,0,2)
                $row.CornerRadius  = New-Object System.Windows.CornerRadius(4)
                $row.Padding       = New-Object System.Windows.Thickness(8,5,8,5)

                $g = New-Object System.Windows.Controls.Grid
                foreach ($cw in @([System.Windows.GridLength]::Auto,
                                  [System.Windows.GridLength]::Auto,
                                  (New-Object System.Windows.GridLength(1,[System.Windows.GridUnitType]::Star)),
                                  [System.Windows.GridLength]::Auto,
                                  [System.Windows.GridLength]::Auto)) {
                    $cd = New-Object System.Windows.Controls.ColumnDefinition; $cd.Width = $cw
                    $g.ColumnDefinitions.Add($cd)
                }

                $tbDot = New-Object System.Windows.Controls.TextBlock
                $tbDot.Text = $dotCh; $tbDot.FontSize = 9
                $tbDot.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($typeFg))
                $tbDot.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                $tbDot.Margin = New-Object System.Windows.Thickness(0,0,6,0)
                [System.Windows.Controls.Grid]::SetColumn($tbDot, 0)

                $tbTime = New-Object System.Windows.Controls.TextBlock
                $tbTime.Text = $tStr; $tbTime.FontSize = 11; $tbTime.MinWidth = 84
                $tbTime.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
                $tbTime.Foreground = [System.Windows.Media.Brushes]::LightGray
                $tbTime.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                [System.Windows.Controls.Grid]::SetColumn($tbTime, 1)

                $tbGlc = New-Object System.Windows.Controls.TextBlock
                $tbGlc.Text = "$val $unit"; $tbGlc.FontSize = 13
                $tbGlc.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI Semibold")
                $tbGlc.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($color))
                $tbGlc.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                $tbGlc.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
                [System.Windows.Controls.Grid]::SetColumn($tbGlc, 3)

                $tbTyp = New-Object System.Windows.Controls.TextBlock
                $tbTyp.Text = "  $arrow  $typeStr"; $tbTyp.FontSize = 10; $tbTyp.MinWidth = 70
                $tbTyp.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
                $tbTyp.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($typeFg))
                $tbTyp.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                $tbTyp.TextAlignment = [System.Windows.TextAlignment]::Right
                [System.Windows.Controls.Grid]::SetColumn($tbTyp, 4)

                foreach ($c in @($tbDot, $tbTime, $tbGlc, $tbTyp)) { $g.Children.Add($c) | Out-Null }
                $row.Child = $g
                $script:LogStack.Children.Add($row) | Out-Null
            }
        } catch {
            Write-Log "Logbook render ERR: $($_.Exception.Message)"
            try { $script:LogStatusTxt.Text = "Error: $($_.Exception.Message)" } catch {}
        }
    })

    $script:LogWin.Show()
    } catch {
        Write-Log "Show-LogbookWindow error: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("$(if($script:LangEn){'Logbook error'}else{'Blad dziennika'}): $($_.Exception.Message)", "Glucose Monitor") | Out-Null
    }
}

# ======================== BACKUP / RESTORE ========================
function Backup-History {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title  = if ($script:LangEn) { "Save history backup" } else { "Zapisz kopie zapasowa historii" }
    $dlg.Filter = "JSONL files (*.jsonl)|*.jsonl|All files (*.*)|*.*"
    $dlg.FileName = "glucose_backup_$(Get-Date -Format 'yyyyMMdd').jsonl"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            Copy-Item -Path $script:HistoryFile -Destination $dlg.FileName -Force
            [System.Windows.MessageBox]::Show(
                $(if ($script:LangEn) { "Backup saved to:`n$($dlg.FileName)" } else { "Kopia zapisana do:`n$($dlg.FileName)" }),
                "Glucose Monitor") | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show("Blad: $($_.Exception.Message)", "Glucose Monitor") | Out-Null
        }
    }
}

function Restore-History {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title  = if ($script:LangEn) { "Open history backup" } else { "Otworz kopie zapasowa historii" }
    $dlg.Filter = "JSONL files (*.jsonl)|*.jsonl|All files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $existing = [System.Collections.Generic.HashSet[string]]::new()
            if (Test-Path $script:HistoryFile) {
                Get-Content $script:HistoryFile -Encoding UTF8 | ForEach-Object {
                    try { $existing.Add(($_ | ConvertFrom-Json).ts) | Out-Null } catch {}
                }
            }
            $added = 0
            Get-Content $dlg.FileName -Encoding UTF8 | ForEach-Object {
                try {
                    $obj = $_ | ConvertFrom-Json
                    if ($obj.ts -and -not $existing.Contains($obj.ts)) {
                        Add-Content -Path $script:HistoryFile -Value $_ -Encoding UTF8
                        $existing.Add($obj.ts) | Out-Null
                        $added++
                    }
                } catch {}
            }
            $script:HistKnownTs = $null  # wymus reinicjalizacje HashSet
            [System.Windows.MessageBox]::Show(
                $(if ($script:LangEn) { "Restored $added new entries." } else { "Przywrocono $added nowych wpisow." }),
                "Glucose Monitor") | Out-Null
            if ($script:HistWin -and $script:HistWin.IsLoaded) { Render-HistGraph $script:HistDays }
        } catch {
            [System.Windows.MessageBox]::Show("Blad: $($_.Exception.Message)", "Glucose Monitor") | Out-Null
        }
    }
}

# ======================== USTAWIENIA ========================
function Show-SettingsWindow {
    $loVal  = if ($script:UseMgDl) { [Math]::Round($script:Config.AlertLow * 18.018, 0) } else { $script:Config.AlertLow }
    $hiVal  = if ($script:UseMgDl) { [Math]::Round($script:Config.AlertHigh * 18.018, 0) } else { $script:Config.AlertHigh }
    $unit   = if ($script:UseMgDl) { "mg/dL" } else { "mmol/L" }
    $xamlS  = [xml]@"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$(if ($script:LangEn){'Settings'}else{'Ustawienia'})" Width="320" Height="280"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" WindowStartupLocation="CenterOwner">
    <Border Background="#111827" CornerRadius="10" BorderBrush="#334466" BorderThickness="1">
        <Grid Margin="16">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="$(if ($script:LangEn){'Settings'}else{'Ustawienia'})"
                       Foreground="White" FontSize="14" FontFamily="Segoe UI Semibold" Margin="0,0,0,12"/>
            <StackPanel Grid.Row="1" VerticalAlignment="Top">
                <TextBlock Text="$(if ($script:LangEn){"Alert threshold LOW ($unit)"}else{"Prog alarmu NISKI ($unit)"})"
                           Foreground="#7777aa" FontSize="10" FontFamily="Segoe UI" Margin="0,0,0,2"/>
                <TextBox Name="sAlertLow" Text="$loVal" Background="#1e2a3a" Foreground="White"
                         BorderBrush="#334466" BorderThickness="1" Padding="6,4" FontSize="12"
                         FontFamily="Segoe UI" Margin="0,0,0,10"/>
                <TextBlock Text="$(if ($script:LangEn){"Alert threshold HIGH ($unit)"}else{"Prog alarmu WYSOKI ($unit)"})"
                           Foreground="#7777aa" FontSize="10" FontFamily="Segoe UI" Margin="0,0,0,2"/>
                <TextBox Name="sAlertHigh" Text="$hiVal" Background="#1e2a3a" Foreground="White"
                         BorderBrush="#334466" BorderThickness="1" Padding="6,4" FontSize="12"
                         FontFamily="Segoe UI" Margin="0,0,0,10"/>
                <TextBlock Text="$(if ($script:LangEn){'Refresh interval (seconds)'}else{'Interwal odswiezania (sekundy)'})"
                           Foreground="#7777aa" FontSize="10" FontFamily="Segoe UI" Margin="0,0,0,2"/>
                <TextBox Name="sInterval" Text="$($script:Config.Interval)" Background="#1e2a3a" Foreground="White"
                         BorderBrush="#334466" BorderThickness="1" Padding="6,4" FontSize="12"
                         FontFamily="Segoe UI"/>
            </StackPanel>
            <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
                <Button Name="sCancel" Content="$(if ($script:LangEn){'Cancel'}else{'Anuluj'})" Width="80" Height="28"
                        Background="#1e2a3a" Foreground="#7777aa" BorderThickness="0" Margin="0,0,8,0" Cursor="Hand"/>
                <Button Name="sSave" Content="$(if ($script:LangEn){'Save'}else{'Zapisz'})" Width="80" Height="28"
                        Background="#334488" Foreground="White" BorderThickness="0" Cursor="Hand"/>
            </StackPanel>
        </Grid>
    </Border>
</Window>
"@
    try {
        $sw = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xamlS))
        $sw.FindName("sCancel").Add_Click({ $sw.Close() })
        $sw.FindName("sSave").Add_Click({
            try {
                $newLo  = [double]($sw.FindName("sAlertLow").Text  -replace ',','.')
                $newHi  = [double]($sw.FindName("sAlertHigh").Text -replace ',','.')
                $newInt = [int]($sw.FindName("sInterval").Text)
                # Konwertuj z powrotem na mmol jesli trzeba
                if ($script:UseMgDl) { $newLo = [Math]::Round($newLo / 18.018, 1); $newHi = [Math]::Round($newHi / 18.018, 1) }
                $script:Config.AlertLow  = $newLo
                $script:Config.AlertHigh = $newHi
                if ($newInt -ge 30) { $script:Config.Interval = $newInt; $script:SecondsLeft = $newInt }
                Save-Config
                $sw.Close()
            } catch {
                [System.Windows.MessageBox]::Show("$(if ($script:LangEn){'Invalid value'}else{'Nieprawidlowa wartosc'}): $($_.Exception.Message)") | Out-Null
            }
        })
        $sw.Add_MouseLeftButtonDown({ $sw.DragMove() })
        $sw.ShowDialog() | Out-Null
    } catch { Write-Log "Settings ERR: $($_.Exception.Message)" }
}

# ======================== SVG CHARTS (dla PDF) ========================

function Get-HistSvg {
    param([object[]]$Data, [double]$LoN, [double]$HiN, [bool]$UseMgDl, [string]$Unit, [bool]$LangEn)
    try {
        $pts = [System.Collections.Generic.List[object]]::new()
        foreach ($d in ($Data | Sort-Object { [DateTime]::Parse($_.ts) })) {
            try {
                $pts.Add([PSCustomObject]@{
                    ts = [DateTime]::Parse($d.ts)
                    v  = if ($UseMgDl) { [double]$d.mgdl } else { [Math]::Round([double]$d.mgdl / 18.018, 2) }
                })
            } catch {}
        }
        if ($pts.Count -lt 2) {
            return "<p style='color:#889;font-size:11px;padding:10px 0'>$(if($LangEn){'Not enough data for chart.'}else{'Za malo danych dla wykresu.'})</p>"
        }

        [double]$W = 760; [double]$H = 195
        [double]$pL = 44; [double]$pR = 8; [double]$pT = 12; [double]$pB = 26
        [double]$gW = $W - $pL - $pR; [double]$gH = $H - $pT - $pB

        $t0 = $pts[0].ts; $t1 = $pts[$pts.Count - 1].ts
        [double]$tRng = ($t1 - $t0).TotalMinutes
        if ($tRng -lt 1) { return "" }

        [double]$vMin = [double]::MaxValue; [double]$vMax = [double]::MinValue
        foreach ($p in $pts) { if ($p.v -lt $vMin) { $vMin = $p.v }; if ($p.v -gt $vMax) { $vMax = $p.v } }
        [double]$minPad = if ($UseMgDl) { 5.0 } else { 0.3 }
        [double]$pad = [Math]::Max(($vMax - $vMin) * 0.08, $minPad)
        [double]$yLo = [Math]::Max(0, $vMin - $pad)
        [double]$yHi = $vMax + $pad
        [double]$yRng = $yHi - $yLo
        if ($yRng -lt 0.01) { return "" }

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 $W $H' style='width:100%;height:auto;display:block'>")
        [void]$sb.Append("<rect width='$W' height='$H' fill='#f8f9ff' rx='3'/>")

        # Strefa normy (inline)
        [double]$nLoY = [Math]::Min([Math]::Max([Math]::Round($pT + $gH - ($LoN - $yLo) / $yRng * $gH, 1), $pT), $pT + $gH)
        [double]$nHiY = [Math]::Min([Math]::Max([Math]::Round($pT + $gH - ($HiN - $yLo) / $yRng * $gH, 1), $pT), $pT + $gH)
        [double]$nTop = [Math]::Min($nLoY, $nHiY); [double]$nHt = [Math]::Abs($nLoY - $nHiY)
        [void]$sb.Append("<rect x='$pL' y='$nTop' width='$gW' height='$nHt' fill='#d4edda'/>")

        # Linie poziome Y
        [double]$step = if ($UseMgDl) { 50 } else { 2 }
        [double]$yv = [Math]::Ceiling($yLo / $step) * $step
        while ($yv -le $yHi) {
            [double]$yp = [Math]::Round($pT + $gH - ($yv - $yLo) / $yRng * $gH, 1)
            $lbl = if ($UseMgDl) { [int]$yv } else { $yv.ToString("0.0") }
            [void]$sb.Append("<line x1='$pL' y1='$yp' x2='$($pL+$gW)' y2='$yp' stroke='#ccd' stroke-width='0.5' stroke-dasharray='3,2'/>")
            [void]$sb.Append("<text x='$($pL-3)' y='$($yp+3)' text-anchor='end' font-family='Arial' font-size='10' fill='#333'>$lbl</text>")
            $yv += $step
        }

        # Linie pionowe X
        [double]$totalDays = ($t1 - $t0).TotalDays
        if ($totalDays -gt 2) {
            $cur = $t0.Date.AddDays(1)
            while ($cur -le $t1) {
                [double]$xp = [Math]::Round($pL + ($cur - $t0).TotalMinutes / $tRng * $gW, 1)
                [void]$sb.Append("<line x1='$xp' y1='$pT' x2='$xp' y2='$($pT+$gH)' stroke='#dde' stroke-width='0.5'/>")
                [void]$sb.Append("<text x='$xp' y='$($pT+$gH+15)' text-anchor='middle' font-family='Arial' font-size='9' fill='#333'>$($cur.ToString('dd.MM'))</text>")
                $cur = $cur.AddDays(1)
            }
        } else {
            $cur = $t0.Date.AddHours([int]($t0.Hour / 4) * 4)
            while ($cur -le $t1) {
                if ($cur -ge $t0) {
                    [double]$xp = [Math]::Round($pL + ($cur - $t0).TotalMinutes / $tRng * $gW, 1)
                    [void]$sb.Append("<line x1='$xp' y1='$pT' x2='$xp' y2='$($pT+$gH)' stroke='#dde' stroke-width='0.5'/>")
                    [void]$sb.Append("<text x='$xp' y='$($pT+$gH+15)' text-anchor='middle' font-family='Arial' font-size='9' fill='#333'>$($cur.ToString('HH:mm'))</text>")
                }
                $cur = $cur.AddHours(4)
            }
        }

        # Linia danych - wygladzona Catmull-Rom -> cubic Bezier z wykrywaniem przerw
        $n = $pts.Count
        $pxArr = [double[]]::new($n)
        $pyArr = [double[]]::new($n)
        for ($i = 0; $i -lt $n; $i++) {
            $pxArr[$i] = [Math]::Round($pL + ($pts[$i].ts - $t0).TotalMinutes / $tRng * $gW, 1)
            $pyArr[$i] = [Math]::Round($pT + $gH - ($pts[$i].v - $yLo) / $yRng * $gH, 1)
        }
        [double]$maxGapMin = 20.0  # przerwa > 20 min = brak odczytu = nowy segment
        [double]$tension = 0.2     # wspolczynnik wygladzania (0=proste, 1/6=Catmull-Rom)
        $pathSb = New-Object System.Text.StringBuilder
        $segStart = 0
        while ($segStart -lt $n) {
            $segEnd = $segStart
            while ($segEnd + 1 -lt $n -and ($pts[$segEnd+1].ts - $pts[$segEnd].ts).TotalMinutes -le $maxGapMin) {
                $segEnd++
            }
            if ($segEnd -gt $segStart) {
                [void]$pathSb.Append("M $($pxArr[$segStart]),$($pyArr[$segStart])")
                for ($i = $segStart; $i -lt $segEnd; $i++) {
                    $i0 = if ($i -gt $segStart) { $i - 1 } else { $segStart }
                    $i3 = if ($i + 2 -le $segEnd) { $i + 2 } else { $segEnd }
                    [double]$cp1x = [Math]::Round($pxArr[$i]   + ($pxArr[$i+1] - $pxArr[$i0]) * $tension, 1)
                    [double]$cp1y = [Math]::Round($pyArr[$i]   + ($pyArr[$i+1] - $pyArr[$i0]) * $tension, 1)
                    [double]$cp2x = [Math]::Round($pxArr[$i+1] - ($pxArr[$i3]  - $pxArr[$i])  * $tension, 1)
                    [double]$cp2y = [Math]::Round($pyArr[$i+1] - ($pyArr[$i3]  - $pyArr[$i])  * $tension, 1)
                    [void]$pathSb.Append(" C $cp1x,$cp1y $cp2x,$cp2y $($pxArr[$i+1]),$($pyArr[$i+1])")
                }
            }
            $segStart = $segEnd + 1
        }
        [void]$sb.Append("<path d='$($pathSb.ToString())' fill='none' stroke='#3366cc' stroke-width='1.5'/>")

        # Punkty poza norma (inline, bez scriptblokow)
        $dotN = 0
        foreach ($pt3 in $pts) {
            if ($dotN -ge 300) { break }
            if ($pt3.v -lt $LoN -or $pt3.v -gt $HiN) {
                [double]$cx = [Math]::Round($pL + ($pt3.ts - $t0).TotalMinutes / $tRng * $gW, 1)
                [double]$cy = [Math]::Round($pT + $gH - ($pt3.v - $yLo) / $yRng * $gH, 1)
                $col = if ($pt3.v -lt $LoN) { '#cc1111' } else { '#bb5500' }
                [void]$sb.Append("<circle cx='$cx' cy='$cy' r='2' fill='$col'/>")
                $dotN++
            }
        }

        # Osie
        [void]$sb.Append("<line x1='$pL' y1='$pT' x2='$pL' y2='$($pT+$gH)' stroke='#99a' stroke-width='1'/>")
        [void]$sb.Append("<line x1='$pL' y1='$($pT+$gH)' x2='$($pL+$gW)' y2='$($pT+$gH)' stroke='#99a' stroke-width='1'/>")
        [void]$sb.Append("<text x='$pL' y='$($pT-2)' font-family='Arial' font-size='9' fill='#333'>$Unit</text>")
        [void]$sb.Append("</svg>")
        return $sb.ToString()
    } catch { Write-Log "HistSvg ERR: $($_.Exception.Message)"; return "" }
}

function Get-AgpSvg {
    param([object[]]$Data, [double]$LoN, [double]$HiN, [bool]$UseMgDl, [string]$Unit, [bool]$LangEn)
    try {
        # Tablica 24 list - bezposredni indeks godziny, bez hashtable i bez pipeline
        # (pipeline enumeruje List<double> i niszczy je - trzeba petla for)
        $hourLists = [object[]]::new(24)
        for ($i = 0; $i -lt 24; $i++) { $hourLists[$i] = [System.Collections.Generic.List[double]]::new() }
        foreach ($pt in $Data) {
            try {
                $hi = [DateTime]::Parse($pt.ts).Hour
                [double]$v = if ($UseMgDl) { [double]$pt.mgdl } else { [Math]::Round([double]$pt.mgdl / 18.018, 2) }
                $hourLists[$hi].Add($v)
            } catch {}
        }
        $hoursWithData = 0
        foreach ($hl in $hourLists) { if ($hl.Count -gt 0) { $hoursWithData++ } }
        if ($hoursWithData -lt 4) {
            return "<p style='color:#889;font-size:11px;padding:10px 0'>$(if($LangEn){'Not enough data for daily pattern.'}else{'Za malo danych dla wzorca dobowego.'})</p>"
        }

        [double]$W = 760; [double]$H = 175
        [double]$pL = 44; [double]$pR = 8; [double]$pT = 12; [double]$pB = 26
        [double]$gW = $W - $pL - $pR; [double]$gH = $H - $pT - $pB
        [double]$yLo = if ($UseMgDl) { 40.0 } else { 2.0 }
        [double]$yHi = if ($UseMgDl) { 280.0 } else { 15.5 }
        [double]$yRng = $yHi - $yLo
        [double]$barW = $gW / 24 * 0.55

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 $W $H' style='width:100%;height:auto;display:block'>")
        [void]$sb.Append("<rect width='$W' height='$H' fill='#f8f9ff' rx='3'/>")

        # Strefa normy (inline)
        [double]$nLoY = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($LoN,$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
        [double]$nHiY = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($HiN,$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
        [double]$nTop = [Math]::Min($nLoY, $nHiY); [double]$nHt = [Math]::Abs($nLoY - $nHiY)
        [void]$sb.Append("<rect x='$pL' y='$nTop' width='$gW' height='$nHt' fill='#d4edda'/>")

        # Linie poziome Y
        [double]$step = if ($UseMgDl) { 50 } else { 2 }
        [double]$yv = [Math]::Ceiling($yLo / $step) * $step
        while ($yv -le $yHi) {
            [double]$yp = [Math]::Round($pT + $gH - ($yv - $yLo) / $yRng * $gH, 1)
            $lbl = if ($UseMgDl) { [int]$yv } else { $yv.ToString("0.0") }
            [void]$sb.Append("<line x1='$pL' y1='$yp' x2='$($pL+$gW)' y2='$yp' stroke='#ccd' stroke-width='0.5' stroke-dasharray='3,2'/>")
            [void]$sb.Append("<text x='$($pL-3)' y='$($yp+3)' text-anchor='end' font-family='Arial' font-size='10' fill='#333'>$lbl</text>")
            $yv += $step
        }

        # Slupki per godzina + linia srednich (inline, bez scriptblokow)
        $avgPts = New-Object System.Collections.Generic.List[string]
        for ($hi = 0; $hi -lt 24; $hi++) {
            [double]$xC = [Math]::Round($pL + ($hi + 0.5) / 24 * $gW, 1)
            if ($hi % 3 -eq 0) {
                [void]$sb.Append("<text x='$xC' y='$($pT+$gH+16)' text-anchor='middle' font-family='Arial' font-size='9' fill='#333'>${hi}h</text>")
            }
            $hiVals = $hourLists[$hi]
            if ($hiVals.Count -gt 0) {
                $sm = $hiVals | Measure-Object -Average -Minimum -Maximum
                [double]$avg  = $sm.Average
                [double]$minV = $sm.Minimum
                [double]$maxV = $sm.Maximum
                [double]$yAvg = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($avg,$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
                [double]$yMin = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($minV,$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
                [double]$yMax = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($maxV,$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
                [double]$bTop = [Math]::Min($yMin, $yMax)
                [double]$bHt  = [Math]::Max(2, [Math]::Abs($yMin - $yMax))
                [double]$xL   = $xC - $barW / 2
                $fill = if ($avg -lt $LoN) { '#ffc8c8' } elseif ($avg -gt $HiN) { '#ffdcaa' } else { '#c0e8c0' }
                $scol = if ($avg -lt $LoN) { '#cc1111' } elseif ($avg -gt $HiN) { '#bb5500' } else { '#118811' }
                [void]$sb.Append("<rect x='$xL' y='$bTop' width='$barW' height='$bHt' fill='$fill' stroke='$scol' stroke-width='0.8' rx='1'/>")
                $avgPts.Add("$xC,$yAvg")
            }
        }

        # Linia srednich - wygladzona Catmull-Rom -> cubic Bezier
        if ($avgPts.Count -gt 1) {
            $na = $avgPts.Count
            $axArr = [double[]]::new($na)
            $ayArr = [double[]]::new($na)
            for ($i = 0; $i -lt $na; $i++) {
                $xy = $avgPts[$i] -split ','
                $axArr[$i] = [double]($xy[0]); $ayArr[$i] = [double]($xy[1])
            }
            [double]$ta = 0.3
            $apSb = New-Object System.Text.StringBuilder
            [void]$apSb.Append("M $($axArr[0]),$($ayArr[0])")
            for ($i = 0; $i -lt $na - 1; $i++) {
                $i0 = if ($i -gt 0) { $i - 1 } else { 0 }
                $i3 = if ($i + 2 -lt $na) { $i + 2 } else { $na - 1 }
                [double]$cp1x = [Math]::Round($axArr[$i]   + ($axArr[$i+1] - $axArr[$i0]) * $ta, 1)
                [double]$cp1y = [Math]::Round($ayArr[$i]   + ($ayArr[$i+1] - $ayArr[$i0]) * $ta, 1)
                [double]$cp2x = [Math]::Round($axArr[$i+1] - ($axArr[$i3]  - $axArr[$i])  * $ta, 1)
                [double]$cp2y = [Math]::Round($ayArr[$i+1] - ($ayArr[$i3]  - $ayArr[$i])  * $ta, 1)
                [void]$apSb.Append(" C $cp1x,$cp1y $cp2x,$cp2y $($axArr[$i+1]),$($ayArr[$i+1])")
            }
            [void]$sb.Append("<path d='$($apSb.ToString())' fill='none' stroke='#1144aa' stroke-width='1.8'/>")
            for ($i = 0; $i -lt $na; $i++) {
                [void]$sb.Append("<circle cx='$($axArr[$i])' cy='$($ayArr[$i])' r='2.5' fill='#1144aa'/>")
            }
        }

        # Osie
        [void]$sb.Append("<line x1='$pL' y1='$pT' x2='$pL' y2='$($pT+$gH)' stroke='#99a' stroke-width='1'/>")
        [void]$sb.Append("<line x1='$pL' y1='$($pT+$gH)' x2='$($pL+$gW)' y2='$($pT+$gH)' stroke='#99a' stroke-width='1'/>")
        [void]$sb.Append("<text x='$pL' y='$($pT-2)' font-family='Arial' font-size='9' fill='#333'>$Unit</text>")
        [void]$sb.Append("</svg>")
        return $sb.ToString()
    } catch { Write-Log "AgpSvg ERR: $($_.Exception.Message)"; return "" }
}

# Wykres słupkowy BARS (8 slotow 3h) - taki sam jak Show-AvgBarsWindow, do PDF
# UWAGA: wspolrzedne SVG musza uzywac kropki jako separatora dziesietnego (InvariantCulture).
# Bezposrednia interpolacja "$zmienna" w PowerShell uzywa InvariantCulture dla liczb - OK.
# Nigdy nie uzywac .ToString("0.0") dla wspolrzednych SVG - daje przecinek na pl-PL i lamie SVG.
function Get-BarsSvg {
    param([object[]]$Data, [double]$LoN, [double]$HiN, [bool]$UseMgDl, [string]$Unit, [bool]$LangEn)
    try {
        $ic = [System.Globalization.CultureInfo]::InvariantCulture
        $slotSums   = [double[]]::new(8)
        $slotCounts = [int[]]::new(8)
        foreach ($pt in $Data) {
            try {
                $ts   = [DateTime]::Parse($pt.ts)
                $slot = [int][Math]::Floor($ts.Hour / 3)
                $val  = if ($UseMgDl) { [double]$pt.mgdl } else { [double]$pt.mgdl / 18.018 }
                $slotSums[$slot]   += $val
                $slotCounts[$slot] += 1
            } catch {}
        }

        $avgs    = [double[]]::new(8)
        $hasData = [bool[]]::new(8)
        $allAvgs = [System.Collections.Generic.List[double]]::new()
        for ($s = 0; $s -lt 8; $s++) {
            if ($slotCounts[$s] -gt 0) {
                $avgs[$s]    = $slotSums[$s] / $slotCounts[$s]
                $hasData[$s] = $true
                $allAvgs.Add($avgs[$s])
            }
        }
        if ($allAvgs.Count -lt 2) { return "" }

        [double]$W = 760; [double]$H = 180
        [double]$pL = 44; [double]$pR = 16; [double]$pT = 14; [double]$pB = 30
        [double]$gW = $W - $pL - $pR; [double]$gH = $H - $pT - $pB

        $sortedAvgs = $allAvgs.ToArray() | Sort-Object
        [double]$dataMin = $sortedAvgs[0]; [double]$dataMax = $sortedAvgs[$sortedAvgs.Count-1]
        [double]$yLo = [Math]::Max(0, $dataMin * 0.75)
        [double]$yHi = $dataMax * 1.15
        if ($yHi -le $yLo) { $yHi = $yLo + 1 }
        [double]$yRng = $yHi - $yLo

        $fmt  = if ($UseMgDl) { "0" } else { "0.0" }   # format etykiet (locale OK - to tekst, nie wspolrzedna)
        [double]$barW = [Math]::Round($gW / 8 * 0.6, 1)

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 $W $H' style='width:100%;height:auto;display:block'>")
        [void]$sb.Append("<rect width='$W' height='$H' fill='#f8f9ff' rx='3'/>")

        # Strefa normy - wartosci LoN/HiN clampowane do widocznego zakresu Y
        [double]$nLoV = [Math]::Min([Math]::Max($LoN, $yLo), $yHi)
        [double]$nHiV = [Math]::Min([Math]::Max($HiN, $yLo), $yHi)
        [double]$nLoY = [Math]::Round($pT + $gH - ($nLoV - $yLo) / $yRng * $gH, 1)
        [double]$nHiY = [Math]::Round($pT + $gH - ($nHiV - $yLo) / $yRng * $gH, 1)
        [double]$nTop = [Math]::Min($nLoY, $nHiY)
        [double]$nHt  = [Math]::Abs($nLoY - $nHiY)
        # Bezposrednia interpolacja -> InvariantCulture -> kropka w SVG
        [void]$sb.Append("<rect x='$pL' y='$nTop' width='$gW' height='$nHt' fill='#d4edda'/>")

        # Linie poziome Y + etykiety osi
        [double]$step = if ($UseMgDl) { 50 } else { 2 }
        [double]$yv = [Math]::Ceiling($yLo / $step) * $step
        while ($yv -le $yHi) {
            [double]$yp = [Math]::Round($pT + $gH - ($yv - $yLo) / $yRng * $gH, 1)
            $lbl = if ($UseMgDl) { [int]$yv } else { $yv.ToString('F1', $ic) }
            [void]$sb.Append("<line x1='$pL' y1='$yp' x2='$($pL+$gW)' y2='$yp' stroke='#ccd' stroke-width='0.5' stroke-dasharray='3,2'/>")
            [void]$sb.Append("<text x='$($pL-3)' y='$($yp+3)' text-anchor='end' font-family='Arial' font-size='10' fill='#333'>$lbl</text>")
            $yv += $step
        }

        # Słupki + etykiety wartości
        for ($s = 0; $s -lt 8; $s++) {
            [double]$xC = [Math]::Round($pL + ($s + 0.5) / 8 * $gW, 1)
            $hLabel = "$($s * 3):00"
            [void]$sb.Append("<text x='$xC' y='$($pT+$gH+18)' text-anchor='middle' font-family='Arial' font-size='9' fill='#333'>$hLabel</text>")
            if (-not $hasData[$s]) { continue }

            [double]$avg  = $avgs[$s]
            [double]$frac = [Math]::Min([Math]::Max($avg - $yLo, 0), $yRng) / $yRng
            [double]$yAvg = [Math]::Round($pT + $gH - $frac * $gH, 1)
            [double]$barH = [Math]::Round([Math]::Max(4, $pT + $gH - $yAvg), 1)
            [double]$xL   = [Math]::Round($xC - $barW / 2, 1)

            $fill = if ($avg -lt $LoN) { '#ffc8c8' } elseif ($avg -gt $HiN) { '#ffdcaa' } else { '#a8dba8' }
            $scol = if ($avg -lt $LoN) { '#cc1111' } elseif ($avg -gt $HiN) { '#bb5500' } else { '#228822' }
            $tcol = if ($avg -lt $LoN) { '#cc1111' } elseif ($avg -gt $HiN) { '#bb5500' } else { '#1a6e1a' }

            # Bezposrednia interpolacja zmiennych [double] -> InvariantCulture -> kropka w SVG
            [void]$sb.Append("<rect x='$xL' y='$yAvg' width='$barW' height='$barH' fill='$fill' stroke='$scol' stroke-width='1' rx='3'/>")
            # Etykieta wartosci nad slupkiem (tekst dla uzytkownika - format locale jest OK)
            [double]$lblY = [Math]::Round([Math]::Max($yAvg - 6, $pT + 12), 1)
            $avgLbl = $avg.ToString($fmt)
            [void]$sb.Append("<text x='$xC' y='$lblY' text-anchor='middle' font-family='Arial' font-size='11' font-weight='bold' fill='$tcol'>$avgLbl</text>")
        }

        # Etykieta 24:00 + osie
        [double]$x24 = [Math]::Round($pL + $gW, 1)
        [void]$sb.Append("<text x='$x24' y='$($pT+$gH+18)' text-anchor='middle' font-family='Arial' font-size='9' fill='#333'>24:00</text>")
        [void]$sb.Append("<line x1='$pL' y1='$pT' x2='$pL' y2='$($pT+$gH)' stroke='#99a' stroke-width='1'/>")
        [void]$sb.Append("<line x1='$pL' y1='$($pT+$gH)' x2='$($pL+$gW)' y2='$($pT+$gH)' stroke='#99a' stroke-width='1'/>")
        [void]$sb.Append("<text x='$pL' y='$($pT-2)' font-family='Arial' font-size='9' fill='#333'>$Unit</text>")
        [void]$sb.Append("</svg>")
        return $sb.ToString()
    } catch { Write-Log "BarsSvg ERR: $($_.Exception.Message)"; return "" }
}

# Wykres wzorca dobowego AGP z pasmami percentylowymi (p10/p25/p50/p75/p90) - do PDF
# Na jasnym tle (odwrocona kolorystyka wzgledem WPF)
function Get-DailyPatternSvg {
    param([object[]]$Data, [double]$LoN, [double]$HiN, [bool]$UseMgDl, [string]$Unit, [bool]$LangEn, [int]$Days)
    try {
        $ic = [System.Globalization.CultureInfo]::InvariantCulture

        # Grupuj po slotach 5-min (0..287)
        # Uwaga: uzyj tablicy listow (nie Dictionary) - Dictionary zwraca kopie listy w PS
        $slotLists = [object[]]::new(288)
        for ($i = 0; $i -lt 288; $i++) { $slotLists[$i] = [System.Collections.Generic.List[double]]::new() }
        foreach ($pt in $Data) {
            try {
                $ts = [DateTime]::Parse($pt.ts)
                $si = $ts.Hour * 12 + [int]($ts.Minute / 5)
                [double]$mg = [double]$pt.mgdl
                if ($mg -le 20) { continue }
                $v = if ($UseMgDl) { $mg } else { [Math]::Round($mg / 18.018, 2) }
                $slotLists[$si].Add($v)
            } catch {}
        }

        # Percentyle dla kazdego slotu
        $hasS = [bool[]]::new(288)
        $sP10 = [double[]]::new(288); $sP25 = [double[]]::new(288)
        $sP50 = [double[]]::new(288); $sP75 = [double[]]::new(288); $sP90 = [double[]]::new(288)
        for ($s = 0; $s -lt 288; $s++) {
            if ($slotLists[$s].Count -ge 1) {
                $sv = $slotLists[$s].ToArray(); [Array]::Sort($sv)
                $sP10[$s] = Get-Percentile $sv 10; $sP25[$s] = Get-Percentile $sv 25
                $sP50[$s] = Get-Percentile $sv 50; $sP75[$s] = Get-Percentile $sv 75
                $sP90[$s] = Get-Percentile $sv 90; $hasS[$s] = $true
            }
        }
        $validCount = ($hasS | Where-Object { $_ }).Count
        if ($validCount -lt 12) { return "" }

        # Wygladz krzywe (ta sama logika co Render-HistGraph)
        $smW = if ($Days -le 7) { 30 } elseif ($Days -le 14) { 15 } elseif ($Days -le 30) { 8 } else { 4 }
        $sP10 = Smooth-Array $sP10 $hasS $smW; $sP25 = Smooth-Array $sP25 $hasS $smW
        $sP50 = Smooth-Array $sP50 $hasS $smW; $sP75 = Smooth-Array $sP75 $hasS $smW
        $sP90 = Smooth-Array $sP90 $hasS $smW

        # Wymiary SVG
        [double]$W = 760; [double]$H = 175
        [double]$pL = 44; [double]$pR = 8; [double]$pT = 12; [double]$pB = 26
        [double]$gW = $W - $pL - $pR; [double]$gH = $H - $pT - $pB
        [double]$yLo = if ($UseMgDl) { 40.0 } else { 2.2 }
        [double]$yHi = if ($UseMgDl) { 280.0 } else { 15.5 }
        [double]$yRng = $yHi - $yLo

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 $W $H' style='width:100%;height:auto;display:block'>")
        [void]$sb.Append("<rect width='$W' height='$H' fill='#f8f9ff' rx='3'/>")

        # Strefa normy
        [double]$nLoV = [Math]::Min([Math]::Max($LoN, $yLo), $yHi)
        [double]$nHiV = [Math]::Min([Math]::Max($HiN, $yLo), $yHi)
        [double]$nLoY = [Math]::Round($pT + $gH - ($nLoV - $yLo) / $yRng * $gH, 1)
        [double]$nHiY = [Math]::Round($pT + $gH - ($nHiV - $yLo) / $yRng * $gH, 1)
        [double]$nTop = [Math]::Min($nLoY, $nHiY); [double]$nHt = [Math]::Abs($nLoY - $nHiY)
        [void]$sb.Append("<rect x='$pL' y='$nTop' width='$gW' height='$nHt' fill='#d4edda'/>")

        # Siatka Y
        $gridVals = if ($UseMgDl) { @(70,100,140,180,250) } else { @(3.9,5.5,7.0,10.0,13.9) }
        foreach ($gVal in $gridVals) {
            if ($gVal -lt $yLo -or $gVal -gt $yHi) { continue }
            [double]$gy = [Math]::Round($pT + $gH - ($gVal - $yLo) / $yRng * $gH, 1)
            $lbl = if ($UseMgDl) { [int]$gVal } else { $gVal.ToString('F1', $ic) }
            [void]$sb.Append("<line x1='$pL' y1='$gy' x2='$($pL+$gW)' y2='$gy' stroke='#ccd' stroke-width='0.5' stroke-dasharray='3,2'/>")
            [void]$sb.Append("<text x='$($pL-3)' y='$($gy+3)' text-anchor='end' font-family='Arial' font-size='10' fill='#333'>$lbl</text>")
        }

        # Oblicz wspolrzedne pikseli (subsample co 3 sloty + slot 287)
        $xA   = [System.Collections.Generic.List[double]]::new()
        $y90L = [System.Collections.Generic.List[double]]::new(); $y10L = [System.Collections.Generic.List[double]]::new()
        $y75L = [System.Collections.Generic.List[double]]::new(); $y25L = [System.Collections.Generic.List[double]]::new()
        $y50L = [System.Collections.Generic.List[double]]::new()
        for ($s = 0; $s -lt 287; $s += 3) {
            if (-not $hasS[$s]) { continue }
            [double]$xs   = [Math]::Round($pL + $s / 287.0 * $gW, 1)
            [double]$y90v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP90[$s],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            [double]$y75v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP75[$s],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            [double]$y50v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP50[$s],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            [double]$y25v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP25[$s],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            [double]$y10v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP10[$s],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            $xA.Add($xs); $y90L.Add($y90v); $y75L.Add($y75v); $y50L.Add($y50v); $y25L.Add($y25v); $y10L.Add($y10v)
        }
        if ($hasS[287]) {
            [double]$xs   = [Math]::Round($pL + $gW, 1)
            [double]$y90v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP90[287],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            [double]$y75v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP75[287],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            [double]$y50v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP50[287],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            [double]$y25v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP25[287],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            [double]$y10v = [Math]::Round($pT + $gH - ([Math]::Min([Math]::Max($sP10[287],$yLo),$yHi) - $yLo) / $yRng * $gH, 1)
            $xA.Add($xs); $y90L.Add($y90v); $y75L.Add($y75v); $y50L.Add($y50v); $y25L.Add($y25v); $y10L.Add($y10v)
        }
        $nPts = $xA.Count; if ($nPts -lt 3) { return "" }

        # Pasmo p10-p90 (zewnetrzne, jasne)
        $polyO = New-Object System.Text.StringBuilder
        for ($i = 0; $i -lt $nPts; $i++) { [void]$polyO.Append("$($xA[$i]),$($y90L[$i]) ") }
        for ($i = $nPts-1; $i -ge 0; $i--) { [void]$polyO.Append("$($xA[$i]),$($y10L[$i]) ") }
        [void]$sb.Append("<polygon points='$($polyO.ToString().Trim())' fill='#7aaac8' fill-opacity='0.28' stroke='none'/>")

        # Pasmo p25-p75 IQR (wewnetrzne, ciemniejsze)
        $polyI = New-Object System.Text.StringBuilder
        for ($i = 0; $i -lt $nPts; $i++) { [void]$polyI.Append("$($xA[$i]),$($y75L[$i]) ") }
        for ($i = $nPts-1; $i -ge 0; $i--) { [void]$polyI.Append("$($xA[$i]),$($y25L[$i]) ") }
        [void]$sb.Append("<polygon points='$($polyI.ToString().Trim())' fill='#4477aa' fill-opacity='0.45' stroke='none'/>")

        # Mediana - Catmull-Rom -> cubic Bezier
        $xAr = $xA.ToArray(); $yAr = $y50L.ToArray(); [double]$tens = 0.35
        $pathSb = New-Object System.Text.StringBuilder
        [void]$pathSb.Append("M $($xAr[0]),$($yAr[0])")
        for ($i = 0; $i -lt $nPts-1; $i++) {
            $i0 = if ($i -gt 0) { $i-1 } else { 0 }; $i3 = if ($i+2 -lt $nPts) { $i+2 } else { $nPts-1 }
            [double]$cp1x = [Math]::Round($xAr[$i]   + ($xAr[$i+1] - $xAr[$i0]) * $tens, 1)
            [double]$cp1y = [Math]::Round($yAr[$i]   + ($yAr[$i+1] - $yAr[$i0]) * $tens, 1)
            [double]$cp2x = [Math]::Round($xAr[$i+1] - ($xAr[$i3]  - $xAr[$i])  * $tens, 1)
            [double]$cp2y = [Math]::Round($yAr[$i+1] - ($yAr[$i3]  - $yAr[$i])  * $tens, 1)
            [void]$pathSb.Append(" C $cp1x,$cp1y $cp2x,$cp2y $($xAr[$i+1]),$($yAr[$i+1])")
        }
        [void]$sb.Append("<path d='$($pathSb.ToString())' fill='none' stroke='#1a3a6e' stroke-width='2.5' stroke-linecap='round'/>")

        # Etykiety osi X co 3h
        for ($h = 0; $h -le 24; $h += 3) {
            $hh = $h % 24
            [double]$xPos = if ($h -eq 24) { $pL + $gW } else { [Math]::Round($pL + ($hh * 12) / 287.0 * $gW, 1) }
            [void]$sb.Append("<text x='$xPos' y='$($pT+$gH+15)' text-anchor='middle' font-family='Arial' font-size='9' fill='#333'>$($hh.ToString('00')):00</text>")
        }

        # Osie
        [void]$sb.Append("<line x1='$pL' y1='$pT' x2='$pL' y2='$($pT+$gH)' stroke='#99a' stroke-width='1'/>")
        [void]$sb.Append("<line x1='$pL' y1='$($pT+$gH)' x2='$($pL+$gW)' y2='$($pT+$gH)' stroke='#99a' stroke-width='1'/>")
        [void]$sb.Append("<text x='$pL' y='$($pT-2)' font-family='Arial' font-size='9' fill='#333'>$Unit</text>")
        [void]$sb.Append("</svg>")
        $result = $sb.ToString()
        return $result
    } catch { Write-Log "DailyPatternSvg ERR: $($_.Exception.Message) | $($_.ScriptStackTrace)"; return "" }
}

# ======================== DOCTOR SUMMARY ========================
function Get-DoctorSummary {
    param([object[]]$Data, [double]$LoN, [double]$HiN, [bool]$UseMgDl, [string]$Unit, [bool]$LangEn, [int]$Days)
    try {
        if ($Data.Count -lt 10) { return "" }
        $ic = [System.Globalization.CultureInfo]::InvariantCulture
        $dp = if ($UseMgDl) { 0 } else { 1 }
        $fmt = if ($UseMgDl) { "0" } else { "0.0" }

        # Zbierz wartosci + grupuj po godzinach
        $vals = [System.Collections.Generic.List[double]]::new()
        $byHour = [object[]]::new(24)
        for ($i = 0; $i -lt 24; $i++) { $byHour[$i] = [System.Collections.Generic.List[double]]::new() }
        foreach ($pt in $Data) {
            try {
                [double]$mg = [double]$pt.mgdl
                if ($mg -le 20) { continue }
                [double]$v = if ($UseMgDl) { $mg } else { $mg / 18.018 }
                $vals.Add($v)
                $byHour[[DateTime]::Parse($pt.ts).Hour].Add($v)
            } catch {}
        }
        if ($vals.Count -lt 10) { return "" }

        # Statystyki (single-pass)
        $vArr = $vals.ToArray(); $nV = $vArr.Count
        $sum = 0.0; $sumSq = 0.0
        $tirN = 0; $tbrN = 0; $tarN = 0; $tbrVN = 0; $tarVN = 0
        $vLoThr = if ($UseMgDl) { 54.0 } else { 3.0 }
        $vHiThr = if ($UseMgDl) { 250.0 } else { 13.9 }
        foreach ($v in $vArr) {
            $sum += $v; $sumSq += $v * $v
            if ($v -ge $LoN -and $v -le $HiN) { $tirN++ }
            elseif ($v -lt $LoN) { $tbrN++; if ($v -lt $vLoThr) { $tbrVN++ } }
            else { $tarN++; if ($v -gt $vHiThr) { $tarVN++ } }
        }
        $avg = $sum / $nV
        $sd  = [Math]::Sqrt([Math]::Max(0, $sumSq / $nV - $avg * $avg))
        $cv  = [Math]::Round($sd / $avg * 100, 1)
        [double]$avgMgDl = if ($UseMgDl) { $avg } else { $avg * 18.018 }
        [double]$eHbA1c  = [Math]::Round(($avgMgDl + 46.7) / 28.7, 1)
        $tirP  = [Math]::Round($tirN / $nV * 100, 1)
        $tbrP  = [Math]::Round($tbrN / $nV * 100, 1)
        $tarP  = [Math]::Round($tarN / $nV * 100, 1)
        $tbrVP = [Math]::Round($tbrVN / $nV * 100, 1)
        $tarVP = [Math]::Round($tarVN / $nV * 100, 1)
        $avgD  = [Math]::Round($avg, $dp)
        $sdD   = [Math]::Round($sd,  $dp)

        # Srednie czasowe
        function HourAvg([object[]]$lists, [int[]]$hours) {
            $s = 0.0; $c = 0
            foreach ($h in $hours) { foreach ($x in $lists[$h]) { $s += $x; $c++ } }
            if ($c -gt 3) { return [Math]::Round($s / $c, 1) } else { return $null }
        }
        $fastD = HourAvg $byHour @(5,6,7)
        $noctD = HourAvg $byHour @(0,1,2,3,4)
        $ppD   = HourAvg $byHour @(8,9,12,13,18,19)

        $sb = New-Object System.Text.StringBuilder

        if ($LangEn) {
            # --- ENGLISH ---
            $ctrl = if ($tirP -ge 70) { "good" } elseif ($tirP -ge 50) { "moderate, requiring optimization" } else { "suboptimal, requiring significant therapeutic review" }
            $hbC  = if ($eHbA1c -lt 7.0) { "within the recommended target (&lt;7.0%)" } `
                    elseif ($eHbA1c -lt 8.0) { "slightly above the recommended target" } `
                    elseif ($eHbA1c -lt 9.0) { "above target &mdash;requires attention" } `
                    else { "significantly above target &mdash;immediate therapeutic review recommended" }

            $p1 = "The analysis covers <strong>$Days days</strong> of continuous glucose monitoring comprising <strong>$nV readings</strong>. Mean glucose was <strong>$avgD $Unit</strong>, corresponding to an estimated HbA1c of <strong>$($eHbA1c.ToString('0.0',$ic))%</strong>, which is $hbC. Overall glycemic control is <strong>$ctrl</strong> with a Time in Range (TIR) of <strong>$tirP%</strong> (recommended target &ge;70%)."
            [void]$sb.Append("<p style='margin:0 0 7px 0'>$p1</p>")

            $hypo = if ($tbrP -lt 1) {
                "Hypoglycemia was negligible ($tbrP% of readings below target range)."
            } elseif ($tbrP -lt 4) {
                "Minor hypoglycemia was observed ($tbrP% of readings below target; recommended &lt;4%). Further reduction is advisable."
            } else {
                "Clinically significant hypoglycemia was recorded: <strong>$tbrP%</strong> of time below target$(if($tbrVP -gt 1){", including <strong>$tbrVP%</strong> in the very low range (&lt;$($vLoThr.ToString($fmt,$ic)) $Unit)"})&nbsp;&mdash; urgent therapeutic adjustment recommended."
            }
            $hyper = if ($tarP -lt 10) {
                "Hyperglycemia was minimal ($tarP% above target)."
            } elseif ($tarP -lt 25) {
                "Moderate hyperglycemia was present (<strong>$tarP%</strong> of time above target); therapy optimisation is advisable."
            } else {
                "Significant hyperglycemia was recorded: <strong>$tarP%</strong> of time above target$(if($tarVP -gt 5){", including <strong>$tarVP%</strong> in the very high range (&gt;$($vHiThr.ToString($fmt,$ic)) $Unit)"})&nbsp;&mdash; intensification of treatment should be considered."
            }
            [void]$sb.Append("<p style='margin:0 0 7px 0'>$hypo $hyper</p>")

            $cvA = if ($cv -lt 27) { "excellent glycemic stability (CV&nbsp;=&nbsp;$cv%; target &lt;36%)" } `
                   elseif ($cv -lt 36) { "acceptable glycemic stability (CV&nbsp;=&nbsp;$cv%; target &lt;36%)" } `
                   else { "elevated glycemic variability (CV&nbsp;=&nbsp;$cv%; target &lt;36%) &mdash;this independently increases cardiovascular risk and may reflect undetected hypoglycemia" }
            $p3 = "Standard deviation was <strong>$sdD $Unit</strong>. The profile demonstrates $cvA."
            if ($null -ne $fastD) { $p3 += " Mean fasting glucose (05:00&ndash;08:00): <strong>$fastD $Unit</strong>." }
            if ($null -ne $noctD) {
                $noctCmt = if ($noctD -lt $LoN) { " &mdash;<em>nocturnal hypoglycemia risk</em>" } elseif ($noctD -gt $HiN) { " &mdash;<em>nocturnal hyperglycemia</em>" } else { " &mdash;within target" }
                $p3 += " Nocturnal average (00:00&ndash;05:00): <strong>$noctD $Unit</strong>$noctCmt."
            }
            if ($null -ne $ppD) { $p3 += " Postprandial average: <strong>$ppD $Unit</strong>." }
            [void]$sb.Append("<p style='margin:0 0 7px 0'>$p3</p>")

            # Recommendations
            $recs = [System.Collections.Generic.List[string]]::new()
            if ($tirP -lt 70) { $recs.Add("Review therapy to improve Time in Range (current $tirP% vs target &ge;70%)") }
            if ($tbrP -ge 4)  { $recs.Add("Investigate hypoglycemia &mdash;consider adjusting insulin doses/basal rates (TBR $tbrP%, target &lt;4%)") }
            if ($tbrVP -gt 1) { $recs.Add("Very low glucose episodes recorded ($tbrVP%) &mdash;urgent hypoglycemia prevention strategy review") }
            if ($tarP -ge 25) { $recs.Add("Significant hyperglycemia ($tarP% TAR) &mdash;consider treatment intensification") }
            if ($cv -ge 36)   { $recs.Add("High variability (CV $cv%) &mdash;review meal composition, activity patterns and insulin timing") }
            if ($null -ne $noctD -and $noctD -lt $LoN) { $recs.Add("Nocturnal hypoglycemia risk &mdash;review bedtime insulin and pre-sleep carbohydrate intake") }
            if ($null -ne $noctD -and $noctD -gt $HiN) { $recs.Add("Nocturnal hyperglycemia &mdash;consider adjusting basal insulin or evening meal management") }
            if ($recs.Count -eq 0) { $recs.Add("Continue current therapy &mdash;glycemic parameters are within or close to recommended targets") }
            [void]$sb.Append("<p style='margin:0 0 4px 0'><strong>Clinical Recommendations:</strong></p>")
            [void]$sb.Append("<ul style='margin:0 0 8px 20px;padding:0'>")
            foreach ($r in $recs) { [void]$sb.Append("<li style='margin-bottom:3px'>$r</li>") }
            [void]$sb.Append("</ul>")
            [void]$sb.Append("<p style='font-size:11px;color:#999;font-style:italic;margin:8px 0 0 0'>This analysis was generated automatically by Glucose Monitor software based on CGM data. It is intended as a clinical decision-support tool and does not replace physician judgment.</p>")

        } else {
            # --- POLISH (wszystkie polskie znaki jako HTML entities - bezpieczne dla kazdego kodowania pliku) ---
            $ctrl = if ($tirP -ge 70) { "dobra" } elseif ($tirP -ge 50) { "umiarkowana, wymagaj&#261;ca optymalizacji" } else { "niewystarczaj&#261;ca, wymagaj&#261;ca weryfikacji leczenia" }
            $hbC  = if ($eHbA1c -lt 7.0) { "w granicach zalecanego celu terapeutycznego (&lt;7,0%)" } `
                    elseif ($eHbA1c -lt 8.0) { "nieznacznie powy&#380;ej zalecanego celu" } `
                    elseif ($eHbA1c -lt 9.0) { "powy&#380;ej celu terapeutycznego &mdash; wymaga uwagi" } `
                    else { "znacznie powy&#380;ej celu &mdash; zalecana pilna weryfikacja leczenia" }

            $p1 = "Analiza obejmuje <strong>$Days dni</strong> ci&#261;g&#322;ego monitorowania glikemii, &#322;&#261;cznie <strong>$nV odczyt&#243;w</strong>. &#346;rednia glikemia wynios&#322;a <strong>$avgD $Unit</strong>, co odpowiada szacowanemu HbA1c <strong>$($eHbA1c.ToString('0.0',$ic))%</strong> &mdash; $hbC. Og&#243;lna kontrola glikemii jest <strong>$ctrl</strong>; czas w zakresie docelowym (TIR) wyni&#243;s&#322; <strong>$tirP%</strong> (zalecany cel &ge;70%)."
            [void]$sb.Append("<p style='margin:0 0 7px 0'>$p1</p>")

            $hypo = if ($tbrP -lt 1) {
                "Hipoglikemia by&#322;a nieistotna ($tbrP% odczyt&#243;w poni&#380;ej zakresu docelowego)."
            } elseif ($tbrP -lt 4) {
                "Odnotowano nieznaczn&#261; hipoglikemi&#281; ($tbrP% poni&#380;ej zakresu; zalecany cel &lt;4%). Wskazana dalsza redukcja."
            } else {
                "Odnotowano klinicznie istotn&#261; hipoglikemi&#281;: <strong>$tbrP%</strong> czasu poni&#380;ej zakresu$(if($tbrVP -gt 1){", w tym <strong>$tbrVP%</strong> w zakresie bardzo niskim (&lt;$($vLoThr.ToString($fmt,$ic)) $Unit)"})&nbsp;&mdash; wymagana pilna weryfikacja leczenia."
            }
            $hyper = if ($tarP -lt 10) {
                "Hiperglikemia by&#322;a minimalna ($tarP% powy&#380;ej zakresu docelowego)."
            } elseif ($tarP -lt 25) {
                "Odnotowano umiarkown&#261; hiperglikemi&#281; (<strong>$tarP%</strong> czasu powy&#380;ej zakresu); wskazana optymalizacja terapii."
            } else {
                "Stwierdzono istotn&#261; hiperglikemi&#281;: <strong>$tarP%</strong> czasu powy&#380;ej zakresu$(if($tarVP -gt 5){", w tym <strong>$tarVP%</strong> w zakresie bardzo wysokim (&gt;$($vHiThr.ToString($fmt,$ic)) $Unit)"})&nbsp;&mdash; nale&#380;y rozwa&#380;y&#263; intensyfikacj&#281; leczenia."
            }
            [void]$sb.Append("<p style='margin:0 0 7px 0'>$hypo $hyper</p>")

            $cvA = if ($cv -lt 27) { "bardzo dobr&#261; stabilno&#347;ci&#261; glikemii (CV&nbsp;=&nbsp;$cv%; cel &lt;36%)" } `
                   elseif ($cv -lt 36) { "akceptowaln&#261; stabilno&#347;ci&#261; glikemii (CV&nbsp;=&nbsp;$cv%; cel &lt;36%)" } `
                   else { "podwy&#380;szon&#261; zmienno&#347;ci&#261; glikemii (CV&nbsp;=&nbsp;$cv%; cel &lt;36%) &mdash; niezale&#380;nie zwi&#281;ksza ryzyko sercowo-naczyniowe i mo&#380;e wi&#261;za&#263; si&#281; z nierozpoznanymi epizodami hipoglikemii" }
            $p3 = "Odchylenie standardowe wynios&#322;o <strong>$sdD $Unit</strong>. Profil cechowa&#322; si&#281; $cvA."
            if ($null -ne $fastD) { $p3 += " &#346;rednia glikemia na czczo (05:00&ndash;08:00): <strong>$fastD $Unit</strong>." }
            if ($null -ne $noctD) {
                $noctCmt = if ($noctD -lt $LoN) { " &mdash; <em>ryzyko hipoglikemii nocnej</em>" } elseif ($noctD -gt $HiN) { " &mdash; <em>hiperglikemia nocna</em>" } else { " &mdash; w zakresie docelowym" }
                $p3 += " &#346;rednia glikemia nocna (00:00&ndash;05:00): <strong>$noctD $Unit</strong>$noctCmt."
            }
            if ($null -ne $ppD) { $p3 += " &#346;rednia poposi&#322;kowa: <strong>$ppD $Unit</strong>." }
            [void]$sb.Append("<p style='margin:0 0 7px 0'>$p3</p>")

            # Zalecenia
            $recs = [System.Collections.Generic.List[string]]::new()
            if ($tirP -lt 70) { $recs.Add("Weryfikacja i optymalizacja leczenia w celu poprawy TIR (aktualnie $tirP%, cel &ge;70%)") }
            if ($tbrP -ge 4)  { $recs.Add("Analiza i redukcja ryzyka hipoglikemii &mdash; rozwa&#380;enie korekty dawek insuliny/dawki podstawowej (TBR $tbrP%, cel &lt;4%)") }
            if ($tbrVP -gt 1) { $recs.Add("Zarejestrowano epizody bardzo niskiej glikemii ($tbrVP%) &mdash; pilna weryfikacja strategii prewencji hipoglikemii") }
            if ($tarP -ge 25) { $recs.Add("Istotna hiperglikemia ($tarP% TAR) &mdash; rozwa&#380;enie intensyfikacji leczenia") }
            if ($cv -ge 36)   { $recs.Add("Wysoka zmienno&#347;&#263; glikemii (CV $cv%) &mdash; analiza sk&#322;adu posi&#322;k&#243;w, aktywno&#347;ci fizycznej i harmonogramu podawania insuliny") }
            if ($null -ne $noctD -and $noctD -lt $LoN) { $recs.Add("Ryzyko hipoglikemii nocnej &mdash; weryfikacja dawki insuliny podstawowej i w&#281;glowodan&#243;w przed snem") }
            if ($null -ne $noctD -and $noctD -gt $HiN) { $recs.Add("Hiperglikemia nocna &mdash; rozwa&#380;enie korekty dawki insuliny podstawowej lub zarz&#261;dzania kolacj&#261;") }
            if ($recs.Count -eq 0) { $recs.Add("Kontynuacja aktualnego leczenia &mdash; parametry glikemii mieszcz&#261; si&#281; w granicach zalecanych lub s&#261; bliskie celom terapeutycznym") }
            [void]$sb.Append("<p style='margin:0 0 4px 0'><strong>Zalecenia kliniczne:</strong></p>")
            [void]$sb.Append("<ul style='margin:0 0 8px 20px;padding:0'>")
            foreach ($r in $recs) { [void]$sb.Append("<li style='margin-bottom:3px'>$r</li>") }
            [void]$sb.Append("</ul>")
            [void]$sb.Append("<p style='font-size:11px;color:#999;font-style:italic;margin:8px 0 0 0'>Niniejsza analiza zosta&#322;a wygenerowana automatycznie przez oprogramowanie Glucose Monitor na podstawie danych CGM. Ma charakter pomocniczy i nie zast&#281;puje oceny klinicznej lekarza.</p>")
        }

        return $sb.ToString()
    } catch { Write-Log "DoctorSummary ERR: $($_.Exception.Message)"; return "" }
}

# ======================== PDF EXPORT ========================
function Export-PdfReport {
    $days = if ($script:HistWin -and $script:HistWin.IsLoaded) { $script:HistDays } else { 14 }
    $data = Load-HistoryData $days
    if ($data.Count -lt 2) {
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) { "Not enough data for report." } else { "Za malo danych do raportu." }),
            "Glucose Monitor") | Out-Null; return
    }

    # --- Dialog: imie, nazwisko, jezyk raportu ---
    $script:PdfLangEn = $script:LangEn  # domyslnie: jezyk UI
    $xamlN = [xml]@"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Eksport PDF" Width="320" Height="242"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
  <Border Background="#111827" CornerRadius="10" BorderBrush="#334466" BorderThickness="1">
    <StackPanel Margin="20,16,20,16">
      <TextBlock Text="$(if ($script:LangEn){'Patient data for PDF'}else{'Dane pacjenta do PDF'})"
                 Foreground="White" FontSize="13" FontFamily="Segoe UI Semibold" Margin="0,0,0,12"/>
      <TextBox Name="tbFirst" Tag="$(if ($script:LangEn){'First name'}else{'Imie'})"
               Background="#1a2540" Foreground="White" CaretBrush="White"
               BorderBrush="#334466" BorderThickness="1" Padding="8,6"
               FontSize="13" Margin="0,0,0,8"/>
      <TextBox Name="tbLast" Tag="$(if ($script:LangEn){'Last name'}else{'Nazwisko'})"
               Background="#1a2540" Foreground="White" CaretBrush="White"
               BorderBrush="#334466" BorderThickness="1" Padding="8,6"
               FontSize="13" Margin="0,0,0,12"/>
      <Grid Margin="0,0,0,12">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="$(if ($script:LangEn){'Report language:'}else{'Jezyk raportu:'})"
                   Foreground="#7777aa" FontSize="11" VerticalAlignment="Center" Margin="0,0,8,0"/>
        <Border Grid.Column="2" CornerRadius="5" BorderBrush="#334466" BorderThickness="1"
                Margin="0,0,0,0" Height="28">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="38"/><ColumnDefinition Width="38"/></Grid.ColumnDefinitions>
            <Button Grid.Column="0" Name="btnLangPl" Content="PL"
                    Background="#2244aa" Foreground="White" BorderThickness="0"
                    FontSize="11" FontWeight="SemiBold" Cursor="Hand"/>
            <Button Grid.Column="1" Name="btnLangEn" Content="EN"
                    Background="Transparent" Foreground="#556688" BorderThickness="0"
                    FontSize="11" Cursor="Hand"/>
          </Grid>
        </Border>
      </Grid>
      <Grid>
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <Button Grid.Column="0" Name="btnNOk" Content="$(if ($script:LangEn){'Export'}else{'Eksportuj'})"
                Margin="0,0,4,0" Padding="0,8" Background="#2244aa" Foreground="White"
                BorderThickness="0" Cursor="Hand" FontSize="12"/>
        <Button Grid.Column="1" Name="btnNCancel" Content="$(if ($script:LangEn){'Cancel'}else{'Anuluj'})"
                Margin="4,0,0,0" Padding="0,8" Background="#1a2540" Foreground="#7777aa"
                BorderThickness="0" Cursor="Hand" FontSize="12"/>
      </Grid>
    </StackPanel>
  </Border>
</Window>
"@
    try {
        $nd   = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xamlN))
        $tbF  = $nd.FindName("tbFirst"); $tbL = $nd.FindName("tbLast")
        $bPl  = $nd.FindName("btnLangPl"); $bEn = $nd.FindName("btnLangEn")

        # Placeholder (hint)
        $tbF.Text = $tbF.Tag; $tbF.Foreground = [System.Windows.Media.Brushes]::Gray
        $tbL.Text = $tbL.Tag; $tbL.Foreground = [System.Windows.Media.Brushes]::Gray
        $tbF.Add_GotFocus({  if ($tbF.Foreground -eq [System.Windows.Media.Brushes]::Gray) { $tbF.Text = ""; $tbF.Foreground = [System.Windows.Media.Brushes]::White } })
        $tbF.Add_LostFocus({ if ($tbF.Text -eq "") { $tbF.Text = $tbF.Tag; $tbF.Foreground = [System.Windows.Media.Brushes]::Gray } })
        $tbL.Add_GotFocus({  if ($tbL.Foreground -eq [System.Windows.Media.Brushes]::Gray) { $tbL.Text = ""; $tbL.Foreground = [System.Windows.Media.Brushes]::White } })
        $tbL.Add_LostFocus({ if ($tbL.Text -eq "") { $tbL.Text = $tbL.Tag; $tbL.Foreground = [System.Windows.Media.Brushes]::Gray } })

        # Funkcja ustawiajaca aktywny jezyk (lokalna przez script scope)
        function Set-PdfLang([bool]$en) {
            $script:PdfLangEn = $en
            if ($en) {
                $bEn.Background  = [System.Windows.Media.Brushes]::RoyalBlue
                $bEn.Foreground  = [System.Windows.Media.Brushes]::White
                $bEn.FontWeight  = [System.Windows.FontWeights]::SemiBold
                $bPl.Background  = [System.Windows.Media.Brushes]::Transparent
                $bPl.Foreground  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#556688"))
                $bPl.FontWeight  = [System.Windows.FontWeights]::Normal
            } else {
                $bPl.Background  = [System.Windows.Media.Brushes]::RoyalBlue
                $bPl.Foreground  = [System.Windows.Media.Brushes]::White
                $bPl.FontWeight  = [System.Windows.FontWeights]::SemiBold
                $bEn.Background  = [System.Windows.Media.Brushes]::Transparent
                $bEn.Foreground  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#556688"))
                $bEn.FontWeight  = [System.Windows.FontWeights]::Normal
            }
        }

        # Ustaw stan poczatkowy (pamietaj ostatni wybor)
        Set-PdfLang $script:PdfLangEn

        $bPl.Add_Click({ Set-PdfLang $false })
        $bEn.Add_Click({ Set-PdfLang $true  })

        $script:PdfOk = $false
        $nd.FindName("btnNOk").Add_Click({ $script:PdfOk = $true; $nd.Close() })
        $nd.FindName("btnNCancel").Add_Click({ $nd.Close() })
        $nd.Add_MouseLeftButtonDown({ $nd.DragMove() })
        $nd.ShowDialog() | Out-Null
        if (-not $script:PdfOk) { return }
        $firstName = if ($tbF.Foreground -ne [System.Windows.Media.Brushes]::Gray) { $tbF.Text.Trim() } else { "" }
        $lastName  = if ($tbL.Foreground -ne [System.Windows.Media.Brushes]::Gray) { $tbL.Text.Trim() } else { "" }
    } catch { Write-Log "PDF Name ERR: $($_.Exception.Message)"; return }

    # --- Oblicz statystyki ---
    try {
        $vals  = $data | ForEach-Object { if ($script:UseMgDl) { [double]$_.mgdl } else { [Math]::Round([double]$_.mgdl/18.018,1) } }
        $sm    = $vals | Measure-Object -Min -Max -Average
        $avg   = [Math]::Round($sm.Average, 1)
        $loN   = if ($script:UseMgDl) { 70.0  } else { 3.9  }
        $hiN   = if ($script:UseMgDl) { 180.0 } else { 10.0 }
        $tir   = [Math]::Round(($vals | Where-Object { $_ -ge $loN -and $_ -le $hiN }).Count / $vals.Count * 100, 0)
        $mean  = $sm.Average
        $sd    = [Math]::Round([Math]::Sqrt(($vals | ForEach-Object { ($_ - $mean)*($_ - $mean) } | Measure-Object -Sum).Sum / $vals.Count), 1)
        $cv    = if ($mean -gt 0) { [Math]::Round($sd/$mean*100, 0) } else { 0 }
        $avgMg = if ($script:UseMgDl) { $mean } else { $mean * 18.018 }
        $hba1c = [Math]::Round(($avgMg + 46.7) / 28.7, 1)
        $unit  = if ($script:UseMgDl) { "mg/dL" } else { "mmol/L" }
        $fmt   = if ($script:UseMgDl) { "0" } else { "0.0" }
        $dateFrom = (Get-Date).AddDays(-$days).ToString("dd.MM.yyyy")
        $dateTo   = (Get-Date).ToString("dd.MM.yyyy")
        $lEn      = $script:PdfLangEn   # jezyk wybrany w dialogu, nie jezyk UI
        $patFull  = "$firstName $lastName".Trim()
        if (-not $patFull) { $patFull = "---" }

        # Generuj wykresy SVG
        $svgHist    = Get-HistSvg          -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn
        $svgAgp     = Get-AgpSvg           -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn
        $svgBars    = Get-BarsSvg          -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn
        $svgPattern  = Get-DailyPatternSvg  -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn -Days $days
        $doctorSummary = Get-DoctorSummary -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn -Days $days
        $secTitle = if ($lEn) { "Clinical Analysis" } else { "Analiza kliniczna" }
        $doctorSection = if ($doctorSummary) {
            "<div style='margin-top:18px;border-top:2px solid #1a1a4a;padding-top:12px'>" +
            "<div class='section-title'>$secTitle</div>" +
            "<div style='font-size:13px;color:#1a1a2e;line-height:1.8;background:#f8f9ff;border:1px solid #dde2ff;border-radius:6px;padding:14px 18px'>" +
            $doctorSummary + "</div></div>"
        } else { "" }

        $html = @"
<!DOCTYPE html>
<html lang="$(if($lEn){'en'}else{'pl'})">
<head>
<meta charset="UTF-8">
<style>
  @page { margin: 14mm 12mm 14mm 12mm; size: A4; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; color: #1a1a2e; background: #fff; }

  .header { display: flex; justify-content: space-between; align-items: flex-start;
            border-bottom: 2px solid #1a1a4a; padding-bottom: 10px; margin-bottom: 14px; }
  .header-left h1 { font-size: 22px; color: #1a1a4a; font-weight: 700; }
  .header-left .sub { font-size: 12px; color: #667; margin-top: 2px; }
  .header-right { font-size: 11px; color: #667; text-align: right; line-height: 1.6; }

  .patient-box { background: #f0f2ff; border-left: 4px solid #3344aa;
                 padding: 9px 14px; margin-bottom: 14px; border-radius: 0 6px 6px 0; }
  .patient-name { font-size: 17px; font-weight: 700; color: #1a1a4a; }
  .patient-meta { font-size: 12px; color: #556; margin-top: 3px; }

  .grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-bottom: 16px; }
  .card { border: 1px solid #dde2ff; border-radius: 6px; padding: 9px 6px; text-align: center; }
  .lbl { font-size: 10px; color: #889; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
  .val { font-size: 18px; font-weight: 700; line-height: 1; }
  .unit { font-size: 11px; font-weight: 400; color: #667; }
  .c-green  { color: #0a7a0a; }
  .c-blue   { color: #1144aa; }
  .c-orange { color: #bb5500; }
  .c-purple { color: #6622aa; }
  .c-teal   { color: #006677; }
  .c-gray   { color: #445566; }

  .section-title { font-size: 11px; font-weight: 700; color: #3344aa;
                   text-transform: uppercase; letter-spacing: 0.5px;
                   border-bottom: 1px solid #dde2ff; padding-bottom: 4px; margin-bottom: 8px; }

  .footer { margin-top: 14px; font-size: 10px; color: #aaa;
            text-align: center; border-top: 1px solid #eee; padding-top: 6px; }
  .range-legend { font-size: 11px; color: #667; margin-bottom: 12px; }
  .range-legend span { margin-right: 14px; }
  .dot-green  { display: inline-block; width: 8px; height: 8px; background: #0a7a0a; border-radius: 50%; margin-right: 3px; }
  .dot-red    { display: inline-block; width: 8px; height: 8px; background: #cc1111; border-radius: 50%; margin-right: 3px; }
  .dot-orange { display: inline-block; width: 8px; height: 8px; background: #bb5500; border-radius: 50%; margin-right: 3px; }
  .chart-box  { margin-bottom: 16px; border: 1px solid #dde2ff; border-radius: 6px; padding: 10px 8px 6px; background: #f8f9ff; }
  .chart-legend { font-size: 10px; color: #778; margin-top: 5px; }
  .chart-legend span { margin-right: 12px; }
  .leg-line  { display: inline-block; width: 18px; height: 2px; background: #3366cc; vertical-align: middle; margin-right: 3px; border-radius: 1px; }
  .leg-avg   { display: inline-block; width: 18px; height: 2px; background: #1144aa; vertical-align: middle; margin-right: 3px; border-radius: 1px; }
  .leg-green { display: inline-block; width: 10px; height: 10px; background: #d4edda; border: 1px solid #118811; border-radius: 2px; vertical-align: middle; margin-right: 3px; }
  .leg-red   { display: inline-block; width: 8px;  height: 8px;  background: #cc1111; border-radius: 50%; vertical-align: middle; margin-right: 3px; }
  .leg-or    { display: inline-block; width: 8px;  height: 8px;  background: #bb5500; border-radius: 50%; vertical-align: middle; margin-right: 3px; }
  @media print { .chart-box { break-inside: avoid; } }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Glucose Monitor</h1>
    <div class="sub">$(if($lEn){'Blood glucose report'}else{'Raport glikemii'})</div>
  </div>
  <div class="header-right">
    <div>$(if($lEn){'Generated'}else{'Wygenerowano'}): $(Get-Date -Format 'dd.MM.yyyy HH:mm')</div>
    <div>$(if($lEn){'Period'}else{'Okres'}): $dateFrom &ndash; $dateTo</div>
    <div>$days $(if($lEn){'days'}else{'dni'}) &bull; $($vals.Count) $(if($lEn){'readings'}else{'odczytow'})</div>
  </div>
</div>

<div class="patient-box">
  <div class="patient-name">$patFull</div>
  <div class="patient-meta">$(if($lEn){'Unit'}else{'Jednostka'}): $unit &nbsp;&bull;&nbsp; $(if($lEn){'Normal range'}else{'Norma'}): $($loN.ToString($fmt)) &ndash; $($hiN.ToString($fmt)) $unit</div>
</div>

<div class="grid">
  <div class="card"><div class="lbl">$(if($lEn){'Average'}else{'Srednia'})</div><div class="val c-blue">$($avg.ToString($fmt)) <span class="unit">$unit</span></div></div>
  <div class="card"><div class="lbl">TIR $(if($lEn){'in range'}else{'w normie'})</div><div class="val c-green">$tir%</div></div>
  <div class="card"><div class="lbl">eHbA1c</div><div class="val c-purple">$($hba1c.ToString('0.0'))%</div></div>
  <div class="card"><div class="lbl">CV%</div><div class="val c-teal">$cv%</div></div>
  <div class="card"><div class="lbl">Min</div><div class="val c-blue">$($sm.Minimum.ToString($fmt)) <span class="unit">$unit</span></div></div>
  <div class="card"><div class="lbl">Max</div><div class="val c-orange">$($sm.Maximum.ToString($fmt)) <span class="unit">$unit</span></div></div>
  <div class="card"><div class="lbl">SD</div><div class="val c-gray">$($sd.ToString($fmt)) <span class="unit">$unit</span></div></div>
  <div class="card"><div class="lbl">$(if($lEn){'Readings'}else{'Odczyty'})</div><div class="val c-gray">$($vals.Count)</div></div>
</div>

<div class="section-title">$(if($lEn){'Glucose history'}else{'Historia glukozy'})</div>
<div class="chart-box">
$svgHist
<div class="chart-legend">
  <span><span class="leg-line"></span>$(if($lEn){'Glucose readings'}else{'Odczyty glukozy'})</span>
  <span><span class="leg-green"></span>$(if($lEn){'Normal range'}else{'Norma'}) ($($loN.ToString($fmt))&ndash;$($hiN.ToString($fmt)) $unit)</span>
  <span><span class="leg-red"></span>$(if($lEn){'Low'}else{'Hipoglikemia'})</span>
  <span><span class="leg-or"></span>$(if($lEn){'High'}else{'Hiperglikemia'})</span>
</div>
</div>

<div class="section-title">$(if($lEn){'Daily pattern (AGP)'}else{'Wzorzec dobowy (AGP)'})</div>
<div class="chart-box">
$svgAgp
<div class="chart-legend">
  <span><span class="leg-avg"></span>$(if($lEn){'Hourly average'}else{'Srednia godzinowa'})</span>
  <span><span class="leg-green"></span>$(if($lEn){'Normal range'}else{'Norma'})</span>
  <span>$(if($lEn){'Bars: min-max range per hour'}else{'Slupki: zakres min-max na godzine'})</span>
</div>
</div>

<div class="section-title">$(if($lEn){'Average glucose by time of day (3h)'}else{'Srednie stezenie glukozy (przedzialy 3h)'})</div>
<div class="chart-box">
$svgBars
<div class="chart-legend">
  <span><span class="leg-green"></span>$(if($lEn){'In range'}else{'W normie'}) ($($loN.ToString($fmt))&ndash;$($hiN.ToString($fmt)) $unit)</span>
  <span><span class="dot-red"></span>$(if($lEn){'Below range'}else{'Ponizej normy'})</span>
  <span><span class="dot-orange"></span>$(if($lEn){'Above range'}else{'Powyzej normy'})</span>
</div>
</div>

<div class="section-title">$(if($lEn){'Daily glucose pattern (percentiles)'}else{'Wzorzec dobowy glukozy (percentyle)'})</div>
<div class="chart-box">
$svgPattern
<div class="chart-legend">
  <span style="display:inline-flex;align-items:center;margin-right:14px">
    <span style="display:inline-block;width:28px;height:10px;background:#7aaac8;opacity:0.4;border-radius:2px;margin-right:4px"></span>
    $(if($lEn){'10&ndash;90th percentile'}else{'Percentyl 10&ndash;90'})
  </span>
  <span style="display:inline-flex;align-items:center;margin-right:14px">
    <span style="display:inline-block;width:28px;height:10px;background:#4477aa;opacity:0.55;border-radius:2px;margin-right:4px"></span>
    $(if($lEn){'25&ndash;75th percentile (IQR)'}else{'Percentyl 25&ndash;75 (IQR)'})
  </span>
  <span style="display:inline-flex;align-items:center">
    <span style="display:inline-block;width:28px;height:3px;background:#1a3a6e;border-radius:1px;margin-right:4px"></span>
    $(if($lEn){'Median (50th percentile)'}else{'Mediana (percentyl 50)'})
  </span>
</div>
</div>

$doctorSection

<div class="footer">Glucose Monitor &bull; $(Get-Date -Format 'dd.MM.yyyy HH:mm') &bull; $patFull</div>
</body>
</html>
"@
    } catch { Write-Log "PDF HTML ERR: $($_.Exception.Message)"; return }

    # --- Zapisz plik ---
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title  = if ($script:LangEn) { "Save PDF report" } else { "Zapisz raport PDF" }
    $dlg.Filter = "PDF (*.pdf)|*.pdf"
    $safeName   = ($lastName -replace '[\\/:*?"<>|]', '') + "_$(Get-Date -Format 'yyyyMMdd')"
    $dlg.FileName = "glucose_$safeName.pdf"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $pdfPath = $dlg.FileName

    # --- Konwersja HTML do PDF przez Edge headless ---
    $tmpHtml = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "gm_report_$([System.IO.Path]::GetRandomFileName()).html")
    $html | Set-Content -Path $tmpHtml -Encoding UTF8

    $edge = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:LocalAppData\Microsoft\Edge\Application\msedge.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    $success = $false
    if ($edge) {
        try {
            $fileUri = "file:///" + ($tmpHtml -replace '\\', '/' -replace ' ', '%20')
            $edgeArgs = @(
                "--headless=new", "--disable-gpu", "--no-sandbox",
                "--run-all-compositor-stages-before-draw",
                "--disable-extensions",
                "--print-to-pdf=`"$pdfPath`"",
                "`"$fileUri`""
            )
            $p = Start-Process -FilePath $edge -ArgumentList $edgeArgs -Wait -PassThru -WindowStyle Hidden
            if ((Test-Path $pdfPath) -and (Get-Item $pdfPath).Length -gt 1000) {
                $success = $true
            }
        } catch { Write-Log "PDF Edge ERR: $($_.Exception.Message)" }
    }

    Remove-Item $tmpHtml -ErrorAction SilentlyContinue

    if ($success) {
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) { "PDF saved:`n$pdfPath" } else { "PDF zapisany:`n$pdfPath" }),
            "Glucose Monitor") | Out-Null
        Start-Process $pdfPath
    } else {
        # Fallback - zapisz HTML, otworz w przegladarce
        $htmlPath = $pdfPath -replace '\.pdf$', '.html'
        $html | Set-Content -Path $htmlPath -Encoding UTF8
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) {
                "Microsoft Edge not found for PDF conversion.`nOpening HTML report in browser - use Ctrl+P -> Save as PDF."
              } else {
                "Nie znaleziono Edge do konwersji PDF.`nOtwarto raport HTML w przegladarce - uzyj Ctrl+P -> Zapisz jako PDF."
              }),
            "Glucose Monitor") | Out-Null
        Start-Process $htmlPath
    }
}

# ======================== SREDNIE STEZENIE (SLUPKI) ========================
function Show-AvgBarsWindow {
    $data = Load-HistoryData $script:HistDays $script:HistOffset
    if ($data.Count -lt 2) {
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) { "Not enough data for this period." } else { "Za malo danych dla wybranego okresu." }),
            "Glucose Monitor") | Out-Null; return
    }

    # ── Pre-compute slot averages BEFORE window creation ──────────────────────
    # Używamy prostych tablic [double[]] i [bool[]] – bez List[double] ani Measure-Object w closurze
    $slotSums   = [double[]]::new(8)
    $slotCounts = [int[]]::new(8)
    $allSum     = 0.0
    $allCount   = 0
    $daySet     = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($pt in $data) {
        try {
            $ts   = [DateTime]::Parse($pt.ts)
            $slot = [int][Math]::Floor($ts.Hour / 3)   # 0..7
            $val  = if ($script:UseMgDl) { [double]$pt.mgdl } else { [double]$pt.mgdl / 18.018 }
            $slotSums[$slot]   += $val
            $slotCounts[$slot] += 1
            $allSum   += $val
            $allCount += 1
            $daySet.Add($ts.ToString("yyyy-MM-dd")) | Out-Null
        } catch {}
    }

    # Tablice pre-computed dla closure
    $script:AvgBarsAvgs    = [double[]]::new(8)
    $script:AvgBarsHasData = [bool[]]::new(8)
    $validMin = [double]::MaxValue
    $validMax = [double]::MinValue
    for ($s = 0; $s -lt 8; $s++) {
        if ($slotCounts[$s] -gt 0) {
            $script:AvgBarsAvgs[$s]    = $slotSums[$s] / $slotCounts[$s]
            $script:AvgBarsHasData[$s] = $true
            if ($script:AvgBarsAvgs[$s] -lt $validMin) { $validMin = $script:AvgBarsAvgs[$s] }
            if ($script:AvgBarsAvgs[$s] -gt $validMax) { $validMax = $script:AvgBarsAvgs[$s] }
        }
    }
    if ($validMin -eq [double]::MaxValue) {
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) { "Not enough data for this period." } else { "Za malo danych dla wybranego okresu." }),
            "Glucose Monitor") | Out-Null; return
    }

    # Y range – pre-computed, zapisany w script scope
    $yPad = ($validMax - $validMin) * 0.3; if ($yPad -lt 1.0) { $yPad = 1.0 }
    $script:AvgBarsLoY  = [Math]::Max(0.0, $validMin - $yPad)
    $script:AvgBarsHiY  = $validMax + $yPad
    $script:AvgBarsRngY = $script:AvgBarsHiY - $script:AvgBarsLoY
    if ($script:AvgBarsRngY -lt 0.1) { $script:AvgBarsRngY = 1.0 }

    # Progi kolorow
    $script:AvgBarsLoN     = if ($script:UseMgDl) { 70.0  } else { 3.9  }
    $script:AvgBarsHiN     = if ($script:UseMgDl) { 180.0 } else { 10.0 }
    $script:AvgBarsHiC     = if ($script:UseMgDl) { 250.0 } else { 13.9 }
    $script:AvgBarsUseMgDl = $script:UseMgDl

    # Etykiety okna
    $overallAvg = if ($allCount -gt 0) { $allSum / $allCount } else { 0.0 }
    $fmt        = if ($script:UseMgDl) { "0" } else { "0.0" }
    $unit       = if ($script:UseMgDl) { "mg/dL" } else { "mmol/L" }
    $daysAvail  = $daySet.Count
    $endDate    = (Get-Date).AddDays(-$script:HistOffset)
    $startDate  = $endDate.AddDays(-$script:HistDays)
    $dateRange  = "$($startDate.ToString('d/MMM/yyyy')) - $($endDate.ToString('d/MMM/yyyy'))"
    $avgLabel   = if ($script:LangEn) { "Average: $([Math]::Round($overallAvg,1).ToString($fmt)) $unit" } else { "Srednia: $([Math]::Round($overallAvg,1).ToString($fmt)) $unit" }
    $daysLabel  = if ($script:LangEn) { "Data available for $daysAvail of $($script:HistDays) days" } else { "Dane dostepne dla $daysAvail z $($script:HistDays) dni" }

    $xamlB = [xml]@"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Bars" Width="520" Height="440"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Border Background="#111827" CornerRadius="10" BorderBrush="#334466" BorderThickness="1">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="36"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0" Background="#0d1520" Name="barsTitleBar">
                <TextBlock Name="barsTitle" Text=""
                           Foreground="White" FontSize="12" FontFamily="Segoe UI Semibold"
                           VerticalAlignment="Center" Margin="12,0,0,0"/>
                <Button Name="barsClose" Content="X" Width="30" Height="30" HorizontalAlignment="Right"
                        Background="Transparent" Foreground="#7777aa" BorderThickness="0" Cursor="Hand"
                        FontSize="14" FontWeight="Bold"/>
            </Grid>
            <Grid Grid.Row="1">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Name="barsDateRange" Text=""
                           Foreground="#aaaacc" FontSize="11" FontFamily="Segoe UI"
                           HorizontalAlignment="Center" Margin="0,6,0,2"/>
                <Canvas Grid.Row="1" Name="barsCanvas" ClipToBounds="True" Margin="8,4,8,4"/>
                <TextBlock Grid.Row="2" Name="barsAvgLbl" Text=""
                           Foreground="White" FontSize="13" FontFamily="Segoe UI Semibold"
                           HorizontalAlignment="Center" Margin="0,4,0,2"/>
                <TextBlock Grid.Row="3" Name="barsDaysLbl" Text=""
                           Foreground="#FFAA44" FontSize="10" FontFamily="Segoe UI"
                           HorizontalAlignment="Center" Margin="0,0,0,8"/>
            </Grid>
        </Grid>
    </Border>
</Window>
"@
    try {
        if ($script:AvgBarsWin -and $script:AvgBarsWin.IsLoaded) { $script:AvgBarsWin.Close() }
        $script:AvgBarsWin = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xamlB))
        $script:AvgBarsWin.FindName("barsClose").Add_Click({ $script:AvgBarsWin.Close() })
        $script:AvgBarsWin.FindName("barsTitleBar").Add_MouseLeftButtonDown({ $script:AvgBarsWin.DragMove() })
        $script:AvgBarsWin.FindName("barsTitle").Text     = if ($script:LangEn) { "Average Glucose (bars)" } else { "Srednie stezenie glukozy" }
        $script:AvgBarsWin.FindName("barsDateRange").Text = $dateRange
        $script:AvgBarsWin.FindName("barsAvgLbl").Text    = $avgLabel
        $script:AvgBarsWin.FindName("barsDaysLbl").Text   = $daysLabel

        $script:AvgBarsWin.Add_ContentRendered({
          try {
            $cv = $script:AvgBarsWin.FindName("barsCanvas")
            $script:AvgBarsWin.UpdateLayout()
            $cW = $cv.ActualWidth;  if ($cW -lt 10) { $cW = 480.0 }
            $cH = $cv.ActualHeight; if ($cH -lt 10) { $cH = 300.0 }
            $padL = 46.0; $padR = 10.0; $padT = 22.0; $padB = 26.0
            $gW   = $cW - $padL - $padR
            $gH   = $cH - $padT - $padB

            # Użyj pre-computed Y range
            $loY  = $script:AvgBarsLoY
            $rngY = $script:AvgBarsRngY
            $fmt2  = if ($script:AvgBarsUseMgDl) { "0" } else { "0.0" }
            $decPl = if ($script:AvgBarsUseMgDl) { 0 } else { 1 }

            # Poziome linie siatki (5 linii)
            for ($gi = 0; $gi -le 4; $gi++) {
                $gv = $loY + $rngY * ($gi / 4.0)
                $yp = $padT + $gH - ($gv - $loY) / $rngY * $gH
                $gl = New-Object System.Windows.Shapes.Line
                $gl.X1 = $padL; $gl.X2 = $padL + $gW; $gl.Y1 = $yp; $gl.Y2 = $yp
                $gl.Stroke = $script:CBrGrid33
                $gl.StrokeThickness = 0.5
                $cv.Children.Add($gl) | Out-Null
                $yl = New-Object System.Windows.Controls.TextBlock
                $yl.Text       = [Math]::Round($gv, $decPl).ToString($fmt2)
                $yl.FontSize   = 8
                $yl.Foreground = $script:CBrLabel2
                $yl.FontFamily = $script:CFontUI
                [System.Windows.Controls.Canvas]::SetLeft($yl, 2)
                [System.Windows.Controls.Canvas]::SetTop($yl, $yp - 7)
                $cv.Children.Add($yl) | Out-Null
            }

            # Słupki – 9 kolumn: 00:00..21:00 + 00:00 (zamknięcie cyklu dobowego)
            # Slot 8 (ostatni "00:00") używa tych samych danych co slot 0 (północ)
            $slotW = $gW / 9.0
            $barW  = $slotW * 0.60

            for ($s = 0; $s -lt 9; $s++) {
                $xCenter  = $padL + ($s + 0.5) * $slotW
                $dataSlot = $s % 8   # slot 8 → te same dane co slot 0

                # Etykieta godziny na osi X (dla s=8: 24%24=0 → "00:00")
                $hh   = ($s * 3) % 24
                $xLbl = New-Object System.Windows.Controls.TextBlock
                $xLbl.Text       = "$($hh.ToString('00')):00"
                $xLbl.FontSize   = 8
                $xLbl.Foreground = $script:CBrLabel
                $xLbl.FontFamily = $script:CFontUI
                [System.Windows.Controls.Canvas]::SetLeft($xLbl, $xCenter - 14)
                [System.Windows.Controls.Canvas]::SetTop($xLbl,  $padT + $gH + 5)
                $cv.Children.Add($xLbl) | Out-Null

                if (-not $script:AvgBarsHasData[$dataSlot]) { continue }
                $avg = $script:AvgBarsAvgs[$dataSlot]

                # Kolor słupka
                $barBr = if   ($avg -lt $script:AvgBarsLoN -or $avg -gt $script:AvgBarsHiC) { $script:CBrCC4444 } `
                          elseif ($avg -gt $script:AvgBarsHiN) { $script:CBrOrange } `
                          else { $script:CBrGreen2 }

                # Wysokość słupka od dołu osi
                $barH   = [Math]::Max(2.0, ($avg - $loY) / $rngY * $gH)
                $barTop = $padT + $gH - $barH

                $bar = New-Object System.Windows.Shapes.Rectangle
                $bar.Width   = $barW
                $bar.Height  = $barH
                $bar.RadiusX = 3; $bar.RadiusY = 3
                $bar.Fill    = $barBr
                [System.Windows.Controls.Canvas]::SetLeft($bar, $xCenter - $barW / 2.0)
                [System.Windows.Controls.Canvas]::SetTop($bar,  $barTop)
                $cv.Children.Add($bar) | Out-Null

                # Wartość średnia nad słupkiem
                $valLbl = New-Object System.Windows.Controls.TextBlock
                $valLbl.Text       = [Math]::Round($avg, $decPl).ToString($fmt2)
                $valLbl.FontSize   = 9
                $valLbl.FontFamily = $script:CFontUIBold
                $valLbl.Foreground = $script:CBrWhite
                [System.Windows.Controls.Canvas]::SetLeft($valLbl, $xCenter - 10)
                [System.Windows.Controls.Canvas]::SetTop($valLbl,  $barTop - 14)
                $cv.Children.Add($valLbl) | Out-Null
            }

          } catch { Write-Log "AvgBars Render ERR: $($_.Exception.Message)" }
        })

        # Zapamietaj pozycje przy zamknieciu
        $script:AvgBarsWin.Add_Closing({
            $script:AvgBarsWinLeft = $script:AvgBarsWin.Left
            $script:AvgBarsWinTop  = $script:AvgBarsWin.Top
        })

        # Przywroc zapamietana pozycje (jesli istnieje)
        if ($null -ne $script:AvgBarsWinLeft -and $null -ne $script:AvgBarsWinTop) {
            $script:AvgBarsWin.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
            $script:AvgBarsWin.Left = $script:AvgBarsWinLeft
            $script:AvgBarsWin.Top  = $script:AvgBarsWinTop
        }

        $script:AvgBarsWin.Show()
    } catch { Write-Log "AvgBars ERR: $($_.Exception.Message)" }
}

# ======================== AGP (WZORZEC DOBOWY) ========================
function Show-AgpWindow {
    $allData = Load-HistoryData 90
    if ($allData.Count -lt 10) {
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) { "Not enough data for AGP (min. 10 readings)." } else { "Za malo danych dla AGP (min. 10 odczytow)." }),
            "Glucose Monitor") | Out-Null; return
    }
    # Grupuj po godzinie - zapisz w script scope bo closure nie widzi lokalnych zmiennych funkcji
    $script:AgpByHour = @{}
    foreach ($pt in $allData) {
        try {
            $h   = [DateTime]::Parse($pt.ts).Hour
            $val = if ($script:UseMgDl) { [double]$pt.mgdl } else { [Math]::Round([double]$pt.mgdl / 18.018, 1) }
            if (-not $script:AgpByHour.ContainsKey($h)) { $script:AgpByHour[$h] = [System.Collections.Generic.List[double]]::new() }
            $script:AgpByHour[$h].Add($val)
        } catch {}
    }
    $script:AgpUseMgDl = $script:UseMgDl
    $xamlA = [xml]@"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AGP" Width="620" Height="400"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Border Background="#111827" CornerRadius="10" BorderBrush="#334466" BorderThickness="1">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="36"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0" Background="#0d1520" Name="agpTitleBar">
                <TextBlock Name="agpTitle" Text="AGP"
                           Foreground="White" FontSize="12" FontFamily="Segoe UI Semibold"
                           VerticalAlignment="Center" Margin="12,0,0,0"/>
                <Button Name="agpClose" Content="X" Width="30" Height="30" HorizontalAlignment="Right"
                        Background="Transparent" Foreground="#7777aa" BorderThickness="0" Cursor="Hand"
                        FontSize="14" FontWeight="Bold"/>
            </Grid>
            <Canvas Name="agpCanvas" Grid.Row="1" ClipToBounds="True" Margin="8"/>
        </Grid>
    </Border>
</Window>
"@
    try {
        if ($script:AgpWin -and $script:AgpWin.IsLoaded) { $script:AgpWin.Close() }
        $script:AgpWin = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xamlA))
        $script:AgpWin.FindName("agpClose").Add_Click({ $script:AgpWin.Close() })
        $script:AgpWin.FindName("agpTitleBar").Add_MouseLeftButtonDown({ $script:AgpWin.DragMove() })
        $script:AgpWin.FindName("agpTitle").Text = if ($script:LangEn) { "Daily Pattern (AGP) - last 90 days" } else { "Wzorzec dobowy (AGP) - ostatnie 90 dni" }
        $script:AgpWin.Add_ContentRendered({
          try {
            $cv  = $script:AgpWin.FindName("agpCanvas")
            $script:AgpWin.UpdateLayout()
            $cW  = $cv.ActualWidth;  if ($cW -lt 10) { $cW  = 580.0 }
            $cH  = $cv.ActualHeight; if ($cH -lt 10) { $cH  = 330.0 }
            $padL = 38.0; $padR = 8.0; $padT = 8.0; $padB = 28.0
            $gW   = $cW - $padL - $padR
            $gH   = $cH - $padT  - $padB
            $loY  = if ($script:AgpUseMgDl) { 40.0  } else { 2.2  }
            $hiY  = if ($script:AgpUseMgDl) { 280.0 } else { 15.5 }
            $rngY = $hiY - $loY
            $barW = $gW / 24 * 0.6
            $loN  = if ($script:AgpUseMgDl) { 70.0  } else { 3.9  }
            $hiN  = if ($script:AgpUseMgDl) { 180.0 } else { 10.0 }
            # Strefa normy
            $yLo  = $padT + $gH - ($loN - $loY)/$rngY * $gH
            $yHi  = $padT + $gH - ($hiN - $loY)/$rngY * $gH
            $normR = New-Object System.Windows.Shapes.Rectangle
            $normR.Fill = $script:CBrNormZone3
            $normR.Width = $gW; $normR.Height = [Math]::Abs($yLo - $yHi)
            [System.Windows.Controls.Canvas]::SetLeft($normR, $padL)
            [System.Windows.Controls.Canvas]::SetTop($normR,  [Math]::Min($yLo,$yHi))
            $cv.Children.Add($normR) | Out-Null
            # Slupki i srednie per godzina
            for ($h = 0; $h -lt 24; $h++) {
                $xCenter = $padL + ($h + 0.5) / 24 * $gW
                if ($script:AgpByHour.ContainsKey($h) -and $script:AgpByHour[$h].Count -gt 0) {
                    $hVals  = $script:AgpByHour[$h]
                    $avg    = ($hVals | Measure-Object -Average).Average
                    $minV   = ($hVals | Measure-Object -Minimum).Minimum
                    $maxV   = ($hVals | Measure-Object -Maximum).Maximum
                    $minV   = [Math]::Max($minV, $loY); $maxV = [Math]::Min($maxV, $hiY)
                    $yAvg   = $padT + $gH - ($avg  - $loY)/$rngY * $gH
                    $yMin   = $padT + $gH - ($minV - $loY)/$rngY * $gH
                    $yMax   = $padT + $gH - ($maxV - $loY)/$rngY * $gH
                    # Slupek min-max
                    $bar  = New-Object System.Windows.Shapes.Rectangle
                    $barBr = if ($avg -lt $loN) { $script:CBrAgpHypo } elseif ($avg -gt $hiN) { $script:CBrAgpHyper } else { $script:CBrAgpNorm }
                    $bar.Fill   = $barBr
                    $bar.Width  = $barW; $bar.Height = [Math]::Max(2, [Math]::Abs($yMin - $yMax))
                    [System.Windows.Controls.Canvas]::SetLeft($bar, $xCenter - $barW/2)
                    [System.Windows.Controls.Canvas]::SetTop($bar,  [Math]::Min($yMin,$yMax))
                    $cv.Children.Add($bar) | Out-Null
                    # Punkt sredniej
                    $dot = New-Object System.Windows.Shapes.Ellipse
                    $dot.Width = 6; $dot.Height = 6
                    $dot.Fill  = $script:CBrWhite
                    [System.Windows.Controls.Canvas]::SetLeft($dot, $xCenter - 3)
                    [System.Windows.Controls.Canvas]::SetTop($dot,  $yAvg - 3)
                    $cv.Children.Add($dot) | Out-Null
                }
                # Etykieta godziny co 3h
                if ($h % 3 -eq 0) {
                    $lbl = New-Object System.Windows.Controls.TextBlock
                    $lbl.Text = "${h}h"; $lbl.FontSize = 8
                    $lbl.Foreground = $script:CBrLabel
                    $lbl.FontFamily = $script:CFontUI
                    [System.Windows.Controls.Canvas]::SetLeft($lbl, $xCenter - 8)
                    [System.Windows.Controls.Canvas]::SetTop($lbl,  $padT + $gH + 4)
                    $cv.Children.Add($lbl) | Out-Null
                }
            }
            # Os Y (linie siatki + etykiety)
            foreach ($yv in @($loY, ($loY+$rngY*0.25), ($loY+$rngY*0.5), ($loY+$rngY*0.75), $hiY)) {
                $yp = $padT + $gH - ($yv - $loY)/$rngY * $gH
                $gl = New-Object System.Windows.Shapes.Line
                $gl.X1=$padL; $gl.X2=$padL+$gW; $gl.Y1=$yp; $gl.Y2=$yp
                $gl.Stroke = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#22ffffff"))
                $gl.StrokeThickness = 0.5; $cv.Children.Add($gl) | Out-Null
                $yl = New-Object System.Windows.Controls.TextBlock
                $dispVal = if ($script:AgpUseMgDl) { [Math]::Round($yv,0).ToString() } else { $yv.ToString("0.0") }
                $yl.Text = $dispVal; $yl.FontSize = 8
                $yl.Foreground = $script:CBrLabel
                $yl.FontFamily = $script:CFontUI
                [System.Windows.Controls.Canvas]::SetLeft($yl, 2)
                [System.Windows.Controls.Canvas]::SetTop($yl,  $yp - 7)
                $cv.Children.Add($yl) | Out-Null
            }
          } catch { Write-Log "AGP Render ERR: $($_.Exception.Message)" }
        })

        # Zapamietaj pozycje przy zamknieciu
        $script:AgpWin.Add_Closing({
            $script:AgpWinLeft = $script:AgpWin.Left
            $script:AgpWinTop  = $script:AgpWin.Top
        })

        # Przywroc zapamietana pozycje (jesli istnieje)
        if ($null -ne $script:AgpWinLeft -and $null -ne $script:AgpWinTop) {
            $script:AgpWin.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
            $script:AgpWin.Left = $script:AgpWinLeft
            $script:AgpWin.Top  = $script:AgpWinTop
        }

        $script:AgpWin.Show()
    } catch { Write-Log "AGP ERR: $($_.Exception.Message)" }
}

# ======================== RAPORT HTML ========================
function Export-HtmlReport {
    $days = if ($script:HistWin -and $script:HistWin.IsLoaded) { $script:HistDays } else { 14 }
    $data = Load-HistoryData $days
    if ($data.Count -lt 2) {
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) { "Not enough data for report." } else { "Za malo danych do raportu." }),
            "Glucose Monitor") | Out-Null; return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title  = if ($script:LangEn) { "Save HTML report" } else { "Zapisz raport HTML" }
    $dlg.Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
    $dlg.FileName = "glucose_report_$(Get-Date -Format 'yyyyMMdd').html"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    try {
        $vals  = $data | ForEach-Object { if ($script:UseMgDl) { [double]$_.mgdl } else { [Math]::Round([double]$_.mgdl/18.018,1) } }
        $sm    = $vals | Measure-Object -Min -Max -Average
        $avg   = [Math]::Round($sm.Average, 1)
        $loN   = if ($script:UseMgDl) { 70.0 } else { 3.9 }
        $hiN   = if ($script:UseMgDl) { 180.0 } else { 10.0 }
        $tir   = [Math]::Round(($vals | Where-Object { $_ -ge $loN -and $_ -le $hiN }).Count / $vals.Count * 100, 0)
        $mean  = $sm.Average
        $sd    = [Math]::Round([Math]::Sqrt(($vals | ForEach-Object { ($_ - $mean)*($_ - $mean) } | Measure-Object -Sum).Sum / $vals.Count), 1)
        $cv    = if ($mean -gt 0) { [Math]::Round($sd/$mean*100, 0) } else { 0 }
        $avgMg = if ($script:UseMgDl) { $mean } else { $mean * 18.018 }
        $hba1c = [Math]::Round(($avgMg + 46.7) / 28.7, 1)
        $unit  = if ($script:UseMgDl) { "mg/dL" } else { "mmol/L" }
        $fmt   = if ($script:UseMgDl) { "0" } else { "0.0" }
        $patientName = if ($txtPatient -and $txtPatient.Text) { $txtPatient.Text } else { "---" }
        $dateFrom = (Get-Date).AddDays(-$days).ToString("dd.MM.yyyy")
        $dateTo   = (Get-Date).ToString("dd.MM.yyyy")
        $lEn = $script:LangEn
        $svgHist    = Get-HistSvg          -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn
        $svgAgp     = Get-AgpSvg           -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn
        $svgBars    = Get-BarsSvg          -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn
        $svgPattern = Get-DailyPatternSvg  -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn -Days $days
        $html = @"
<!DOCTYPE html>
<html lang="$(if ($lEn){'en'}else{'pl'})">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Glucose Monitor</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:Segoe UI,Arial,sans-serif;background:#0d1520;color:#ccc;padding:16px}
  h1{color:#fff;font-size:18px;margin-bottom:2px}
  .sub{color:#7777aa;font-size:12px;margin-bottom:16px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px}
  .card{background:#1a2540;border-radius:8px;padding:12px;text-align:center}
  .lbl{color:#7777aa;font-size:11px;margin-bottom:4px}
  .val{font-size:20px;font-weight:600;line-height:1.2}
  .unit{font-size:12px;font-weight:400}
  .green{color:#44dd44}.blue{color:#44aaff}.orange{color:#ffaa44}
  .purple{color:#cc88ff}.cyan{color:#88ccff}.yellow{color:#ffcc66}
  .section-title{color:#7777aa;font-size:12px;font-weight:600;margin:18px 0 8px;text-transform:uppercase;letter-spacing:0.05em}
  .footer{color:#445566;font-size:10px;margin-top:14px;text-align:center}
  svg{max-width:100%;height:auto;display:block}
  @media(min-width:480px){.grid{grid-template-columns:repeat(4,1fr)}}
</style>
</head>
<body>
<h1>Glucose Monitor</h1>
<div class="sub">$patientName &nbsp;&#183;&nbsp; $dateFrom &ndash; $dateTo &nbsp;&#183;&nbsp; $days $(if ($lEn){'days'}else{'dni'})</div>
<div class="grid">
  <div class="card"><div class="lbl">$(if ($lEn){'Average'}else{'Srednia'})</div><div class="val">$($avg.ToString($fmt)) <span class="unit">$unit</span></div></div>
  <div class="card"><div class="lbl">TIR</div><div class="val green">$tir%</div></div>
  <div class="card"><div class="lbl">eHbA1c</div><div class="val purple">$($hba1c.ToString('0.0'))%</div></div>
  <div class="card"><div class="lbl">SD / CV%</div><div class="val cyan">$($sd.ToString($fmt)) <span class="unit">/ $cv%</span></div></div>
  <div class="card"><div class="lbl">Min</div><div class="val blue">$($sm.Minimum.ToString($fmt)) <span class="unit">$unit</span></div></div>
  <div class="card"><div class="lbl">Max</div><div class="val orange">$($sm.Maximum.ToString($fmt)) <span class="unit">$unit</span></div></div>
  <div class="card"><div class="lbl">$(if ($lEn){'Readings'}else{'Odczyty'})</div><div class="val">$($vals.Count)</div></div>
  <div class="card"><div class="lbl">$(if ($lEn){'Low alert'}else{'Prog niski'})</div><div class="val">$(if ($script:UseMgDl){[Math]::Round($script:Config.AlertLow*18.018,0)}else{$script:Config.AlertLow}) <span class="unit">$unit</span></div></div>
</div>
<div class="section-title">$(if ($lEn){'Glucose history'}else{'Historia glukozy'})</div>
$svgHist
<div class="section-title">$(if ($lEn){'Daily pattern (AGP)'}else{'Wzorzec dobowy (AGP)'})</div>
$svgAgp
<div class="section-title">$(if ($lEn){'Distribution by time (BARS)'}else{'Rozklad wg czasu (BARS)'})</div>
$svgBars
<div class="section-title">$(if ($lEn){'Daily glucose pattern (percentiles)'}else{'Wzorzec dobowy glukozy (percentyle)'})</div>
$svgPattern
<div class="footer">Glucose Monitor &bull; $(Get-Date -Format 'dd.MM.yyyy HH:mm')</div>
</body>
</html>
"@
        $html | Set-Content -Path $dlg.FileName -Encoding UTF8
        Start-Process $dlg.FileName
    } catch { [System.Windows.MessageBox]::Show("Blad: $($_.Exception.Message)", "Glucose Monitor") | Out-Null }
}

# ======================== SKROT NA PULPICIE ========================
function Create-DesktopShortcut {
    try {
        $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $args  = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$script:ScriptPath`""
        $lnk   = Join-Path ([Environment]::GetFolderPath("Desktop")) "Glucose Monitor.lnk"

        $wsh      = New-Object -ComObject WScript.Shell
        $shortcut = $wsh.CreateShortcut($lnk)
        $shortcut.TargetPath       = $psExe
        $shortcut.Arguments        = $args
        $shortcut.WorkingDirectory = $script:ScriptDir
        $shortcut.Description      = "Glucose Monitor"
        $shortcut.WindowStyle      = 7   # 7 = zminimalizowany (nie pokazuj konsoli)
        $shortcut.Save()
        Write-Log "Skrot na pulpicie utworzony: $lnk"
    } catch {
        Write-Log "Create-DesktopShortcut ERR: $($_.Exception.Message)"
    }
}

# ======================== AUTOSTART ========================
$script:TaskName = "GlucoseMonitor"

function Get-AutoStartStatus {
    $task = Get-ScheduledTask -TaskName $script:TaskName -ErrorAction SilentlyContinue
    return ($null -ne $task)
}

function Install-AutoStart([switch]$Silent) {
    $psExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if ($psExe -notmatch 'powershell|pwsh') {
        $psExe = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Source
        if (-not $psExe) { $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" }
    }
    $scriptPath = $script:ScriptPath
    $taskArgs   = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -AutoStart"
    $action     = New-ScheduledTaskAction -Execute $psExe -Argument $taskArgs
    $trigger    = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    $settings   = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0

    # Probuj najpierw z Highest (wymaga admina), jesli blad - uzyj Limited
    $registered = $false
    foreach ($level in @('Highest','Limited')) {
        try {
            $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel $level -LogonType Interactive
            Register-ScheduledTask -TaskName $script:TaskName -Action $action -Trigger $trigger `
                -Settings $settings -Principal $principal -Force -ErrorAction Stop | Out-Null
            Write-Log "AutoStart: zadanie zarejestrowane (RunLevel=$level)"
            $registered = $true
            break
        } catch {
            Write-Log "AutoStart blad RunLevel=$level : $($_.Exception.Message)"
        }
    }

    if ($registered) {
        if (-not $Silent) {
            [System.Windows.MessageBox]::Show(
                $(if ($script:LangEn) { "Autostart enabled.`nGlucose Monitor will start with Windows." }
                  else                { "Autostart wlaczony.`nGlucose Monitor bedzie uruchamiany z Windows." }),
                "Glucose Monitor") | Out-Null
        }
    } else {
        # Zawsze informuj o niepowodzeniu - nawet w trybie Silent
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) { "Could not register autostart task.`nPlease run the program as Administrator once to enable autostart." }
              else                { "Nie mozna zarejestr. autostartu.`nUruchom program jako Administrator aby wlaczyc autostart." }),
            "Glucose Monitor") | Out-Null
    }
}

function Uninstall-AutoStart {
    try {
        Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false -ErrorAction Stop
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) { "Autostart disabled." }
              else                { "Autostart wylaczony." }),
            "Glucose Monitor") | Out-Null
    } catch {
        [System.Windows.MessageBox]::Show(
            $(if ($script:LangEn) { "Failed to disable autostart:`n$($_.Exception.Message)" }
              else                { "Blad przy wylaczaniu autostartu:`n$($_.Exception.Message)" }),
            "Glucose Monitor") | Out-Null
    }
}

# ======================== TRAY ========================
$script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
$script:NotifyIcon.Text = "Glucose Monitor"; $script:NotifyIcon.Visible = $true

function Update-TrayIcon {
    param([double]$mmol = 0, [int]$trend = 0)
    try {
        $bmp = New-Object System.Drawing.Bitmap(32, 32)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode   = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.Clear([System.Drawing.Color]::Transparent)

        # Kolor tla wg poziomu glukozy
        if ($mmol -le 0) {
            $c1 = [System.Drawing.Color]::FromArgb(80, 80, 100)
            $c2 = [System.Drawing.Color]::FromArgb(40, 40, 60)
        } elseif ($mmol -lt 3.9) {
            $c1 = [System.Drawing.Color]::FromArgb(255, 80,  40)
            $c2 = [System.Drawing.Color]::FromArgb(180,  0,   0)
        } elseif ($mmol -le 10.0) {
            $c1 = [System.Drawing.Color]::FromArgb(60, 200, 80)
            $c2 = [System.Drawing.Color]::FromArgb( 0, 120, 40)
        } elseif ($mmol -le 13.9) {
            $c1 = [System.Drawing.Color]::FromArgb(255, 180,  0)
            $c2 = [System.Drawing.Color]::FromArgb(180, 100,  0)
        } else {
            $c1 = [System.Drawing.Color]::FromArgb(255, 60, 60)
            $c2 = [System.Drawing.Color]::FromArgb(160,  0,  0)
        }

        # Zaokraglone tlo - rysuj lukami
        $bgPath = New-Object System.Drawing.Drawing2D.GraphicsPath
        [float]$r  = 6.0
        [float]$d  = $r * 2       # srednica luku = 12
        [float]$ed = 32.0 - $d   # przesuniecie prawej/dolnej krawedzi = 20
        $bgPath.AddArc([float]0,  [float]0,  $d, $d, [float]180, [float]90)
        $bgPath.AddArc($ed,       [float]0,  $d, $d, [float]270, [float]90)
        $bgPath.AddArc($ed,       $ed,       $d, $d, [float]0,   [float]90)
        $bgPath.AddArc([float]0,  $ed,       $d, $d, [float]90,  [float]90)
        $bgPath.CloseFigure()

        $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            (New-Object System.Drawing.PointF([float]0, [float]0)),
            (New-Object System.Drawing.PointF([float]0, [float]32)), $c1, $c2)
        $g.FillPath($bgBrush, $bgPath)
        $bgBrush.Dispose()
        $bgPath.Dispose()

        $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $sf    = New-Object System.Drawing.StringFormat
        $sf.Alignment     = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center

        if ($mmol -gt 0) {
            # Wartosc glukozy
            $valStr   = if ($script:UseMgDl) { [Math]::Round($mmol * 18.018, 0).ToString() } else { $mmol.ToString("0.0") }
            $fontSize = if ($valStr.Length -ge 4) { 10.0 } else { 13.0 }
            $font     = New-Object System.Drawing.Font("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold)
            $g.DrawString($valStr, $font, $white,
                (New-Object System.Drawing.RectangleF([float]0, [float]1, [float]32, [float]26)), $sf)
            $font.Dispose()

            # Strzalka trendu - prawy dolny rog
            $dn = [char]0x2193; $up = [char]0x2191
            $se = [char]0x2198; $ne = [char]0x2197  # ↘ ↗
            $arrow = switch ($trend) {
                1 { "$dn$dn" }   # ↓↓
                2 { "$dn"    }   # ↓
                6 { "$se"    }   # ↘
                4 { "$ne"    }   # ↗
                7 { "$up"    }   # ↑
                5 { "$up$up" }   # ↑↑
                default { "" }
            }
            if ($arrow) {
                $sfR               = New-Object System.Drawing.StringFormat
                $sfR.Alignment     = [System.Drawing.StringAlignment]::Far
                $sfR.LineAlignment = [System.Drawing.StringAlignment]::Far
                $fontA = New-Object System.Drawing.Font("Segoe UI", 7.0, [System.Drawing.FontStyle]::Bold)
                $g.DrawString($arrow, $fontA, $white,
                    (New-Object System.Drawing.RectangleF([float]0, [float]0, [float]31, [float]31)), $sfR)
                $fontA.Dispose(); $sfR.Dispose()
            }
        } else {
            # Brak danych - znak zapytania
            $gray = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(180, 180, 200))
            $font = New-Object System.Drawing.Font("Segoe UI", 8.0, [System.Drawing.FontStyle]::Bold)
            $g.DrawString("?", $font, $gray,
                (New-Object System.Drawing.RectangleF([float]0, [float]0, [float]32, [float]32)), $sf)
            $font.Dispose(); $gray.Dispose()
        }

        $white.Dispose(); $sf.Dispose(); $g.Dispose()

        $hicon   = $bmp.GetHicon()
        $newIcon = [System.Drawing.Icon]::FromHandle($hicon)
        $bmp.Dispose()
        $oldIcon = $script:NotifyIcon.Icon
        $script:NotifyIcon.Icon = $newIcon
        try { if ($oldIcon) { $oldIcon.Dispose() } } catch {}
    } catch {
        Write-Log "Update-TrayIcon ERR: $($_.Exception.Message)"
    }
}

Update-TrayIcon

$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuShow = $trayMenu.Items.Add((t "ShowWin"))
$menuShow.Add_Click({
    if ($script:IsCompact) {
        $btnCompact.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    }
    $window.ShowInTaskbar = $true
    $window.Show()
    $window.WindowState = [System.Windows.WindowState]::Normal
    $window.Topmost = $true
    $window.Activate()
})
$menuSettings = $trayMenu.Items.Add((t "SettingsMenu"))
$menuSettings.Add_Click({ Show-SettingsWindow })
$menuBackup = $trayMenu.Items.Add((t "BackupMenu"))
$menuBackup.Add_Click({ Backup-History })
$menuRestore = $trayMenu.Items.Add((t "RestoreMenu"))
$menuRestore.Add_Click({ Restore-History })
$trayMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
$menuLogout = $trayMenu.Items.Add((t "SwitchAcc"))
$menuLogout.Add_Click({
    $script:AuthToken = $null; $script:PatientId = $null
    if (Test-Path $script:ConfigFile) { Remove-Item $script:ConfigFile -Force }
    $result = Show-LoginWindow
    if ($result) { Update-Display; $script:SecondsLeft = $script:Config.Interval }
})
$menuAutoStart = $trayMenu.Items.Add($(if (Get-AutoStartStatus) { (t "AutoStartOn") } else { (t "AutoStartOff") }))
$menuAutoStart.Add_Click({
    if (Get-AutoStartStatus) {
        Uninstall-AutoStart
        $menuAutoStart.Text = (t "AutoStartOff")
    } else {
        Install-AutoStart
        $menuAutoStart.Text = (t "AutoStartOn")
    }
})
$menuExit = $trayMenu.Items.Add((t "CloseApp"))
$menuExit.Add_Click({ $window.Close() })
$script:NotifyIcon.ContextMenuStrip = $trayMenu
$script:NotifyIcon.Add_DoubleClick({
    if ($script:IsCompact) {
        $btnCompact.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    }
    $window.ShowInTaskbar = $true
    $window.Show()
    $window.WindowState = [System.Windows.WindowState]::Normal
    $window.Topmost = $true
    $window.Activate()
})

function Update-TrayTooltip([string]$Text) {
    try { if($Text.Length -gt 63){$Text=$Text.Substring(0,63)}; $script:NotifyIcon.Text=$Text } catch {}
}

# ======================== START ========================

# Timer do odswiezania i odliczania
$script:Timer=New-Object System.Windows.Threading.DispatcherTimer
$script:Timer.Interval=[TimeSpan]::FromSeconds(1)
$script:SecondsLeft = $script:Config.Interval

# Timer wymuszajacy HWND_TOPMOST przez Win32 API
$script:HWND_TOPMOST   = [IntPtr]::new(-1)
$script:HWND_NOTOPMOST = [IntPtr]::new(-2)
$script:SWP_FLAGS = 0x0002 -bor 0x0001  # SWP_NOMOVE | SWP_NOSIZE
$script:ForceTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:ForceTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:ForceTimer.Add_Tick({
    if ($script:IsCompact -and $script:CompactTopMost) {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
        [Native.Win32]::SetWindowPos($hwnd, $script:HWND_TOPMOST, 0, 0, 0, 0, $script:SWP_FLAGS) | Out-Null
    }
})

$script:Timer.Add_Tick({
    $script:SecondsLeft--
    # Sprawdz backoff 429
    if ($script:BackoffUntil -and (Get-Date) -lt $script:BackoffUntil) {
        $bSecs = [Math]::Max(0, [Math]::Ceiling(($script:BackoffUntil - (Get-Date)).TotalSeconds))
        $bMins = [Math]::Ceiling($bSecs / 60)
        $txtNextUpdate.Text = "429 backoff: ${bMins}min"
        $txtStatus.Text = (t "TooMany")
        $txtStatus.Foreground = [System.Windows.Media.Brushes]::OrangeRed
        if ($script:SecondsLeft -le 0) { $script:SecondsLeft = $script:Config.Interval }
        return
    }
    if ($script:BackoffUntil) { $script:BackoffUntil = $null }  # backoff wygasl - wyczysc
    if ($script:SecondsLeft -le 0) {
        $script:SecondsLeft = $script:Config.Interval
        Update-Display
    } else {
        $txtNextUpdate.Text = "$(t 'NextUpdate') $($script:SecondsLeft)s"
        Update-ReadingAge   # odswiez "X min temu" co sekunde
    }
})

$btnRefresh.Add_Click({
    $script:SecondsLeft = $script:Config.Interval
    Update-Display
})

$btnUnit.Add_Click({
    $script:UseMgDl = -not $script:UseMgDl
    Render-GlucoseUI
    if ($script:HistWin -and $script:HistWin.IsLoaded) { Render-HistGraph $script:HistDays }
    Save-Config
})

$btnHist.Add_Click({ Show-HistoryWindow })

function Update-HistLabels {
    if (-not ($script:HistWin -and $script:HistWin.IsLoaded)) { return }
    if ($script:HistLblTitleTxt) { $script:HistLblTitleTxt.Text = (t "HistTitle") }
    if ($script:HistLblAvg)      { $script:HistLblAvg.Text      = (t "HistAvg") }
    if ($script:HistLblTIR)      { $script:HistLblTIR.Text      = (t "HistTIR") }
    if ($script:HistNoData)      { $script:HistNoData.Text       = (t "HistNoData") }
    if ($script:HistBtns) {
        $script:HistBtns[0].Content = "7 "  + (t "HistDays")
        $script:HistBtns[1].Content = "14 " + (t "HistDays")
        $script:HistBtns[2].Content = "30 " + (t "HistDays")
        $script:HistBtns[3].Content = "90 " + (t "HistDays")
    }
}

function Apply-Language {
    # Tray menu
    $menuShow.Text      = (t "ShowWin")
    $menuSettings.Text  = (t "SettingsMenu")
    $menuBackup.Text    = (t "BackupMenu")
    $menuRestore.Text   = (t "RestoreMenu")
    $menuLogout.Text    = (t "SwitchAcc")
    $menuAutoStart.Text = if (Get-AutoStartStatus) { (t "AutoStartOn") } else { (t "AutoStartOff") }
    $menuExit.Text      = (t "CloseApp")
    # Etykiety statyczne
    $lblSred.Text    = (t "AvgLbl")
    # Tooltips i etykiety przyciskow
    $btnUnit.ToolTip    = (t "UnitTip")
    $btnRefresh.Content = (t "RefreshBtn")
    $btnHist.ToolTip    = (t "HistTitle")
    # Przycisk jezyka: pokazuje jezyk DO KTOREGO mozna przelczyc
    $btnLang.Content = if ($script:LangEn) { "PL" } else { "ENG" }
    # Odswiez wszystkie dynamiczne teksty z cache
    if ($null -ne $script:CachedMgDl) { Render-GlucoseUI }
    # Status bar jesli brak danych
    if ($txtGlucoseValue.Text -eq "---") {
        if ($txtStatus.Text -ne (t "Fetching")) {
            $txtStatus.Text = (t "NoData")
        }
    }
    # Odswiez okno historii jesli otwarte
    Update-HistLabels
    if ($script:HistWin -and $script:HistWin.IsLoaded) { Render-HistGraph $script:HistDays }
}

$btnLang.Add_Click({
    $script:LangEn = -not $script:LangEn
    Apply-Language
    Save-Config
})

$btnSmooth.Add_Click({
    $script:SmoothMode = -not $script:SmoothMode
    $col = if ($script:SmoothMode) { [System.Windows.Media.Brushes]::White } else { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#6677aa")) }
    $bg  = if ($script:SmoothMode) { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#2a3a5a")) } else { New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#1e2a3a")) }
    $btnSmooth.Foreground = $col; $btnSmooth.Background = $bg
    if ($script:HistBtnSmooth) { $script:HistBtnSmooth.Foreground = $col; $script:HistBtnSmooth.Background = $bg }
    Save-Config
    if ($script:CachedGraphData) { Update-Graph $script:CachedGraphData }
    if ($script:HistWin -and $script:HistWin.IsLoaded) { Render-HistGraph $script:HistDays }
})

$window.Add_ContentRendered({
    Update-Display
    $txtNextUpdate.Text = "$(t 'NextUpdate') $($script:SecondsLeft)s"
    $script:Timer.Start()
    # Zastosuj wczytany stan SmoothMode do przycisku ~
    if ($script:SmoothMode) {
        $btnSmooth.Foreground = [System.Windows.Media.Brushes]::White
        $btnSmooth.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#2a3a5a"))
    }
    # Jesli uruchomiony przez Task Scheduler (autostart) i dane logowania juz istnieja - schowaj do tray
    # Jesli brak konfiguracji (pierwsze uruchomienie) - zostaw okno widoczne
    if ($AutoStart -and $configLoaded) {
        $window.Hide()
    }
})

$window.Add_Closed({
    # Zapisz aktualna pozycje okna przed zamknieciem
    if ($script:IsCompact) {
        $script:CompactLeft = $window.Left
        $script:CompactTop  = $window.Top
    } else {
        $script:FullLeft = $window.Left
        $script:FullTop  = $window.Top
    }
    try { Save-Config } catch {}
    if($script:Timer){$script:Timer.Stop()}
    if($script:ForceTimer){$script:ForceTimer.Stop()}
    $script:NotifyIcon.Visible=$false; $script:NotifyIcon.Dispose()
    [Native.Win32]::ShowWindow($consoleHwnd, 5)|Out-Null
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()
})

Write-Log "=== Glucose Monitor v5 START ==="

# Wczytaj zapisane dane lub pokaz okno logowania
$configLoaded = Load-Config
if ($configLoaded) {
    # Weryfikuj dane z serwerem
    $loginOk = Invoke-LibreLogin
    if (-not $loginOk) {
        # Dane bledne lub nieaktualne - pokaz okno logowania
        $script:Config.Email = ""; $script:Config.Password = ""
        $configLoaded = $false
    }
}
if (-not $configLoaded) {
    $result = Show-LoginWindow
    if (-not $result) { Write-Log "Anulowano logowanie"; exit }
}

# Automatyczna rejestracja w harmonogramie zadan przy pierwszym uruchomieniu
if (-not (Get-AutoStartStatus)) {
    Write-Log "AutoStart: zadanie nie istnieje, rejestruje..."
    Install-AutoStart -Silent
}

# Przywroc zapamietana pozycje okna glownego (jesli zapisana w konfiguracji)
if ($null -ne $script:FullLeft -and $null -ne $script:FullTop) {
    $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
    $window.Left = $script:FullLeft
    $window.Top  = $script:FullTop
}
if ($null -ne $script:WindowOpacity -and $script:WindowOpacity -gt 0) {
    $window.Opacity = $script:WindowOpacity
}

$window.Show()
[System.Windows.Threading.Dispatcher]::Run()
Write-Log "=== STOP ==="