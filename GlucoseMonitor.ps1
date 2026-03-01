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
}
function t([string]$key) { $script:T[$key][[int]$script:LangEn] }

$script:HistoryFile = Join-Path $script:ScriptDir "history.jsonl"

$script:HistKnownTs = $null  # HashSet znanych timestampow (yyyy-MM-ddTHH:mm) - inicjowany przy pierwszym uzyciu

function Save-HistoryEntry([double]$mgdl, [int]$trend) {
    try {
        $now = Get-Date
        if ($script:LastHistorySave -and ($now - $script:LastHistorySave).TotalMinutes -lt 2) { return }
        $script:LastHistorySave = $now
        $key = $now.ToString("yyyy-MM-ddTHH:mm")
        if ($script:HistKnownTs) { $script:HistKnownTs.Add($key) | Out-Null }
        $entry = '{"ts":"' + $now.ToString("yyyy-MM-ddTHH:mm:ss") + '","mgdl":' + [Math]::Round($mgdl,1) + ',"trend":' + $trend + '}'
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

function Load-HistoryData([int]$days, [int]$offset = 0) {
    $result    = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path $script:HistoryFile)) { return $result }
    $endDate   = (Get-Date).AddDays(-$offset)
    $startDate = $endDate.AddDays(-$days)
    try {
        Get-Content $script:HistoryFile -Encoding UTF8 | ForEach-Object {
            try {
                $obj = $_ | ConvertFrom-Json
                $ts  = [DateTime]::Parse($obj.ts)
                if ($ts -ge $startDate -and $ts -le $endDate) { $result.Add($obj) }
            } catch {}
        }
    } catch {}
    return ($result | Sort-Object { [DateTime]::Parse($_.ts) })
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

# ======================== HELPERS ========================
# TrendArrow: 0=NotDetermined, 1=FallingQuickly, 2=Falling, 3=Stable, 4=Rising, 5=RisingQuickly
function Get-TrendArrow([int]$v){switch($v){1{[char]0x2193+[char]0x2193}2{[char]0x2193}3{[char]0x2192}4{[char]0x2197}5{[char]0x2191+[char]0x2191}default{"?"}}}
function Get-TrendText([int]$v){
    $li=[int]$script:LangEn
    switch($v){
        1{@("Szybki spadek","Rapidly falling")[$li]}  2{@("Spadek","Falling")[$li]}
        3{@("Stabilny","Stable")[$li]}                4{@("Wzrost","Rising")[$li]}
        5{@("Szybki wzrost","Rapidly rising")[$li]}   default{"---"}
    }
}
function Get-CalculatedTrend([array]$graphData) {
    if (-not $graphData -or $graphData.Count -lt 3) { return $null }
    $pts = $graphData | Select-Object -Last 5
    $deltas = @()
    for ($i = 1; $i -lt $pts.Count; $i++) {
        $mg0 = try { [double]$pts[$i-1].ValueInMgPerDl } catch { 0 }
        $mg1 = try { [double]$pts[$i].ValueInMgPerDl   } catch { 0 }
        if ($mg0 -le 0 -or $mg1 -le 0) { continue }
        $tsRaw0 = if ($pts[$i-1].Timestamp) { $pts[$i-1].Timestamp } else { $pts[$i-1].FactoryTimestamp }
        $tsRaw1 = if ($pts[$i].Timestamp)   { $pts[$i].Timestamp   } else { $pts[$i].FactoryTimestamp   }
        try {
            $t0   = [DateTime]::Parse([string]$tsRaw0)
            $t1   = [DateTime]::Parse([string]$tsRaw1)
            $mins = ($t1 - $t0).TotalMinutes
            if ($mins -gt 0.5) { $deltas += ($mg1 - $mg0) / $mins }
        } catch {}
    }
    if ($deltas.Count -eq 0) { return $null }
    $slope = ($deltas | Measure-Object -Average).Average
    if    ($slope -lt -2)   { return 1 }
    elseif($slope -lt -1)   { return 2 }
    elseif($slope -le  1)   { return 3 }
    elseif($slope -le  2)   { return 4 }
    else                    { return 5 }
}
function Get-TrendColor([int]$v) {
    switch ($v) {
        1 { "#FF3333" }   # Szybki spadek  - czerwony
        2 { "#FF8800" }   # Spadek         - pomaranczowy
        3 { "#44DD44" }   # Stabilny       - zielony
        4 { "#FFAA00" }   # Wzrost         - zolty
        5 { "#FF3333" }   # Szybki wzrost  - czerwony
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
    @(
        "Email=$($script:Config.Email)"
        "EncryptedPassword=$encPass"
        "Interval=$($script:Config.Interval)"
        "AlertLow=$($script:Config.AlertLow)"
        "AlertHigh=$($script:Config.AlertHigh)"
        "LangEn=$($script:LangEn)"
        "UseMgDl=$($script:UseMgDl)"
        "SmoothMode=$($script:SmoothMode)"
    ) | Set-Content -Path $script:ConfigFile -Encoding UTF8
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
                        <TextBlock Name="txtUnitLabel" Text="mmol/L" Foreground="#7777aa" FontSize="10" Margin="2,0,0,0"/>
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

            <!-- Trend + stats -->
            <Border Grid.Row="3" CornerRadius="8" Padding="10,6" Background="#2a2a4a" Margin="12,0,12,4">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <TextBlock Name="lblTrend" Text="Trend" Style="{StaticResource L}"/>
                        <TextBlock Name="txtTrendText" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Min" Style="{StaticResource L}"/>
                            <TextBlock Text=" 12h" Foreground="#8888bb" FontSize="9" VerticalAlignment="Bottom" Margin="2,0,0,1" FontFamily="Segoe UI"/>
                        </StackPanel>
                        <TextBlock Name="txtMin" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Name="lblSred" Text="Sred." Style="{StaticResource L}"/>
                            <TextBlock Text=" 12h" Foreground="#8888bb" FontSize="9" VerticalAlignment="Bottom" Margin="2,0,0,1" FontFamily="Segoe UI"/>
                        </StackPanel>
                        <TextBlock Name="txtAvg" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="3" HorizontalAlignment="Center">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Max" Style="{StaticResource L}"/>
                            <TextBlock Text=" 12h" Foreground="#8888bb" FontSize="9" VerticalAlignment="Bottom" Margin="2,0,0,1" FontFamily="Segoe UI"/>
                        </StackPanel>
                        <TextBlock Name="txtMax" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="4" HorizontalAlignment="Right">
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
$lblTrend=$window.FindName("lblTrend")
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
        @{ lo=$zHiC; hi=$mx;   col="#44FF3333" }
        @{ lo=$zHiH; hi=$zHiC; col="#33FFAA00" }
        @{ lo=$zLoY; hi=$zHiH; col="#2200CC44" }
        @{ lo=$zLoN; hi=$zLoY; col="#44FFEE00" }
        @{ lo=$mn;   hi=$zLoN; col="#44FF3333" }
    )) {
        $zHi = [Math]::Min($z.hi, $mx); $zLo = [Math]::Max($z.lo, $mn)
        if ($zHi -le $zLo) { continue }
        $yT = $m + $dh - (($dh/$rng)*($zHi - $mn))
        $yB = $m + $dh - (($dh/$rng)*($zLo - $mn))
        $zRect = New-Object System.Windows.Shapes.Rectangle
        $zRect.Fill = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($z.col))
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
            $ln.Stroke=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(60,180,180,180))
            $ln.StrokeThickness=0.7
            $canvasGraph.Children.Add($ln)|Out-Null
            $tb=New-Object System.Windows.Controls.TextBlock; $tb.Text="$lim"; $tb.FontSize=8
            $tb.Foreground=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(180,180,180,180))
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
    $mGeoR = New-Object System.Windows.Media.PathGeometry
    $mGeoO = New-Object System.Windows.Media.PathGeometry
    $mGeoG = New-Object System.Windows.Media.PathGeometry
    [double]$tens = 0.2
    $mCurFig = $null; $mCurCol = ""
    for ($i = 0; $i -lt $n - 1; $i++) {
        $avg2 = ($vals[$i]+$vals[$i+1])/2.0
        $sc   = if ($avg2 -lt $loC -or $avg2 -gt $hiC) { "R" } elseif ($avg2 -gt $hiH) { "O" } else { "G" }
        if ($sc -ne $mCurCol) {
            if ($mCurFig) {
                if     ($mCurCol -eq "R") { $mGeoR.Figures.Add($mCurFig) | Out-Null }
                elseif ($mCurCol -eq "O") { $mGeoO.Figures.Add($mCurFig) | Out-Null }
                else                      { $mGeoG.Figures.Add($mCurFig) | Out-Null }
            }
            $mCurFig = New-Object System.Windows.Media.PathFigure
            $mCurFig.StartPoint = [System.Windows.Point]::new($px[$i], $py[$i])
            $mCurCol = $sc
        }
        $i0 = if ($i -gt 0) { $i-1 } else { 0 }
        $i3 = if ($i+2 -lt $n) { $i+2 } else { $n-1 }
        [double]$cp1x = $px[$i]   + ($px[$i+1]-$px[$i0])*$tens
        [double]$cp1y = $py[$i]   + ($py[$i+1]-$py[$i0])*$tens
        [double]$cp2x = $px[$i+1] - ($px[$i3] -$px[$i]) *$tens
        [double]$cp2y = $py[$i+1] - ($py[$i3] -$py[$i]) *$tens
        $bz = New-Object System.Windows.Media.BezierSegment
        $bz.Point1 = [System.Windows.Point]::new($cp1x, $cp1y)
        $bz.Point2 = [System.Windows.Point]::new($cp2x, $cp2y)
        $bz.Point3 = [System.Windows.Point]::new($px[$i+1], $py[$i+1])
        $bz.IsStroked = $true
        $mCurFig.Segments.Add($bz) | Out-Null
    }
    if ($mCurFig) {
        if     ($mCurCol -eq "R") { $mGeoR.Figures.Add($mCurFig) | Out-Null }
        elseif ($mCurCol -eq "O") { $mGeoO.Figures.Add($mCurFig) | Out-Null }
        else                      { $mGeoG.Figures.Add($mCurFig) | Out-Null }
    }
    foreach ($item in @(@{g=$mGeoR;c="#EE4444"},@{g=$mGeoO;c="#FFAA44"},@{g=$mGeoG;c="#44DDAA"})) {
        if ($item.g.Figures.Count -eq 0) { continue }
        $pe = New-Object System.Windows.Shapes.Path; $pe.Data = $item.g
        $pe.Stroke = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($item.c))
        $pe.StrokeThickness=2; $pe.StrokeStartLineCap="Round"; $pe.StrokeEndLineCap="Round"
        $canvasGraph.Children.Add($pe)|Out-Null
    }

    # Kropki skanow (type=1) - male polprzezroczyste kola na wierzchu linii
    for($i=0; $i -lt $vals.Count; $i++) {
        if ($types[$i]) {
            $sx=$m+($i*$step); $sy=$m+$dh-(($dh/$rng)*($vals[$i]-$mn))
            $sc=New-Object System.Windows.Shapes.Ellipse; $sc.Width=5; $sc.Height=5
            $sc.Fill=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(210,255,255,255))
            $sc.Stroke=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(120,255,255,255))
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
    $dotCol = if ($lastAvg -lt $loC -or $lastAvg -gt $hiC) { "#EE4444" } elseif ($lastAvg -gt $hiH) { "#FFAA44" } else { "#44DDAA" }
    $dot=New-Object System.Windows.Shapes.Ellipse; $dot.Width=8;$dot.Height=8
    $dot.Fill=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($dotCol))
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
        $fCol  = if ($f2Val -lt $loC -or $f2Val -gt $hiC) { "#EE4444" } elseif ($f2Val -gt $hiH) { "#FFAA44" } else { "#44DDAA" }
        $fl=New-Object System.Windows.Shapes.Line; $fl.X1=$lastX;$fl.Y1=$lastY;$fl.X2=$f2X;$fl.Y2=$f2Y
        $fl.Stroke=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($fCol))
        $fl.StrokeThickness=1.5; $fl.Opacity=0.55
        $da=New-Object System.Windows.Media.DoubleCollection; $da.Add(4);$da.Add(3); $fl.StrokeDashArray=$da
        $canvasGraph.Children.Add($fl)|Out-Null
        $fdot=New-Object System.Windows.Shapes.Ellipse; $fdot.Width=6;$fdot.Height=6; $fdot.Opacity=0.55
        $fdot.Fill=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($fCol))
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
                        $tb2.Foreground=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(120,150,200,200))
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
        }

        # Czulszy algorytm trendu - oblicz z danych wykresu
        $calcTrend = Get-CalculatedTrend $script:CachedGraphData
        if ($null -ne $calcTrend) { $script:CachedTrend = $calcTrend }

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

# ---- Render-HistGraph na poziomie skryptu (nie zagniezdzona!) ----
function Render-HistGraph([int]$days) {
    if (-not $script:HistCanvas) { return }
    try {
        # Podswietl aktywny przycisk okresu
        if ($script:HistBtns) {
            $dayMap = @(7,14,30,90)
            for ($bi=0; $bi -lt 4; $bi++) {
                $b = $script:HistBtns[$bi]
                if (-not $b) { continue }
                if ($dayMap[$bi] -eq $days) {
                    $b.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#2a3a5a"))
                    $b.Foreground = [System.Windows.Media.Brushes]::White
                } else {
                    $b.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#1e2a3a"))
                    $b.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#7777aa"))
                }
            }
        }

        $data = Load-HistoryData $days $script:HistOffset
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

        # Zbierz wartosci glukozy i oblicz delta
        $vals   = [System.Collections.Generic.List[double]]::new()
        $tss    = [System.Collections.Generic.List[datetime]]::new()
        $deltas = [System.Collections.Generic.List[double]]::new()
        $prevMg = 0.0
        foreach ($pt in $data) {
            $mg = try { [double]$pt.mgdl } catch { 0 }
            if ($mg -le 20) { continue }
            $ptTs = try { [DateTime]::Parse($pt.ts) } catch { [DateTime]::MinValue }
            if ($ptTs -eq [DateTime]::MinValue) { continue }
            $displayV = if ($script:UseMgDl) { [Math]::Round($mg,0) } else { [Math]::Round($mg/18.018,1) }
            $vals.Add($displayV)
            $tss.Add($ptTs)
            if ($prevMg -gt 20) { $deltas.Add([Math]::Abs($mg - $prevMg)) }
            $prevMg = $mg
        }
        if ($vals.Count -lt 2) {
            if ($script:HistNoData) { $script:HistNoData.Visibility = [System.Windows.Visibility]::Visible }
            return
        }

        # Statystyki
        $sm  = $vals | Measure-Object -Min -Max -Average
        $fmt = if ($script:UseMgDl) { "0" } else { "0.0" }
        $loNorm = if ($script:UseMgDl) { 70.0  } else { 3.9  }
        $hiNorm = if ($script:UseMgDl) { 180.0 } else { 10.0 }
        $inR = ($vals | Where-Object { $_ -ge $loNorm -and $_ -le $hiNorm }).Count
        $tirPct = [Math]::Round($inR / $vals.Count * 100, 0)

        if ($script:HistValAvg) {
            $script:HistValAvg.Text = [Math]::Round($sm.Average,1).ToString($fmt)
            $script:HistValMin.Text = $sm.Minimum.ToString($fmt)
            $script:HistValMax.Text = $sm.Maximum.ToString($fmt)
            $script:HistValTIR.Text = "$tirPct%"
            # eHbA1c - formula NGSP: (srednia_mgdl + 46.7) / 28.7
            if ($script:HistValHbA1c) {
                $avgMgDl = if ($script:UseMgDl) { $sm.Average } else { $sm.Average * 18.018 }
                $eHbA1c  = [Math]::Round(($avgMgDl + 46.7) / 28.7, 1)
                $script:HistValHbA1c.Text = "$($eHbA1c.ToString('0.0'))%"
            }
            # SD i CV%
            if ($vals.Count -gt 1) {
                $mean = $sm.Average
                $sd   = [Math]::Sqrt(($vals | ForEach-Object { ($_ - $mean)*($_ - $mean) } | Measure-Object -Sum).Sum / $vals.Count)
                $cv   = if ($mean -gt 0) { [Math]::Round($sd / $mean * 100, 0) } else { 0 }
                $sdDisp = if ($script:UseMgDl) { [Math]::Round($sd,0).ToString("0") } else { [Math]::Round($sd,1).ToString("0.0") }
                if ($script:HistValSD)  { $script:HistValSD.Text  = $sdDisp }
                if ($script:HistValCV)  { $script:HistValCV.Text  = "$cv%" }
            }
            if ($deltas.Count -gt 0) {
                $avgDelta     = ($deltas | Measure-Object -Average).Average
                $avgDeltaDisp = if ($script:UseMgDl) { [Math]::Round($avgDelta,0) } else { [Math]::Round($avgDelta/18.018,1) }
                $script:HistValDelta.Text = $avgDeltaDisp.ToString($fmt)
            } else { $script:HistValDelta.Text = "---" }
        }

        # Wygładzanie Savitzky-Golay tylko dla wykresu (statystyki z oryginalnych danych)
        # $chartVals = tablica do rysowania (moze byc wyglądzona); $vals pozostaje bez zmian (statystyki)
        $chartVals = if ($script:SmoothMode -and $vals.Count -ge 5) {
            Apply-SavitzkyGolay ($vals.ToArray())
        } else {
            $vals.ToArray()
        }

        # Wymiary canvas
        if ($script:HistWin) { $script:HistWin.UpdateLayout() }
        $cW = $script:HistCanvas.ActualWidth;  if ($cW -lt 10) { $cW = 440.0 }
        $cH = $script:HistCanvas.ActualHeight; if ($cH -lt 10) { $cH = 250.0 }
        $padL=38.0; $padR=8.0; $padT=8.0; $padB=22.0
        $gW = $cW - $padL - $padR
        $gH = $cH - $padT  - $padB
        $n  = $chartVals.Length

        # Zakres czasu osi X (z poprawnych punktow - zsynchronizowany z linia danych)
        $firstTs   = $tss[0]
        $lastTs    = $tss[$n - 1]
        $totalSecs = [Math]::Max(1.0, ($lastTs - $firstTs).TotalSeconds)

        # Zakresy osi Y
        $loY = if ($script:UseMgDl) { 40.0 } else { 2.2 }
        $hiY = if ($script:UseMgDl) { 280.0} else { 15.5}
        if ($sm.Minimum -lt $loY) { $loY = [Math]::Floor($sm.Minimum  - 0.5) }
        if ($sm.Maximum -gt $hiY) { $hiY = [Math]::Ceiling($sm.Maximum + 0.5) }
        $rangeY = $hiY - $loY; if ($rangeY -lt 0.1) { $rangeY = 1.0 }

        # Pomocnicze obliczenia (bez zagniezdzonej funkcji!)
        # Y: $padT + $gH - ($v - $loY)/$rangeY * $gH
        # X: $padL + ($tss[$i] - $firstTs).TotalSeconds / $totalSecs * $gW

        # --- Kolorowe tlo stref ---
        $hiHyper  = if ($script:UseMgDl) { 250.0 } else { 13.9 }
        $hiYellow = if ($script:UseMgDl) { 180.0 } else { 10.0 }
        $loNormY  = $loNorm
        $loYellowZ = if ($script:UseMgDl) { 79.0 } else { 4.4 }
        $zones = @(
            @{ yTop=$hiHyper;  yBot=$hiY;        col="#44FF3333" }
            @{ yTop=$hiYellow; yBot=$hiHyper;    col="#33FFAA00" }
            @{ yTop=$loYellowZ;yBot=$hiYellow;   col="#2200CC44" }
            @{ yTop=$loNormY;  yBot=$loYellowZ;  col="#44FFEE00" }
            @{ yTop=$loY;      yBot=$loNormY;    col="#44FF3333" }
        )
        foreach ($z in $zones) {
            $bot = $z.yBot; $top = $z.yTop
            if ($bot -le $loY -or $top -ge $hiY) { continue }
            $yT = $padT + $gH - ([Math]::Min($bot,$hiY) - $loY)/$rangeY * $gH
            $yB = $padT + $gH - ([Math]::Max($top,$loY) - $loY)/$rangeY * $gH
            $h  = [Math]::Abs($yB - $yT); if ($h -lt 0.5) { continue }
            $yTop2 = [Math]::Min($yT,$yB)
            $rect = New-Object System.Windows.Shapes.Rectangle
            $rect.Fill   = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($z.col))
            $rect.Width  = $gW; $rect.Height = $h
            [System.Windows.Controls.Canvas]::SetLeft($rect, $padL)
            [System.Windows.Controls.Canvas]::SetTop($rect,  [Math]::Max($padT,$yTop2))
            $script:HistCanvas.Children.Add($rect) | Out-Null
        }

        # --- Linie siatki + etykiety Y ---
        $gridVals = if ($script:UseMgDl) { @(70,100,140,180,250) } else { @(3.9,5.5,7.0,10.0,13.9) }
        foreach ($gVal in $gridVals) {
            if ($gVal -lt $loY -or $gVal -gt $hiY) { continue }
            $gy = $padT + $gH - ($gVal - $loY)/$rangeY * $gH
            $gl = New-Object System.Windows.Shapes.Line
            $gl.X1=$padL; $gl.X2=$padL+$gW; $gl.Y1=$gy; $gl.Y2=$gy
            $gl.Stroke = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#33FFFFFF"))
            $gl.StrokeThickness = 0.7
            $script:HistCanvas.Children.Add($gl) | Out-Null
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = $gVal.ToString($fmt)
            $lbl.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#7777aa"))
            $lbl.FontSize = 8; $lbl.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
            [System.Windows.Controls.Canvas]::SetLeft($lbl, 1)
            [System.Windows.Controls.Canvas]::SetTop($lbl, $gy - 7)
            $script:HistCanvas.Children.Add($lbl) | Out-Null
        }

        # --- Etykiety osi X (daty) ---
        # firstTs / lastTs / totalSecs obliczone wyzej z poprawnych punktow
        if ($firstTs -and $lastTs) {
            $step = if ($days -le 7) { 1 } elseif ($days -le 14) { 2 } elseif ($days -le 30) { 5 } else { 14 }
            $cur  = $firstTs.Date
            while ($cur -le $lastTs.Date) {
                $frac = ($cur - $firstTs).TotalSeconds / $totalSecs
                $xPos = $padL + $frac * $gW
                $dl   = New-Object System.Windows.Controls.TextBlock
                $dl.Text = $cur.ToString("dd.MM")
                $dl.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#7777aa"))
                $dl.FontSize = 8; $dl.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
                [System.Windows.Controls.Canvas]::SetLeft($dl, $xPos - 10)
                [System.Windows.Controls.Canvas]::SetTop($dl,  $cH - $padB + 4)
                $script:HistCanvas.Children.Add($dl) | Out-Null
                $cur = $cur.AddDays($step)
            }
        }

        # --- Linia danych - wygladzona Catmull-Rom -> cubic Bezier ---
        # 3 PathGeometry (po jednym na kolor) zamiast 2000 osobnych Path - duzo wydajniej
        $hiH2 = if ($script:UseMgDl) { 180.0 } else { 10.0 }
        $hiC2 = if ($script:UseMgDl) { 250.0 } else { 13.9 }
        $lC2  = if ($script:UseMgDl) {  70.0 } else {  3.9 }
        $hpx = [double[]]::new($n); $hpy = [double[]]::new($n)
        for ($i = 0; $i -lt $n; $i++) {
            $hpx[$i] = $padL + ($tss[$i] - $firstTs).TotalSeconds / $totalSecs * $gW
            $hpy[$i] = $padT + $gH - ($chartVals[$i] - $loY) / $rangeY * $gH
        }
        $geoR = New-Object System.Windows.Media.PathGeometry
        $geoO = New-Object System.Windows.Media.PathGeometry
        $geoG = New-Object System.Windows.Media.PathGeometry
        [double]$tens = 0.2
        $curFig = $null; $curCol = ""
        for ($i = 0; $i -lt ($n - 1); $i++) {
            $isGap = ($tss[$i+1] - $tss[$i]).TotalMinutes -gt 30
            $avg2  = ($chartVals[$i]+$chartVals[$i+1])/2.0
            $sc    = if ($avg2 -lt $lC2 -or $avg2 -gt $hiC2) { "R" } elseif ($avg2 -gt $hiH2) { "O" } else { "G" }
            if ($isGap -or $sc -ne $curCol) {
                if ($curFig) {
                    if     ($curCol -eq "R") { $geoR.Figures.Add($curFig) | Out-Null }
                    elseif ($curCol -eq "O") { $geoO.Figures.Add($curFig) | Out-Null }
                    else                     { $geoG.Figures.Add($curFig) | Out-Null }
                }
                if ($isGap) { $curFig = $null; $curCol = ""; continue }
                $curFig = New-Object System.Windows.Media.PathFigure
                $curFig.StartPoint = [System.Windows.Point]::new($hpx[$i], $hpy[$i])
                $curCol = $sc
            }
            $i0 = if ($i -gt 0 -and ($tss[$i] - $tss[$i-1]).TotalMinutes -le 30) { $i-1 } else { $i }
            $i3 = if ($i+2 -lt $n -and ($tss[$i+2] - $tss[$i+1]).TotalMinutes -le 30) { $i+2 } else { $i+1 }
            [double]$cp1x = $hpx[$i]   + ($hpx[$i+1]-$hpx[$i0])*$tens
            [double]$cp1y = $hpy[$i]   + ($hpy[$i+1]-$hpy[$i0])*$tens
            [double]$cp2x = $hpx[$i+1] - ($hpx[$i3] -$hpx[$i]) *$tens
            [double]$cp2y = $hpy[$i+1] - ($hpy[$i3] -$hpy[$i]) *$tens
            $bz = New-Object System.Windows.Media.BezierSegment
            $bz.Point1 = [System.Windows.Point]::new($cp1x, $cp1y)
            $bz.Point2 = [System.Windows.Point]::new($cp2x, $cp2y)
            $bz.Point3 = [System.Windows.Point]::new($hpx[$i+1], $hpy[$i+1])
            $bz.IsStroked = $true
            $curFig.Segments.Add($bz) | Out-Null
        }
        if ($curFig) {
            if     ($curCol -eq "R") { $geoR.Figures.Add($curFig) | Out-Null }
            elseif ($curCol -eq "O") { $geoO.Figures.Add($curFig) | Out-Null }
            else                     { $geoG.Figures.Add($curFig) | Out-Null }
        }
        foreach ($item in @(@{g=$geoR;c="#EE4444"},@{g=$geoO;c="#FFAA44"},@{g=$geoG;c="#44DDAA"})) {
            if ($item.g.Figures.Count -eq 0) { continue }
            $pe = New-Object System.Windows.Shapes.Path; $pe.Data = $item.g
            $pe.Stroke = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($item.c))
            $pe.StrokeThickness = 1.5
            $script:HistCanvas.Children.Add($pe) | Out-Null
        }
    } catch { Write-Log "Render-HistGraph err: $($_.Exception.Message)" }
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
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Name="hTitleTxt" Text="  Historia glukozy"
                               Foreground="#7777aa" FontSize="11" VerticalAlignment="Center"
                               FontFamily="Segoe UI" Margin="6,0,0,0"/>
                    <Button Grid.Column="1" Name="hBtnAgp" Content="AGP" Width="42" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Wzorzec dobowy (AGP)"/>
                    <Button Grid.Column="2" Name="hBtnBars" Content="BARS" Width="42" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Średnie stężenie glukozy (słupkowe) / Average glucose (bars)"/>
                    <Button Grid.Column="3" Name="hBtnReport" Content="HTML" Width="42" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Eksportuj raport HTML"/>
                    <Button Grid.Column="4" Name="hBtnPdf" Content="PDF" Width="42" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Eksportuj raport PDF"/>
                    <Button Grid.Column="5" Name="hBtnCsv" Content="CSV" Width="40" Height="30"
                            Background="Transparent" Foreground="#6677aa" BorderThickness="0"
                            FontSize="9" Cursor="Hand" ToolTip="Eksportuj dane do pliku CSV"/>
                    <Button Grid.Column="6" Name="hBtnSmooth" Content="~" Width="34" Height="30"
                            Background="#1e2a3a" Foreground="#6677aa" BorderThickness="0"
                            FontSize="13" Cursor="Hand" ToolTip="Wygładzanie danych (Savitzky-Golay) / Data smoothing"/>
                    <Button Grid.Column="7" Name="hClose" Content="&#x2715;" Width="34" Height="30"
                            Background="Transparent" Foreground="#aa5555" BorderThickness="0"
                            FontSize="13" Cursor="Hand"/>
                </Grid>
            </Border>

            <Grid Grid.Row="1" Margin="12,8,12,4">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button Grid.Column="0" Name="hBtn7"  Content="7 dni"  Background="#2a3a5a" Foreground="White"
                        BorderThickness="0" Padding="0,6" FontSize="11" Cursor="Hand" FontFamily="Segoe UI"/>
                <Button Grid.Column="2" Name="hBtn14" Content="14 dni" Background="#1e2a3a" Foreground="#7777aa"
                        BorderThickness="0" Padding="0,6" FontSize="11" Cursor="Hand" FontFamily="Segoe UI"/>
                <Button Grid.Column="4" Name="hBtn30" Content="30 dni" Background="#1e2a3a" Foreground="#7777aa"
                        BorderThickness="0" Padding="0,6" FontSize="11" Cursor="Hand" FontFamily="Segoe UI"/>
                <Button Grid.Column="6" Name="hBtn90" Content="90 dni" Background="#1e2a3a" Foreground="#7777aa"
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
    # Stan wizualny przycisku smooth - zsynchronizowany z glownym oknem
    if ($script:SmoothMode) {
        $script:HistBtnSmooth.Foreground = [System.Windows.Media.Brushes]::White
        $script:HistBtnSmooth.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#2a3a5a"))
    }
    $script:HistBtns     = @(
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
    $script:HistBtns[0].Content = "7 "  + (t "HistDays")
    $script:HistBtns[1].Content = "14 " + (t "HistDays")
    $script:HistBtns[2].Content = "30 " + (t "HistDays")
    $script:HistBtns[3].Content = "90 " + (t "HistDays")

    # Drag paska tytuloweg
    $script:HistWin.FindName("hTitleBar").Add_MouseLeftButtonDown({ $script:HistWin.DragMove() })

    $script:HistDays   = 7
    $script:HistOffset = 0

    # Renderuj przy otwarciu (Add_ContentRendered wykonuje sie po Show())
    $script:HistWin.Add_ContentRendered({ Render-HistGraph $script:HistDays })

    # Handlery przyciskow okresu - reset offset przy zmianie okresu
    $script:HistWin.FindName("hClose").Add_Click({ $script:HistWin.Close() })
    $script:HistBtns[0].Add_Click({ $script:HistDays = 7;  $script:HistOffset = 0; Render-HistGraph 7  })
    $script:HistBtns[1].Add_Click({ $script:HistDays = 14; $script:HistOffset = 0; Render-HistGraph 14 })
    $script:HistBtns[2].Add_Click({ $script:HistDays = 30; $script:HistOffset = 0; Render-HistGraph 30 })
    $script:HistBtns[3].Add_Click({ $script:HistDays = 90; $script:HistOffset = 0; Render-HistGraph 90 })

    # Nawigacja wstecz / naprzod
    $script:HistWin.FindName("hBtnPrev").Add_Click({
        $script:HistOffset = [Math]::Max(0, $script:HistOffset - [int]($script:HistDays / 2))
        Render-HistGraph $script:HistDays
    })
    $script:HistWin.FindName("hBtnNext").Add_Click({
        $script:HistOffset += [int]($script:HistDays / 2)
        Render-HistGraph $script:HistDays
    })

    # Eksport CSV
    $script:HistWin.FindName("hBtnCsv").Add_Click({ Export-HistoryCSV })
    $script:HistWin.FindName("hBtnAgp").Add_Click({ Show-AgpWindow })
    $script:HistWin.FindName("hBtnBars").Add_Click({ Show-AvgBarsWindow })
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

    $script:HistWin.Show()

    } catch { Write-Log "Show-HistoryWindow error: $($_.Exception.Message)" }
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
            [void]$sb.Append("<text x='$($pL-3)' y='$($yp+3)' text-anchor='end' font-family='Arial' font-size='9' fill='#778'>$lbl</text>")
            $yv += $step
        }

        # Linie pionowe X
        [double]$totalDays = ($t1 - $t0).TotalDays
        if ($totalDays -gt 2) {
            $cur = $t0.Date.AddDays(1)
            while ($cur -le $t1) {
                [double]$xp = [Math]::Round($pL + ($cur - $t0).TotalMinutes / $tRng * $gW, 1)
                [void]$sb.Append("<line x1='$xp' y1='$pT' x2='$xp' y2='$($pT+$gH)' stroke='#dde' stroke-width='0.5'/>")
                [void]$sb.Append("<text x='$xp' y='$($pT+$gH+15)' text-anchor='middle' font-family='Arial' font-size='8' fill='#778'>$($cur.ToString('dd.MM'))</text>")
                $cur = $cur.AddDays(1)
            }
        } else {
            $cur = $t0.Date.AddHours([int]($t0.Hour / 4) * 4)
            while ($cur -le $t1) {
                if ($cur -ge $t0) {
                    [double]$xp = [Math]::Round($pL + ($cur - $t0).TotalMinutes / $tRng * $gW, 1)
                    [void]$sb.Append("<line x1='$xp' y1='$pT' x2='$xp' y2='$($pT+$gH)' stroke='#dde' stroke-width='0.5'/>")
                    [void]$sb.Append("<text x='$xp' y='$($pT+$gH+15)' text-anchor='middle' font-family='Arial' font-size='8' fill='#778'>$($cur.ToString('HH:mm'))</text>")
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
        [void]$sb.Append("<text x='$pL' y='$($pT-2)' font-family='Arial' font-size='8' fill='#778'>$Unit</text>")
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
            [void]$sb.Append("<text x='$($pL-3)' y='$($yp+3)' text-anchor='end' font-family='Arial' font-size='9' fill='#778'>$lbl</text>")
            $yv += $step
        }

        # Slupki per godzina + linia srednich (inline, bez scriptblokow)
        $avgPts = New-Object System.Collections.Generic.List[string]
        for ($hi = 0; $hi -lt 24; $hi++) {
            [double]$xC = [Math]::Round($pL + ($hi + 0.5) / 24 * $gW, 1)
            if ($hi % 3 -eq 0) {
                [void]$sb.Append("<text x='$xC' y='$($pT+$gH+16)' text-anchor='middle' font-family='Arial' font-size='8' fill='#778'>${hi}h</text>")
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
        [void]$sb.Append("<text x='$pL' y='$($pT-2)' font-family='Arial' font-size='8' fill='#778'>$Unit</text>")
        [void]$sb.Append("</svg>")
        return $sb.ToString()
    } catch { Write-Log "AgpSvg ERR: $($_.Exception.Message)"; return "" }
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
        $svgHist = Get-HistSvg -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn
        $svgAgp  = Get-AgpSvg  -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn

        $html = @"
<!DOCTYPE html>
<html lang="$(if($lEn){'en'}else{'pl'})">
<head>
<meta charset="UTF-8">
<style>
  @page { margin: 14mm 12mm 14mm 12mm; size: A4; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; color: #1a1a2e; background: #fff; }

  .header { display: flex; justify-content: space-between; align-items: flex-start;
            border-bottom: 2px solid #1a1a4a; padding-bottom: 10px; margin-bottom: 14px; }
  .header-left h1 { font-size: 20px; color: #1a1a4a; font-weight: 700; }
  .header-left .sub { font-size: 11px; color: #667; margin-top: 2px; }
  .header-right { font-size: 10px; color: #667; text-align: right; line-height: 1.6; }

  .patient-box { background: #f0f2ff; border-left: 4px solid #3344aa;
                 padding: 9px 14px; margin-bottom: 14px; border-radius: 0 6px 6px 0; }
  .patient-name { font-size: 16px; font-weight: 700; color: #1a1a4a; }
  .patient-meta { font-size: 11px; color: #556; margin-top: 3px; }

  .grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-bottom: 16px; }
  .card { border: 1px solid #dde2ff; border-radius: 6px; padding: 9px 6px; text-align: center; }
  .lbl { font-size: 9px; color: #889; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
  .val { font-size: 17px; font-weight: 700; line-height: 1; }
  .unit { font-size: 10px; font-weight: 400; color: #667; }
  .c-green  { color: #0a7a0a; }
  .c-blue   { color: #1144aa; }
  .c-orange { color: #bb5500; }
  .c-purple { color: #6622aa; }
  .c-teal   { color: #006677; }
  .c-gray   { color: #445566; }

  .section-title { font-size: 10px; font-weight: 700; color: #3344aa;
                   text-transform: uppercase; letter-spacing: 0.5px;
                   border-bottom: 1px solid #dde2ff; padding-bottom: 4px; margin-bottom: 8px; }

  .footer { margin-top: 14px; font-size: 9px; color: #aaa;
            text-align: center; border-top: 1px solid #eee; padding-top: 6px; }
  .range-legend { font-size: 10px; color: #667; margin-bottom: 12px; }
  .range-legend span { margin-right: 14px; }
  .dot-green  { display: inline-block; width: 8px; height: 8px; background: #0a7a0a; border-radius: 50%; margin-right: 3px; }
  .dot-red    { display: inline-block; width: 8px; height: 8px; background: #cc1111; border-radius: 50%; margin-right: 3px; }
  .dot-orange { display: inline-block; width: 8px; height: 8px; background: #bb5500; border-radius: 50%; margin-right: 3px; }
  .chart-box  { margin-bottom: 16px; border: 1px solid #dde2ff; border-radius: 6px; padding: 10px 8px 6px; background: #f8f9ff; }
  .chart-legend { font-size: 9px; color: #778; margin-top: 5px; }
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
                $gl.Stroke = New-Object System.Windows.Media.SolidColorBrush(
                    [System.Windows.Media.ColorConverter]::ConvertFromString("#33ffffff"))
                $gl.StrokeThickness = 0.5
                $cv.Children.Add($gl) | Out-Null
                $yl = New-Object System.Windows.Controls.TextBlock
                $yl.Text       = [Math]::Round($gv, $decPl).ToString($fmt2)
                $yl.FontSize   = 8
                $yl.Foreground = New-Object System.Windows.Media.SolidColorBrush(
                    [System.Windows.Media.ColorConverter]::ConvertFromString("#8888bb"))
                $yl.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
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
                $xLbl.Foreground = New-Object System.Windows.Media.SolidColorBrush(
                    [System.Windows.Media.ColorConverter]::ConvertFromString("#7777aa"))
                $xLbl.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
                [System.Windows.Controls.Canvas]::SetLeft($xLbl, $xCenter - 14)
                [System.Windows.Controls.Canvas]::SetTop($xLbl,  $padT + $gH + 5)
                $cv.Children.Add($xLbl) | Out-Null

                if (-not $script:AvgBarsHasData[$dataSlot]) { continue }
                $avg = $script:AvgBarsAvgs[$dataSlot]

                # Kolor słupka
                $barCol = if   ($avg -lt $script:AvgBarsLoN -or $avg -gt $script:AvgBarsHiC) { "#CC4444" } `
                          elseif ($avg -gt $script:AvgBarsHiN) { "#FFAA44" } `
                          else { "#6CBF26" }

                # Wysokość słupka od dołu osi
                $barH   = [Math]::Max(2.0, ($avg - $loY) / $rngY * $gH)
                $barTop = $padT + $gH - $barH

                $bar = New-Object System.Windows.Shapes.Rectangle
                $bar.Width   = $barW
                $bar.Height  = $barH
                $bar.RadiusX = 3; $bar.RadiusY = 3
                $bar.Fill    = New-Object System.Windows.Media.SolidColorBrush(
                    [System.Windows.Media.ColorConverter]::ConvertFromString($barCol))
                [System.Windows.Controls.Canvas]::SetLeft($bar, $xCenter - $barW / 2.0)
                [System.Windows.Controls.Canvas]::SetTop($bar,  $barTop)
                $cv.Children.Add($bar) | Out-Null

                # Wartość średnia nad słupkiem
                $valLbl = New-Object System.Windows.Controls.TextBlock
                $valLbl.Text       = [Math]::Round($avg, $decPl).ToString($fmt2)
                $valLbl.FontSize   = 9
                $valLbl.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI Semibold")
                $valLbl.Foreground = [System.Windows.Media.Brushes]::White
                [System.Windows.Controls.Canvas]::SetLeft($valLbl, $xCenter - 10)
                [System.Windows.Controls.Canvas]::SetTop($valLbl,  $barTop - 14)
                $cv.Children.Add($valLbl) | Out-Null
            }

          } catch { Write-Log "AvgBars Render ERR: $($_.Exception.Message)" }
        })
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
            $normR.Fill = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#1500CC44"))
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
                    $col  = if ($avg -lt $loN) { "#884444ff" } elseif ($avg -gt $hiN) { "#88ff8800" } else { "#8844cc44" }
                    $bar.Fill   = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($col))
                    $bar.Width  = $barW; $bar.Height = [Math]::Max(2, [Math]::Abs($yMin - $yMax))
                    [System.Windows.Controls.Canvas]::SetLeft($bar, $xCenter - $barW/2)
                    [System.Windows.Controls.Canvas]::SetTop($bar,  [Math]::Min($yMin,$yMax))
                    $cv.Children.Add($bar) | Out-Null
                    # Punkt sredniej
                    $dot = New-Object System.Windows.Shapes.Ellipse
                    $dot.Width = 6; $dot.Height = 6
                    $dot.Fill  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("White"))
                    [System.Windows.Controls.Canvas]::SetLeft($dot, $xCenter - 3)
                    [System.Windows.Controls.Canvas]::SetTop($dot,  $yAvg - 3)
                    $cv.Children.Add($dot) | Out-Null
                }
                # Etykieta godziny co 3h
                if ($h % 3 -eq 0) {
                    $lbl = New-Object System.Windows.Controls.TextBlock
                    $lbl.Text = "${h}h"; $lbl.FontSize = 8
                    $lbl.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#7777aa"))
                    $lbl.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
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
                $yl.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString("#7777aa"))
                $yl.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
                [System.Windows.Controls.Canvas]::SetLeft($yl, 2)
                [System.Windows.Controls.Canvas]::SetTop($yl,  $yp - 7)
                $cv.Children.Add($yl) | Out-Null
            }
          } catch { Write-Log "AGP Render ERR: $($_.Exception.Message)" }
        })
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
        $svgHist = Get-HistSvg -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn
        $svgAgp  = Get-AgpSvg  -Data $data -LoN $loN -HiN $hiN -UseMgDl $script:UseMgDl -Unit $unit -LangEn $lEn
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
            $arrow = switch ($trend) { 1{"$dn$dn"} 2{"$dn"} 4{"$up"} 5{"$up$up"} default{""} }
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
    $lblTrend.Text   = (t "TrendLbl")
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

$window.Show()
[System.Windows.Threading.Dispatcher]::Run()
Write-Log "=== STOP ==="