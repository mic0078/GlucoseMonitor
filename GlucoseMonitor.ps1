# ============================================================================
# GLUCOSE MONITOR - LibreLinkUp v5 (mmol/L, tray, movable)
# ============================================================================

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
    Email    = ""
    Password = ""
    ApiUrl   = "https://api-eu.libreview.io"
    Interval = 60
    Version  = "4.16.0"
    Product  = "llu.ios"
}

# Ustal folder skryptu (robust fallback)
$script:ScriptDir = $PSScriptRoot
if (-not $script:ScriptDir) {
    $def = $MyInvocation.MyCommand.Definition
    if ($def -and $def -notmatch "`n") { try { $script:ScriptDir = Split-Path -Parent $def } catch {} }
}
if (-not $script:ScriptDir) { $script:ScriptDir = $PWD.Path }
if (-not $script:ScriptDir) { $script:ScriptDir = "C:\Glucose" }

$script:ConfigFile = Join-Path $script:ScriptDir "config.ini"

$script:LogFile = Join-Path $script:ScriptDir "glucose_debug.log"
function Write-Log { param([string]$M); $l="[$(Get-Date -Format 'HH:mm:ss')] $M"; Write-Host $l; try{Add-Content -Path $script:LogFile -Value $l -ErrorAction Stop}catch{} }

$script:AuthToken=$null; $script:AccountId=$null; $script:AccountIdHash=$null; $script:PatientId=$null; $script:Timer=$null
$script:UseMgDl = $false
$script:CachedMgDl = $null; $script:CachedTrend = 0; $script:CachedGraphData = $null
$script:LangEn = $false   # false = PL (domyslny), true = EN
$script:LastHistorySave = $null
$script:BackoffUntil = $null  # 429 backoff - nie wywoluj API do tej daty

$script:T = @{
    Fetching   = @("Pobieranie...",                       "Fetching...")
    Connected  = @("Polaczono | Odczyt aktualny",         "Connected | Reading current")
    TooMany    = @("Zbyt wiele zadan - sprobuj pozniej",  "Too many requests - try again later")
    NoData     = @("Brak danych",                         "No data")
    TrayGlc    = @("Glukoza: ",                           "Glucose: ")
    TrayTooM   = @("Glucose Monitor - zbyt wiele zadan",  "Glucose Monitor - too many requests")
    TrayNoDat  = @("Glucose Monitor - brak danych",       "Glucose Monitor - no data")
    ShowWin    = @("Pokaz okno",                          "Show window")
    SwitchAcc  = @("Zmien konto",                         "Switch account")
    CloseApp   = @("Zamknij",                             "Exit")
    TrendLbl   = @("Trend",                               "Trend")
    AvgLbl     = @("Sred.",                               "Avg")
    UnitTip    = @("Przelacz jednostki",                  "Switch units")
    RefreshBtn = @("Odswiez",                             "Refresh")
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

function Save-HistoryEntry([double]$mgdl, [int]$trend) {
    try {
        $now = Get-Date
        if ($script:LastHistorySave -and ($now - $script:LastHistorySave).TotalMinutes -lt 2) { return }
        $script:LastHistorySave = $now
        $entry = '{"ts":"' + $now.ToString("yyyy-MM-ddTHH:mm:ss") + '","mgdl":' + [Math]::Round($mgdl,1) + ',"trend":' + $trend + '}'
        Add-Content -Path $script:HistoryFile -Value $entry -Encoding UTF8 -ErrorAction Stop
    } catch {}
}

function Load-HistoryData([int]$days) {
    $result = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path $script:HistoryFile)) { return $result }
    $cutoff = (Get-Date).AddDays(-$days)
    try {
        Get-Content $script:HistoryFile -Encoding UTF8 | ForEach-Object {
            try {
                $obj = $_ | ConvertFrom-Json
                $ts  = [DateTime]::Parse($obj.ts)
                if ($ts -ge $cutoff) { $result.Add($obj) }
            } catch {}
        }
    } catch {}
    return $result
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
            if ($r.data.connection -and $r.data.connection.glucoseMeasurement) {
                $gm=$r.data.connection.glucoseMeasurement
                $result.CurrentGlucose = if($gm.ValueInMgPerDl){$gm.ValueInMgPerDl}else{$gm.Value}
                $result.Trend=$gm.TrendArrow; $result.Timestamp=$gm.Timestamp
            }
            if ($r.data.graphData) {
                $result.GraphData=@($r.data.graphData)
                # Dolacz biezacy odczyt jako ostatni punkt wykresu (API zwraca go oddzielnie)
                if ($r.data.connection -and $r.data.connection.glucoseMeasurement) {
                    $gmTs = $r.data.connection.glucoseMeasurement.Timestamp
                    $lastTs = if ($result.GraphData.Count -gt 0) { $result.GraphData[-1].Timestamp } else { $null }
                    if ($gmTs -ne $lastTs) { $result.GraphData += $r.data.connection.glucoseMeasurement }
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
    "Email=$($script:Config.Email)`nEncryptedPassword=$encPass" | Set-Content -Path $script:ConfigFile -Encoding UTF8
}

function Load-Config {
    if (Test-Path $script:ConfigFile) {
        $lines = Get-Content $script:ConfigFile -Encoding UTF8
        $encPass = $null
        foreach ($line in $lines) {
            if ($line -match '^Email=(.+)$')             { $script:Config.Email    = $Matches[1] }
            if ($line -match '^Password=(.+)$')          { $script:Config.Password = $Matches[1] }  # stary format - plain text
            if ($line -match '^EncryptedPassword=(.+)$') { $encPass = $Matches[1] }
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
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <TextBlock Name="lblTrend" Text="Trend" Style="{StaticResource L}"/>
                        <TextBlock Name="txtTrendText" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                        <TextBlock Text="Min" Style="{StaticResource L}"/>
                        <TextBlock Name="txtMin" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                        <TextBlock Name="lblSred" Text="Sred." Style="{StaticResource L}"/>
                        <TextBlock Name="txtAvg" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="3" HorizontalAlignment="Right">
                        <TextBlock Text="Max" Style="{StaticResource L}"/>
                        <TextBlock Name="txtMax" Text="---" Style="{StaticResource V}" FontSize="11"/>
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
                <Button Grid.Column="7" Name="btnRefresh" Content="Odswiez" Background="#3a3a5a"
                        Foreground="White" BorderThickness="0" Padding="10,4" FontSize="10" Cursor="Hand"/>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$txtPatient=$window.FindName("txtPatient"); $txtStatus=$window.FindName("txtStatus")
$txtGlucoseValue=$window.FindName("txtGlucoseValue"); $txtTrendArrow=$window.FindName("txtTrendArrow")
$txtGlucoseStatus=$window.FindName("txtGlucoseStatus"); $txtTrendText=$window.FindName("txtTrendText")
$txtDelta=$window.FindName("txtDelta")
$txtTimestamp=$window.FindName("txtTimestamp"); $txtMin=$window.FindName("txtMin")
$txtAvg=$window.FindName("txtAvg"); $txtMax=$window.FindName("txtMax")
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
    $window.WindowState = [System.Windows.WindowState]::Minimized
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
        $script:ForceTimer.Stop()
        $window.Opacity = 1.0
    }
})

# ======================== WYKRES ========================
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

    # Linia wykresu - kolorowe segmenty
    $hiH = if ($script:UseMgDl) { 180.0 } else { 10.0 }
    $hiC = if ($script:UseMgDl) { 250.0 } else { 13.9 }
    $loC = if ($script:UseMgDl) {  70.0 } else {  3.9 }
    $step=$dw/[Math]::Max(1,$vals.Count-1)
    for($i=0;$i -lt $vals.Count-1;$i++) {
        $x0=$m+($i*$step);   $y0=$m+$dh-(($dh/$rng)*($vals[$i]-$mn))
        $x1=$m+(($i+1)*$step); $y1=$m+$dh-(($dh/$rng)*($vals[$i+1]-$mn))
        $avg2=($vals[$i]+$vals[$i+1])/2.0
        $segCol = if ($avg2 -lt $loC -or $avg2 -gt $hiC) { "#EE4444" } elseif ($avg2 -gt $hiH) { "#FFAA44" } else { "#44DDAA" }
        $seg=New-Object System.Windows.Shapes.Line
        $seg.X1=$x0;$seg.Y1=$y0;$seg.X2=$x1;$seg.Y2=$y1
        $seg.Stroke=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($segCol))
        $seg.StrokeThickness=2; $seg.StrokeStartLineCap="Round"; $seg.StrokeEndLineCap="Round"
        $canvasGraph.Children.Add($seg)|Out-Null
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

# ======================== RENDER UI (z cache) ========================
function Render-GlucoseUI {
    if ($null -eq $script:CachedMgDl) { return }
    $mgdl  = $script:CachedMgDl
    $mmol  = MgToMmol $mgdl
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

    if ($script:IsCompact -and $script:TxtCompactGlucose) {
        $script:TxtCompactGlucose.Text       = "$displayVal $(Get-TrendArrow $t)"
        $script:TxtCompactGlucose.Foreground = $brush
    }
}

# ======================== UPDATE ========================
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
        if($data.GraphData -and $data.GraphData.Count -gt 0) { $script:CachedGraphData = $data.GraphData }

        # Czulszy algorytm trendu - oblicz z danych wykresu
        $calcTrend = Get-CalculatedTrend $script:CachedGraphData
        if ($null -ne $calcTrend) { $script:CachedTrend = $calcTrend }

        # Zapisz do historii
        Save-HistoryEntry $script:CachedMgDl $script:CachedTrend

        if($data.PatientName){$txtPatient.Text=$data.PatientName}
        if($data.Timestamp) {
            $ts=[string]$data.Timestamp; $p=[DateTime]::MinValue
            $fmts=@("M/d/yyyy h:mm:ss tt","d/M/yyyy h:mm:ss tt","M/d/yyyy H:mm:ss","d/M/yyyy H:mm:ss","yyyy-MM-ddTHH:mm:ss")
            foreach($f in $fmts){if([DateTime]::TryParseExact($ts,$f,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$p)){$txtTimestamp.Text=$p.ToString("HH:mm");break}}
        }

        Render-GlucoseUI

        $txtStatus.Text=(t "Connected")
        $txtStatus.Foreground=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(100,180,100))
    } else {
        $txtStatus.Text=(t "NoData"); $txtStatus.Foreground=[System.Windows.Media.Brushes]::OrangeRed; $txtGlucoseValue.Text="---"
        Update-TrayTooltip (t "TrayNoDat")
    }
    $txtNextUpdate.Text = "Refresh: $($script:SecondsLeft)s"
}

# ======================== HISTORIA ========================
$script:HistWin        = $null
$script:HistCanvas     = $null
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

        $data = Load-HistoryData $days
        $script:HistCanvas.Children.Clear()

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
            $script:HistValAvg.Text   = [Math]::Round($sm.Average,1).ToString($fmt)
            $script:HistValMin.Text   = $sm.Minimum.ToString($fmt)
            $script:HistValMax.Text   = $sm.Maximum.ToString($fmt)
            $script:HistValTIR.Text   = "$tirPct%"
            if ($deltas.Count -gt 0) {
                $avgDelta = ($deltas | Measure-Object -Average).Average
                $dUnit    = if ($script:UseMgDl) { "" } else { "" }
                $avgDeltaDisp = if ($script:UseMgDl) { [Math]::Round($avgDelta,0) } else { [Math]::Round($avgDelta/18.018,1) }
                $script:HistValDelta.Text = $avgDeltaDisp.ToString($fmt) + $dUnit
            } else { $script:HistValDelta.Text = "---" }
        }

        # Wymiary canvas
        if ($script:HistWin) { $script:HistWin.UpdateLayout() }
        $cW = $script:HistCanvas.ActualWidth;  if ($cW -lt 10) { $cW = 440.0 }
        $cH = $script:HistCanvas.ActualHeight; if ($cH -lt 10) { $cH = 250.0 }
        $padL=38.0; $padR=8.0; $padT=8.0; $padB=22.0
        $gW = $cW - $padL - $padR
        $gH = $cH - $padT  - $padB
        $n  = $vals.Count

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
        $zones = @(
            @{ yTop=$hiHyper;  yBot=$hiY;       col="#44FF3333" }
            @{ yTop=$hiYellow; yBot=$hiHyper;   col="#33FFAA00" }
            @{ yTop=$loNormY;  yBot=$hiYellow;  col="#2200CC44" }
            @{ yTop=$loY;      yBot=$loNormY;   col="#44FF3333" }
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

        # --- Linia danych (segmenty kolorowe, os X proporcjonalna do czasu) ---
        for ($i = 0; $i -lt ($n - 1); $i++) {
            $v0 = $vals[$i]; $v1 = $vals[$i+1]
            # Pomijaj segmenty z przerwa > 30 min (brak sensora / przerwa w danych)
            if (($tss[$i+1] - $tss[$i]).TotalMinutes -gt 30) { continue }
            $x0 = $padL + ($tss[$i]   - $firstTs).TotalSeconds / $totalSecs * $gW
            $x1 = $padL + ($tss[$i+1] - $firstTs).TotalSeconds / $totalSecs * $gW
            $y0 = $padT + $gH - ($v0 - $loY) / $rangeY * $gH
            $y1 = $padT + $gH - ($v1 - $loY) / $rangeY * $gH
            $avg2 = ($v0 + $v1) / 2.0
            $hiH2 = if ($script:UseMgDl) { 180.0 } else { 10.0 }
            $hiC2 = if ($script:UseMgDl) { 250.0 } else { 13.9 }
            $lC2  = if ($script:UseMgDl) {  70.0 } else {  3.9 }
            $segCol = if ($avg2 -lt $lC2 -or $avg2 -gt $hiC2) { "#EE4444" } elseif ($avg2 -gt $hiH2) { "#FFAA44" } else { "#44DDAA" }
            $seg = New-Object System.Windows.Shapes.Line
            $seg.X1=$x0; $seg.Y1=$y0; $seg.X2=$x1; $seg.Y2=$y1
            $seg.Stroke = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($segCol))
            $seg.StrokeThickness = 1.5
            $script:HistCanvas.Children.Add($seg) | Out-Null
        }
    } catch { Write-Log "Render-HistGraph err: $($_.Exception.Message)" }
}

function Show-HistoryWindow {
    # Jesli okno juz jest otwarte - przywroc je
    if ($script:HistWin -and $script:HistWin.IsLoaded) {
        $script:HistWin.Activate(); return
    }

    try {
    [xml]$hXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Glucose History" Width="500" Height="470"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="Transparent" WindowStyle="None" AllowsTransparency="True">
    <Border Background="#1a1a2e" CornerRadius="10">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="30"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Background="#12122a" CornerRadius="10,10,0,0" Name="hTitleBar">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Name="hTitleTxt" Text="  Historia glukozy"
                               Foreground="#7777aa" FontSize="11" VerticalAlignment="Center"
                               FontFamily="Segoe UI" Margin="6,0,0,0"/>
                    <Button Grid.Column="1" Name="hClose" Content="&#x2715;" Width="34" Height="30"
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

            <Border Grid.Row="2" Background="#222244" CornerRadius="6" Margin="12,0,12,6" Padding="10,7">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" HorizontalAlignment="Center">
                        <TextBlock Name="hLblAvg"   Text="Srednia"  Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValAvg"   Text="---"      Foreground="White"   FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                        <TextBlock                  Text="Min"      Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValMin"   Text="---"      Foreground="#44aaff" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                        <TextBlock                  Text="Max"      Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValMax"   Text="---"      Foreground="#ffaa44" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Column="3" HorizontalAlignment="Center">
                        <TextBlock Name="hLblTIR"   Text="W normie" Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValTIR"   Text="---"      Foreground="#44DD44" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                    <StackPanel Grid.Column="4" HorizontalAlignment="Center">
                        <TextBlock Name="hLblDelta" Text="Delta avg" Foreground="#7777aa" FontSize="9"  HorizontalAlignment="Center" FontFamily="Segoe UI"/>
                        <TextBlock Name="hValDelta" Text="---"       Foreground="#aaaacc" FontSize="12" HorizontalAlignment="Center" FontFamily="Segoe UI Semibold"/>
                    </StackPanel>
                </Grid>
            </Border>

            <Grid Grid.Row="3" Margin="12,0,12,12">
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

    $script:HistDays = 7

    # Renderuj przy otwarciu (Add_ContentRendered wykonuje sie po Show())
    $script:HistWin.Add_ContentRendered({ Render-HistGraph $script:HistDays })

    # Handlery przyciskow okresu - uzywaja $script: wiec dzialaja poprawnie
    $script:HistWin.FindName("hClose").Add_Click({ $script:HistWin.Close() })
    $script:HistBtns[0].Add_Click({ $script:HistDays = 7;  Render-HistGraph 7  })
    $script:HistBtns[1].Add_Click({ $script:HistDays = 14; Render-HistGraph 14 })
    $script:HistBtns[2].Add_Click({ $script:HistDays = 30; Render-HistGraph 30 })
    $script:HistBtns[3].Add_Click({ $script:HistDays = 90; Render-HistGraph 90 })

    $script:HistWin.Show()

    } catch { Write-Log "Show-HistoryWindow error: $($_.Exception.Message)" }
}

# ======================== TRAY ========================
$script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
$script:NotifyIcon.Text = "Glucose Monitor"; $script:NotifyIcon.Visible = $true

function New-TrayIcon {
    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode    = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode  = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)

    # Ksztalt kropelki: 3 segmenty Beziera
    #   Czubek (16,2) -> lewa strona -> dolna polokrag -> prawa strona -> czubek
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddBezier(16,  2,   4,  6,   2, 18,   6, 25)   # czubek -> lewy bok
    $path.AddBezier( 6, 25,   7, 32,  25, 32,  26, 25)   # lewy dolny -> prawy dolny (okragle dno)
    $path.AddBezier(26, 25,  30, 18,  28,  6,  16,  2)   # prawy bok -> czubek
    $path.CloseFigure()

    # Gradient: jasny czerwono-pomaranczowy (gora) -> ciemny czerwony (dol)
    $gradBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.PointF  8,  2),
        (New-Object System.Drawing.PointF 26, 32),
        [System.Drawing.Color]::FromArgb(255, 80, 40),
        [System.Drawing.Color]::FromArgb(140,  0,  0)
    )
    $g.FillPath($gradBrush, $path)
    $gradBrush.Dispose()

    # Ciemny obrys
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(90, 0, 0), 1.5)
    $g.DrawPath($pen, $path)
    $pen.Dispose()

    # Blysk: bialy gradient w gornej-lewej czesci dla efektu 3D
    $hlBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.PointF  9,  8),
        (New-Object System.Drawing.PointF 14, 19),
        [System.Drawing.Color]::FromArgb(160, 255, 255, 255),
        [System.Drawing.Color]::FromArgb(  0, 255, 255, 255)
    )
    $g.FillEllipse($hlBrush, 9, 8, 6, 9)
    $hlBrush.Dispose()

    $path.Dispose()
    $g.Dispose()
    return [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}
$script:NotifyIcon.Icon = New-TrayIcon

$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuShow = $trayMenu.Items.Add((t "ShowWin"))
$menuShow.Add_Click({
    $window.Show()
    $window.WindowState = [System.Windows.WindowState]::Normal
    $window.Topmost = $true
    $window.Activate()
})
$menuLogout = $trayMenu.Items.Add((t "SwitchAcc"))
$menuLogout.Add_Click({
    $script:AuthToken = $null; $script:PatientId = $null
    if (Test-Path $script:ConfigFile) { Remove-Item $script:ConfigFile -Force }
    $result = Show-LoginWindow
    if ($result) { Update-Display; $script:SecondsLeft = $script:Config.Interval }
})
$menuExit = $trayMenu.Items.Add((t "CloseApp"))
$menuExit.Add_Click({ $window.Close() })
$script:NotifyIcon.ContextMenuStrip = $trayMenu
$script:NotifyIcon.Add_DoubleClick({
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
        $txtNextUpdate.Text = "Refresh: $($script:SecondsLeft)s"
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
    $menuShow.Text   = (t "ShowWin")
    $menuLogout.Text = (t "SwitchAcc")
    $menuExit.Text   = (t "CloseApp")
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
})

$window.Add_ContentRendered({
    Update-Display
    $txtNextUpdate.Text = "Refresh: $($script:SecondsLeft)s"
    $script:Timer.Start()
})

$window.Add_Closed({
    if($script:Timer){$script:Timer.Stop()}
    if($script:ForceTimer){$script:ForceTimer.Stop()}
    $script:NotifyIcon.Visible=$false; $script:NotifyIcon.Dispose()
    [Native.Win32]::ShowWindow($consoleHwnd, 5)|Out-Null
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

$window.ShowDialog() | Out-Null; Write-Log "=== STOP ==="