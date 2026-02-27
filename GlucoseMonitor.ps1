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
    Product  = "llu.android"
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
            if ($r.data.graphData) { $result.GraphData=@($r.data.graphData) }
            if ($r.data.connection) { $result.PatientName="$($r.data.connection.firstName) $($r.data.connection.lastName)".Trim() }
            return $result
        }
        $script:AuthToken=$null;$script:PatientId=$null; return $null
    } catch {
        Write-Log "Graph ERR: $($_.Exception.Message)"
        if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 429) {
            $script:LastApiError = "429"
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
function Get-TrendText([int]$v){switch($v){1{"Szybki spadek"}2{"Spadek"}3{"Stabilny"}4{"Wzrost"}5{"Szybki wzrost"}default{"---"}}}
function Get-GlucoseColor([double]$mmol){
    if($mmol -lt 3.0){"#FF0000"}elseif($mmol -lt 3.9){"#FF6600"}elseif($mmol -le 10.0){"#00CC00"}elseif($mmol -le 13.9){"#FFAA00"}else{"#FF0000"}
}
function Get-GlucoseStatus([double]$mmol){
    if($mmol -lt 3.0){"BARDZO NISKI!"}elseif($mmol -lt 3.9){"Niski"}elseif($mmol -le 10.0){"W normie"}elseif($mmol -le 13.9){"Wysoki"}else{"BARDZO WYSOKI!"}
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
                        <TextBlock Text="Trend" Style="{StaticResource L}"/>
                        <TextBlock Name="txtTrendText" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                        <TextBlock Text="Min" Style="{StaticResource L}"/>
                        <TextBlock Name="txtMin" Text="---" Style="{StaticResource V}" FontSize="11"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                        <TextBlock Text="Sred." Style="{StaticResource L}"/>
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
                    <ColumnDefinition Width="5"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Name="txtNextUpdate" Text="" Foreground="White" FontSize="11" VerticalAlignment="Center"/>
                <Button Grid.Column="1" Name="btnUnit" Content="mmol/L" Background="#1e2a3a"
                        Foreground="#6677aa" BorderThickness="0" Padding="6,4" FontSize="9" Cursor="Hand"
                        ToolTip="Przelacz jednostki"/>
                <Button Grid.Column="3" Name="btnRefresh" Content="Odswiez" Background="#3a3a5a"
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
    if (-not $script:IsCompact) { return }
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

    $vals=@()
    foreach($i in $GraphData) {
        $mg = if($i.ValueInMgPerDl){[double]$i.ValueInMgPerDl}elseif($i.Value){[double]$i.Value}else{0}
        if($mg -gt 0) {
            if ($script:UseMgDl) { $vals += [Math]::Round($mg, 0) }
            else { $vals += MgToMmol $mg }
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

    # Linia wykresu
    $pl=New-Object System.Windows.Shapes.Polyline; $pl.StrokeThickness=2
    $pl.Stroke=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(50,220,50))
    $pl.StrokeLineJoin="Round"
    $pts=New-Object System.Windows.Media.PointCollection
    $step=$dw/[Math]::Max(1,$vals.Count-1)
    for($i=0;$i -lt $vals.Count;$i++) {
        $x=$m+($i*$step); $y=$m+$dh-(($dh/$rng)*($vals[$i]-$mn))
        $pts.Add((New-Object System.Windows.Point $x,$y))|Out-Null
    }
    $pl.Points=$pts; $canvasGraph.Children.Add($pl)|Out-Null

    # Ostatni punkt
    $lc=$vals.Count-1
    $lastX=$m+($lc*$step); $lastY=$m+$dh-(($dh/$rng)*($vals[$lc]-$mn))
    $dot=New-Object System.Windows.Shapes.Ellipse; $dot.Width=8;$dot.Height=8
    $dot.Fill=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(50,220,50))
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
    }

    if ($script:UseMgDl) {
        Update-TrayTooltip "Glukoza: $([Math]::Round($mgdl,0)) mg/dL $(Get-TrendArrow $t)"
    } else {
        Update-TrayTooltip "Glukoza: $($mmol.ToString('0.0')) mmol/L $(Get-TrendArrow $t)"
    }

    if ($script:IsCompact -and $script:TxtCompactGlucose) {
        $script:TxtCompactGlucose.Text       = "$displayVal $(Get-TrendArrow $t)"
        $script:TxtCompactGlucose.Foreground = $brush
    }
}

# ======================== UPDATE ========================
function Update-Display {
    $txtStatus.Text="Pobieranie..."; $txtStatus.Foreground=[System.Windows.Media.Brushes]::Yellow
    $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render,[Action]{})

    $data = Get-GlucoseData
    if ($script:LastApiError -eq "429") {
        $txtStatus.Text = "Zbyt wiele zadan - sprobuj pozniej"
        $txtStatus.Foreground = [System.Windows.Media.Brushes]::OrangeRed
        $txtGlucoseValue.Text = "---"
        Update-TrayTooltip "Glucose Monitor - zbyt wiele zadan"
        $script:LastApiError = $null
    } elseif ($data -and $data.CurrentGlucose) {
        # Zapisz do cache
        $script:CachedMgDl  = [double]$data.CurrentGlucose
        $script:CachedTrend = if($data.Trend){[int]$data.Trend}else{0}
        if($data.GraphData -and $data.GraphData.Count -gt 0) { $script:CachedGraphData = $data.GraphData }

        if($data.PatientName){$txtPatient.Text=$data.PatientName}
        if($data.Timestamp) {
            $ts=[string]$data.Timestamp; $p=[DateTime]::MinValue
            $fmts=@("M/d/yyyy h:mm:ss tt","d/M/yyyy h:mm:ss tt","M/d/yyyy H:mm:ss","d/M/yyyy H:mm:ss","yyyy-MM-ddTHH:mm:ss")
            foreach($f in $fmts){if([DateTime]::TryParseExact($ts,$f,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$p)){$txtTimestamp.Text=$p.ToString("HH:mm");break}}
        }

        Render-GlucoseUI

        $txtStatus.Text="Polaczono | Odczyt aktualny"
        $txtStatus.Foreground=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(100,180,100))
    } else {
        $txtStatus.Text="Brak danych"; $txtStatus.Foreground=[System.Windows.Media.Brushes]::OrangeRed; $txtGlucoseValue.Text="---"
        Update-TrayTooltip "Glucose Monitor - brak danych"
    }
    $txtNextUpdate.Text = "Refresh: $($script:SecondsLeft)s"
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
$menuShow = $trayMenu.Items.Add("Pokaz okno")
$menuShow.Add_Click({
    $window.Show()
    $window.WindowState = [System.Windows.WindowState]::Normal
    $window.Topmost = $true
    $window.Activate()
})
$menuLogout = $trayMenu.Items.Add("Zmien konto")
$menuLogout.Add_Click({
    $script:AuthToken = $null; $script:PatientId = $null
    if (Test-Path $script:ConfigFile) { Remove-Item $script:ConfigFile -Force }
    $result = Show-LoginWindow
    if ($result) { Update-Display; $script:SecondsLeft = $script:Config.Interval }
})
$menuExit = $trayMenu.Items.Add("Zamknij")
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