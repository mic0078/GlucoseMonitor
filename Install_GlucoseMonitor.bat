@echo off
chcp 65001 >nul

:: ============================================================
:: Glucose Monitor - Instalator
:: ============================================================

:: Wymuszenie uprawnien administratora
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Wymagane uprawnienia administratora. Ponowne uruchomienie...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

setlocal
set "INSTALL_DIR=C:\Glucose"
set "SOURCE_DIR=%~dp0"

echo ============================================================
echo  Glucose Monitor - Instalacja
echo ============================================================
echo.

:: Sprawdz czy pliki istnieja
if not exist "%SOURCE_DIR%GlucoseMonitor.ps1" (
    echo BLAD: Nie znaleziono GlucoseMonitor.ps1
    pause
    exit /b 1
)
if not exist "%SOURCE_DIR%Launch_GlucoseMonitor.ps1" (
    echo BLAD: Nie znaleziono Launch_GlucoseMonitor.ps1
    pause
    exit /b 1
)

:: Utworz folder
if not exist "%INSTALL_DIR%" (
    echo Tworzenie folderu %INSTALL_DIR%...
    mkdir "%INSTALL_DIR%"
)

:: Kopiuj pliki
echo Kopiowanie plikow...
copy /Y "%SOURCE_DIR%GlucoseMonitor.ps1"       "%INSTALL_DIR%\GlucoseMonitor.ps1"       >nul
copy /Y "%SOURCE_DIR%Launch_GlucoseMonitor.ps1" "%INSTALL_DIR%\Launch_GlucoseMonitor.ps1" >nul

:: ExecutionPolicy
echo Ustawianie ExecutionPolicy...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force" >nul 2>&1

:: Nadaj uprawnienia zapisu do folderu Glucose (aby config.ini sie zapisywal bez admina)
echo Ustawianie uprawnien folderu...
icacls "%INSTALL_DIR%" /grant Users:(OI)(CI)M /T >nul 2>&1

:: Generuj ikone kropla krwi
echo Generowanie ikony...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Drawing; $b=New-Object System.Drawing.Bitmap(32,32); $g=[System.Drawing.Graphics]::FromImage($b); $g.SmoothingMode=[System.Drawing.Drawing2D.SmoothingMode]::AntiAlias; $g.PixelOffsetMode=[System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality; $g.Clear([System.Drawing.Color]::Transparent); $p=New-Object System.Drawing.Drawing2D.GraphicsPath; $p.AddBezier(16,2,4,6,2,18,6,25); $p.AddBezier(6,25,7,32,25,32,26,25); $p.AddBezier(26,25,30,18,28,6,16,2); $p.CloseFigure(); $gb=New-Object System.Drawing.Drawing2D.LinearGradientBrush((New-Object System.Drawing.PointF 8,2),(New-Object System.Drawing.PointF 26,32),[System.Drawing.Color]::FromArgb(255,80,40),[System.Drawing.Color]::FromArgb(140,0,0)); $g.FillPath($gb,$p); $gb.Dispose(); $pn=New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(90,0,0),1.5); $g.DrawPath($pn,$p); $pn.Dispose(); $hb=New-Object System.Drawing.Drawing2D.LinearGradientBrush((New-Object System.Drawing.PointF 9,8),(New-Object System.Drawing.PointF 14,19),[System.Drawing.Color]::FromArgb(160,255,255,255),[System.Drawing.Color]::FromArgb(0,255,255,255)); $g.FillEllipse($hb,9,8,6,9); $hb.Dispose(); $p.Dispose(); $g.Dispose(); $ico=[System.Drawing.Icon]::FromHandle($b.GetHicon()); $fs=[System.IO.File]::OpenWrite('C:\Glucose\glucose.ico'); $ico.Save($fs); $fs.Close()"

:: Usun stary skrot (wymusza odczyt nowej ikony zamiast cache)
del /f /q "%PUBLIC%\Desktop\Glucose Monitor.lnk" >nul 2>&1
del /f /q "%USERPROFILE%\Desktop\Glucose Monitor.lnk" >nul 2>&1

:: Skrot na pulpicie
echo Tworzenie skrotu na pulpicie...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$d=[Environment]::GetFolderPath('CommonDesktopDirectory');" ^
  "$ws=New-Object -ComObject WScript.Shell;" ^
  "$sc=$ws.CreateShortcut(\"$d\Glucose Monitor.lnk\");" ^
  "$sc.TargetPath='powershell.exe';" ^
  "$sc.Arguments='-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File \"C:\Glucose\Launch_GlucoseMonitor.ps1\"';" ^
  "$sc.WorkingDirectory='C:\Glucose';" ^
  "$sc.IconLocation='C:\Glucose\glucose.ico,0';" ^
  "$sc.WindowStyle=7;" ^
  "$sc.Save();" ^
  "$bytes=[System.IO.File]::ReadAllBytes(\"$d\Glucose Monitor.lnk\");" ^
  "$bytes[0x15]=$bytes[0x15] -bor 0x20;" ^
  "[System.IO.File]::WriteAllBytes(\"$d\Glucose Monitor.lnk\",$bytes)"

:: Wymus odswiezenie cache ikon Windows
echo Odswiezanie cache ikon...
ie4uinit.exe -ClearIconCache >nul 2>&1
powershell -NoProfile -Command "Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class Shl{[DllImport(\"shell32.dll\")]public static extern void SHChangeNotify(int e,int f,IntPtr a,IntPtr b);}';[Shl]::SHChangeNotify(0x8000000,0,[IntPtr]::Zero,[IntPtr]::Zero)" >nul 2>&1

echo.
echo ============================================================
echo  Instalacja zakonczona!
echo  Folder:  %INSTALL_DIR%
echo  Skrot:   Pulpit - Glucose Monitor
echo ============================================================
echo.
pause
endlocal
