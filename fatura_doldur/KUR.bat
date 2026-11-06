@echo off
chcp 65001 >nul
title Fatura Doldur - Kurulum
cd /d "%~dp0"
echo ============================================
echo   FATURA DOLDUR KURULUMU - lutfen bekleyin
echo ============================================
echo.

set "PYEXE="
call :pybul
if defined PYEXE goto pyvar

echo [1/4] Python kuruluyor...
winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
call :pybul
if defined PYEXE goto pyvar

echo winget olmadi, Python internetten indiriliyor...
powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile \"$env:TEMP\python_kur.exe\""
"%TEMP%\python_kur.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_tcltk=1 Include_launcher=1
call :pybul
if not defined PYEXE (
  echo.
  echo HATA: Python kurulamadi. Bu pencerenin fotografini cekip gonderin.
  pause
  exit /b 1
)

:pyvar
echo Python bulundu: %PYEXE%
echo.
echo [2/4] Gerekli paketler kuruluyor...
"%PYEXE%" -m pip install --upgrade pip -q
"%PYEXE%" -m pip install -r requirements.txt -q
if errorlevel 1 (
  echo HATA: Paketler kurulamadi. Internet baglantisini kontrol edin.
  pause
  exit /b 1
)

echo [3/4] SumatraPDF (yazdirma icin) kuruluyor...
winget install -e --id SumatraPDF.SumatraPDF --scope user --silent --accept-package-agreements --accept-source-agreements >nul 2>&1

echo [4/4] Masaustu kisayolu olusturuluyor...
if not exist "C:\Faturalar" mkdir "C:\Faturalar"
for %%d in ("%PYEXE%") do set "PYW=%%~dpdpythonw.exe"
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Fatura Doldur.lnk'); $s.TargetPath='%PYW%'; $s.Arguments='\"%~dp0fatura_doldur.py\"'; $s.WorkingDirectory='%~dp0'; $s.Save()"

echo.
echo ============================================
echo   KURULUM TAMAM!
echo   Masaustundeki "Fatura Doldur" simgesine tiklayin.
echo ============================================
pause
start "" "%PYW%" "%~dp0fatura_doldur.py"
exit /b 0

:pybul
set "PYEXE="
for /f "delims=" %%i in ('py -3 -c "import sys,tkinter;print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
if defined PYEXE exit /b 0
for /f "delims=" %%i in ('python -c "import sys,tkinter;print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
if defined PYEXE exit /b 0
for %%v in (313 312 311) do (
  if exist "%LOCALAPPDATA%\Programs\Python\Python%%v\python.exe" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python%%v\python.exe"
)
exit /b 0
