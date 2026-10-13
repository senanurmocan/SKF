@echo off
setlocal
chcp 65001 >nul
title SKF Fatura Veri Portalı (Yerel Kullanım)
set "APP_ROOT=%~dp0"
cd /d "%APP_ROOT%"

set "VENV_PY=%APP_ROOT%.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    where py >nul 2>nul
    if errorlevel 1 goto missing_python
    echo Yerel Python ortamı hazırlanıyor...
    py -3.14 -m venv "%APP_ROOT%.venv"
    if errorlevel 1 goto missing_python
)

echo Gerekli uygulama paketleri denetleniyor...
"%VENV_PY%" -m pip install --disable-pip-version-check -r "%APP_ROOT%06-Haziran\requirements.txt"
if errorlevel 1 goto install_failed

echo Portal bilgisayarınızda başlatılıyor. Dosyalarınız bu bilgisayarda kalır.
"%VENV_PY%" "%APP_ROOT%run_app.py"
set "APP_EXIT_CODE=%ERRORLEVEL%"
if not "%APP_EXIT_CODE%"=="0" (
    echo.
    echo Portal kapandı. Hata kodu: %APP_EXIT_CODE%
    pause
)
exit /b %APP_EXIT_CODE%

:missing_python
echo Python 3.14 bulunamadı. Python 3.14 yükleyip bu dosyayı yeniden çalıştırın.
echo https://www.python.org/downloads/
pause
exit /b 1

:install_failed
echo Uygulama paketleri kurulamadı. İnternet bağlantısını kontrol edip yeniden deneyin.
pause
exit /b 1
