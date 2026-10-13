@echo off
setlocal EnableExtensions
chcp 65001 >nul
title SKF Fatura Veri Portali (Yerel Kullanim)

set "APP_ROOT=%~dp0"
set "LOG_FILE=%APP_ROOT%baslatma.log"
set "VENV_PY=%APP_ROOT%.venv\Scripts\python.exe"
cd /d "%APP_ROOT%"

>"%LOG_FILE%" echo Baslatma gunu: %DATE% %TIME%
>>"%LOG_FILE%" echo Proje klasoru: %APP_ROOT%

echo SKF Fatura Veri Portali baslatiliyor...
echo Bu uygulama bilgisayarinizda calisir; fatura dosyalari buluta yuklenmez.
echo.

if not exist "%APP_ROOT%run_app.py" goto project_missing
if not exist "%APP_ROOT%06-Haziran\requirements.txt" goto project_missing

if exist "%VENV_PY%" goto check_venv
goto create_venv

:check_venv
"%VENV_PY%" --version >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto venv_invalid
goto install_packages

:create_venv
echo Bilgisayarinizdaki Python surumu kontrol ediliyor...
where py >nul 2>&1
if not errorlevel 1 (
    for %%V in (3.14 3.13 3.12 3.11) do call :try_python %%V
)

if exist "%VENV_PY%" goto check_venv

where python >nul 2>&1
if errorlevel 1 goto python_missing
echo Varsayilan Python ile yerel ortam hazirlaniyor...
python -m venv "%APP_ROOT%.venv" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto python_missing
if not exist "%VENV_PY%" goto python_missing
goto install_packages

:try_python
if exist "%VENV_PY%" exit /b 0
echo Python %1 deneniyor...
>>"%LOG_FILE%" echo Python %1 deneniyor.
py -%1 -m venv "%APP_ROOT%.venv" >>"%LOG_FILE%" 2>&1
if errorlevel 1 exit /b 1
if exist "%VENV_PY%" exit /b 0
exit /b 1

:install_packages
echo Gerekli paketler kontrol ediliyor. Ilk calistirmada bu islem birkac dakika surebilir...
"%VENV_PY%" -m pip install --disable-pip-version-check -r "%APP_ROOT%06-Haziran\requirements.txt" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto install_failed

echo.
echo Portal baslatiliyor. Lutfen bu siyah pencereyi kapatmayin.
echo Tarayici acilmazsa ekranda yazan localhost adresini tarayiciya kopyalayin.
echo Ortam ve paket kurulum kaydi: "%LOG_FILE%"
echo.
"%VENV_PY%" "%APP_ROOT%run_app.py"
set "APP_EXIT_CODE=%ERRORLEVEL%"
echo.
if "%APP_EXIT_CODE%"=="0" (
    echo Portal kapatildi.
) else (
    echo Portal baslatilamadi veya kapandi. Cikis kodu: %APP_EXIT_CODE%
    echo Ayrintilar baslatma.log dosyasina kaydedildi.
)
pause
exit /b %APP_EXIT_CODE%

:project_missing
echo Proje dosyalari eksik. ZIP dosyasini tamamen bir klasore cikartip BAT dosyasini o klasorden calistirin.
>>"%LOG_FILE%" echo HATA: run_app.py veya requirements.txt bulunamadi.
goto failed

:python_missing
echo Uyumlu Python surumu bulunamadi veya yerel Python ortami hazirlanamadi.
echo Python 3.11 veya daha yenisini kurup yeniden deneyin: https://www.python.org/downloads/
echo Hata ayrintilari baslatma.log dosyasina yazildi.
>>"%LOG_FILE%" echo HATA: Python sanal ortami hazirlanamadi.
goto failed

:venv_invalid
echo Mevcut .venv Python ortami acilamadi.
echo Ortam/paket kaydi baslatma.log dosyasina yazildi. .venv klasorunu silip tekrar deneyin.
>>"%LOG_FILE%" echo HATA: .venv icindeki Python calistirilamadi.
goto failed

:install_failed
echo Python paketleri kurulamadi. Internet baglantinizi kontrol edip yeniden deneyin.
echo Hata ayrintilari baslatma.log dosyasina yazildi.
goto failed

:failed
echo.
pause
exit /b 1
