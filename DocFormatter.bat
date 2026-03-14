@echo off
chcp 65001 >nul 2>&1
title DocFormatter
cd /d "%~dp0"

echo.
echo   ══════════════════════════════════════
echo          DocFormatter starten
echo   ══════════════════════════════════════
echo.

:: ---------- Python prüfen ----------
where python >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=python
    goto :found
)
where python3 >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=python3
    goto :found
)

echo   X Python wurde nicht gefunden.
echo.
echo   ┌──────────────────────────────────────────────────┐
echo   │  So installierst du Python (dauert 2 Minuten):   │
echo   │                                                   │
echo   │  1. Oeffne deinen Browser und geh auf:           │
echo   │     python.org/downloads                          │
echo   │                                                   │
echo   │  2. Klick den gelben Button                      │
echo   │     "Download Python 3.x.x"                      │
echo   │                                                   │
echo   │  3. Starte die heruntergeladene .exe Datei       │
echo   │                                                   │
echo   │  4. WICHTIG: Haken setzen bei                    │
echo   │     "Add Python to PATH"  (ganz unten!)          │
echo   │                                                   │
echo   │  5. Dann "Install Now" klicken                   │
echo   │                                                   │
echo   │  6. Danach dieses Skript nochmal doppelklicken   │
echo   └──────────────────────────────────────────────────┘
echo.
set /p OPEN="  Download-Seite jetzt oeffnen? (j/n): "
if /i "%OPEN%"=="j" start https://www.python.org/downloads/
if /i "%OPEN%"=="y" start https://www.python.org/downloads/
echo.
pause
exit /b 1

:found
for /f "tokens=*" %%v in ('%PYTHON% --version 2^>^&1') do echo   + Python gefunden: %%v

:: ---------- Erstinstallation ----------
if not exist "venv" (
    echo.
    echo   Erstinstallation — wird einmalig eingerichtet...
    echo   (Das kann 1-2 Minuten dauern)
    echo.

    %PYTHON% -m venv venv
    if %errorlevel% neq 0 (
        echo   X Konnte virtuelle Umgebung nicht erstellen.
        pause
        exit /b 1
    )

    call venv\Scripts\activate.bat
    pip install --quiet flask python-docx
    if %errorlevel% neq 0 (
        echo   X Pakete konnten nicht installiert werden.
        pause
        exit /b 1
    )

    echo   + Installation abgeschlossen!
) else (
    call venv\Scripts\activate.bat
    echo   + Umgebung geladen
)

echo.
echo   -^>  App startet auf: http://127.0.0.1:5009
echo   -^>  Browser oeffnet sich gleich...
echo.
echo   Zum Beenden: dieses Fenster schliessen
echo   ──────────────────────────────────────
echo.

:: ---------- Browser öffnen ----------
start /b cmd /c "timeout /t 2 /nobreak >nul && start http://127.0.0.1:5009"

:: ---------- App starten ----------
venv\Scripts\python.exe app.py
