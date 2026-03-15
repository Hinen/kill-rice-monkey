@echo off
setlocal EnableDelayedExpansion

echo ============================================
echo  NOL Mock Ticketing Test
echo ============================================
echo.

if not "%1"=="" (
    set QUEUE_SECONDS=%1
) else (
    set /p QUEUE_SECONDS="Queue duration in seconds (default 60): "
    if "!QUEUE_SECONDS!"=="" set QUEUE_SECONDS=60
)

set EXTRA_FLAGS=
set /p CAPTCHA_YN="Include CAPTCHA? (Y/n, default Y): "
if /i "!CAPTCHA_YN!"=="n" set EXTRA_FLAGS=!EXTRA_FLAGS! --no-captcha

:: Kill any previous mock server
taskkill /FI "WINDOWTITLE eq MockTicketServer" /F > nul 2>&1
timeout /t 1 /nobreak > nul

echo.
echo [1/3] Starting mock server (queue: %QUEUE_SECONDS%s, flags:%EXTRA_FLAGS%, port 8080)...
echo.

start "MockTicketServer" /MIN dotnet run --project "%~dp0MockTicketServer.csproj" -- %QUEUE_SECONDS% %EXTRA_FLAGS%

timeout /t 3 /nobreak > nul

echo [2/3] Starting Chrome (remote-debug + host redirect)...
echo.

set CHROME_PATH=
for %%p in (
    "%ProgramFiles%\Google\Chrome\Application\chrome.exe"
    "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
    "%LocalAppData%\Google\Chrome\Application\chrome.exe"
) do (
    if exist %%p set CHROME_PATH=%%~p
)

if "%CHROME_PATH%"=="" (
    echo [ERROR] Chrome not found.
    pause
    exit /b 1
)

echo Chrome: %CHROME_PATH%
echo.

start "" "%CHROME_PATH%" ^
    --remote-debugging-port=9222 ^
    --host-rules="MAP tickets.interpark.com 127.0.0.1:8080" ^
    --ignore-certificate-errors ^
    --disable-web-security ^
    --user-data-dir="%TEMP%\chrome-nol-mock" ^
    "http://tickets.interpark.com/goods/12345"

echo [3/3] Ready!
echo.
echo ============================================
echo  Config: queue=%QUEUE_SECONDS%s%EXTRA_FLAGS%
echo  Test scenario:
echo  1. Chrome opens NOL Mock page
echo  2. Start KillRiceMonkey program
echo  3. Select NOL template, then run automation
echo     - Date: 2026.04.11
echo     - Round: 1st 19:00
echo  4. Auto: date/round/booking click
echo  5. Queue page waits %QUEUE_SECONDS%s
echo  6. Captcha (if enabled)
echo ============================================
echo.
echo Press any key to stop...
pause > nul

taskkill /FI "WINDOWTITLE eq MockTicketServer" /F > nul 2>&1
