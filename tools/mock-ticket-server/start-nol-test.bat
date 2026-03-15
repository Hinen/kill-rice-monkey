@echo off
setlocal

echo ============================================
echo  NOL Mock 티켓팅 테스트 환경 시작
echo ============================================
echo.

:: 대기열 시간 (초) - 기본 60초
set QUEUE_SECONDS=60
if not "%1"=="" set QUEUE_SECONDS=%1

echo [1/3] Mock 서버 시작 (대기열: %QUEUE_SECONDS%초)...
echo       관리자 권한이 필요합니다 (포트 80 사용)
echo.

:: Mock 서버를 백그라운드로 시작
start "MockTicketServer" /MIN dotnet run --project "%~dp0MockTicketServer.csproj" -- %QUEUE_SECONDS%

:: 서버 시작 대기
timeout /t 3 /nobreak > nul

echo [2/3] Chrome 시작 (remote-debug + host 리다이렉트)...
echo.

:: Chrome 경로 찾기
set CHROME_PATH=
for %%p in (
    "%ProgramFiles%\Google\Chrome\Application\chrome.exe"
    "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
    "%LocalAppData%\Google\Chrome\Application\chrome.exe"
) do (
    if exist %%p set CHROME_PATH=%%~p
)

if "%CHROME_PATH%"=="" (
    echo [오류] Chrome을 찾을 수 없습니다.
    pause
    exit /b 1
)

echo Chrome 경로: %CHROME_PATH%
echo.

:: Chrome 실행 - NOL 도메인을 localhost:8080으로 리다이렉트
start "" "%CHROME_PATH%" ^
    --remote-debugging-port=9222 ^
    --host-rules="MAP tickets.interpark.com 127.0.0.1:8080" ^
    --ignore-certificate-errors ^
    --disable-web-security ^
    --user-data-dir="%TEMP%\chrome-nol-mock" ^
    "http://tickets.interpark.com/goods/12345"

echo [3/3] 준비 완료!
echo.
echo ============================================
echo  테스트 방법:
echo  1. Chrome에서 NOL Mock 페이지가 열립니다
echo  2. KillRiceMonkey 프로그램을 시작하세요
echo  3. NOL 템플릿 선택 후 자동화를 실행하세요
echo     - 관람일: 2026.04.11
echo     - 회차: 1회 19:00
echo  4. 프로그램이 날짜/회차/예매를 순서대로 처리합니다
echo  5. 대기열 %QUEUE_SECONDS%초 대기 후 캡차 페이지로 이동
echo ============================================
echo.
echo 종료하려면 아무 키나 누르세요...
pause > nul

:: 정리
taskkill /FI "WINDOWTITLE eq MockTicketServer" /F > nul 2>&1
