@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0" || exit /b 30
if not "%~2"=="" (
    echo ERROR: Edit pdfs-review.txt; the only BAT option is --no-pause.
    exit /b 30
)
if not "%~1"=="" if /i not "%~1"=="--no-pause" (
    echo ERROR: Edit pdfs-review.txt. For the source CLI use uv run python pdfs_review.py --root "C:\Data\inputs".
    exit /b 30
)
set "POWERSHELL_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%POWERSHELL_EXE%" (
    echo ERROR: Windows PowerShell was not found.
    exit /b 30
)
if not exist "%~dp0support\launcher.ps1" (
    echo ERROR: Download and extract pdfs-review-windows-x64.zip from GitHub Releases.
    echo https://github.com/Grigory-T/pdfs-review/releases/latest
    echo The GitHub source ZIP is for development, not the prepared Windows delivery.
    exit /b 30
)
"%POWERSHELL_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0support\launcher.ps1" -NoPause
set "REVIEW_EXIT=%ERRORLEVEL%"
echo.
if "%REVIEW_EXIT%"=="0" echo COMPLETE - report created beside pdfs-review.bat.
if "%REVIEW_EXIT%"=="2" echo PARTIAL - report saved; review the errors report.
if not "%REVIEW_EXIT%"=="0" if not "%REVIEW_EXIT%"=="2" echo FAILED - read the error above.
if /i "%~1"=="--no-pause" exit /b %REVIEW_EXIT%
echo Press any key to close this window.
pause >nul
exit /b %REVIEW_EXIT%
