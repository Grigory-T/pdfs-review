@echo off
setlocal
cd /d "%~dp0" || exit /b 1

if not exist ".runtime\temp" mkdir ".runtime\temp"
set "TEMP=%CD%\.runtime\temp"
set "TMP=%CD%\.runtime\temp"
set "UV_CACHE_DIR=%CD%\.runtime\uv-cache"
set "UV_EXE="
if exist "%~dp0uv.exe" set "UV_EXE=%~dp0uv.exe"
if not defined UV_EXE for /f "delims=" %%U in ('where uv.exe 2^>nul') do if not defined UV_EXE set "UV_EXE=%%U"
if not defined UV_EXE if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV_EXE=%USERPROFILE%\.local\bin\uv.exe"
if not defined UV_EXE for /d %%P in ("%APPDATA%\Python\Python*") do if exist "%%P\Scripts\uv.exe" if not defined UV_EXE set "UV_EXE=%%P\Scripts\uv.exe"
if not defined UV_EXE for /d %%P in ("%LOCALAPPDATA%\Programs\Python\Python*") do if exist "%%P\Scripts\uv.exe" if not defined UV_EXE set "UV_EXE=%%P\Scripts\uv.exe"
if not defined UV_EXE (
  echo ERROR: Installed uv.exe was not found.
  exit /b 1
)

"%UV_EXE%" sync --frozen --no-dev --no-python-downloads
if errorlevel 1 exit /b 1
".venv\Scripts\python.exe" -u "pdfs_review.py" %*
set "RUN_EXIT=%ERRORLEVEL%"
if "%RUN_EXIT%"=="2" echo Workbook saved. Review the errors report.
exit /b %RUN_EXIT%
