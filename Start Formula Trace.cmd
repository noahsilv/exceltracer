@echo off
setlocal
cd /d "%~dp0"

call :usable
if not errorlevel 1 goto run
for /f "delims=" %%F in ('where node.exe 2^>nul') do (
  call :tryFolder "%%~dpF."
  if not errorlevel 1 goto run
)
for %%D in ("%NVM_SYMLINK%" "%ProgramW6432%\nodejs" "%ProgramFiles%\nodejs" "%ProgramFiles(x86)%\nodejs" "%LocalAppData%\Programs\nodejs" "%LocalAppData%\Volta\bin") do (
  call :tryFolder "%%~D"
  if not errorlevel 1 goto run
)
echo Formula Trace needs a working Node.js 22+ in the same session.
echo If you use a version manager, activate your Node version and start this launcher from that terminal.
echo Commands found on PATH:
where node.exe
pause
exit /b 1

:tryFolder
if "%~1"=="" exit /b 1
if not exist "%~1\node.exe" exit /b 1
set "previousPath=%PATH%"
set "PATH=%~1;%PATH%"
call :usable
if not errorlevel 1 exit /b 0
set "PATH=%previousPath%"
exit /b 1

:usable
node -e "process.exit(Number(process.versions.node.split('.')[0]) >= 22 ? 0 : 1)" >nul 2>&1
exit /b %errorlevel%

:run
echo Running Formula Trace with Node.js
node --version
node scripts\windows.mjs
if errorlevel 1 goto failed
exit /b 0

:failed
echo Formula Trace could not start. See the message above.
pause
exit /b 1
