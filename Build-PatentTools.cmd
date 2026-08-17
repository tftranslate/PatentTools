@echo off
setlocal

set "ROOT=%~dp0"

echo.
echo === PatentTools DOTM Build ===
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ROOT%scripts\Build-PatentTools.ps1"

set "EXITCODE=%ERRORLEVEL%"

echo.
if not "%EXITCODE%"=="0" (
    echo BUILD FAILED - exit code: %EXITCODE%
    echo.
    pause
    exit /b %EXITCODE%
)

echo BUILD SUCCESSFUL.
echo Result:
echo %ROOT%build\PatentTools.dotm
echo.
pause
exit /b 0