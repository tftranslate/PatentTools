@echo off
setlocal

set "ROOT=%~dp0"

echo.
echo === PatentTools DOTM Build ===
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ROOT%scripts\Build-PatentTools.ps1"

set "EXITCODE=%ERRORLEVEL%"

if not "%EXITCODE%"=="0" (
    echo.
    echo DOTM BUILD FEHLGESCHLAGEN - Exit-Code: %EXITCODE%
    echo.
    pause
    exit /b %EXITCODE%
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ROOT%scripts\Build-Installer.ps1"

set "EXITCODE=%ERRORLEVEL%"

if not "%EXITCODE%"=="0" (
    echo.
    echo INSTALLER BUILD FEHLGESCHLAGEN - Exit-Code: %EXITCODE%
    echo.
    pause
    exit /b %EXITCODE%
)


echo.
echo BUILD ERFOLGREICH.
echo Ergebnis:
echo %ROOT%build\PatentTools.dotm
echo %ROOT%build\PatentTools_*_Setup.exe
echo.
pause
exit /b 0