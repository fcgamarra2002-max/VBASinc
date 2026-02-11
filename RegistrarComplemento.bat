@echo off
setlocal EnableDelayedExpansion

:: --- DEFINICIÓN DE COLORES ---
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do set "ESC=%%b"
set "RED=!ESC![91m"
set "GREEN=!ESC![92m"
set "YELLOW=!ESC![93m"
set "BLUE=!ESC![94m"
set "CYAN=!ESC![96m"
set "WHITE=!ESC![97m"
set "RESET=!ESC![0m"

:: --- AUTO-ELEVACIÓN A ADMINISTRADOR ---
net session >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo !YELLOW!======================================!RESET!
    echo !YELLOW! SOLICITANDO PERMISOS DE ADMINISTRADOR !RESET!
    echo !YELLOW!======================================!RESET!
    powershell -Command "Start-Process -FilePath '%0' -Verb RunAs"
    exit /b
)

echo !CYAN!======================================!RESET!
echo !CYAN! REGISTRO DE VBASinc (Any CPU)        !RESET!
echo !CYAN!======================================!RESET!
echo(

set "DLL_PATH=%~dp0bin\Release\VBASinc.dll"
set "REGASM64=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
set "REGASM32=C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe"

echo !WHITE![1/4] Verificando DLL...!RESET!
if not exist "%DLL_PATH%" (
    echo !RED!ERROR: No se encuentra la DLL en:!RESET!
    echo !RED!%DLL_PATH%!RESET!
    echo Asegurese de haber compilado en modo Release.
    echo(
    pause
    exit /b 1
)
echo !GREEN!OK: DLL encontrada.!RESET!
echo(

echo !WHITE![2/4] Des-registrando versiones anteriores...!RESET!
"%REGASM64%" "%DLL_PATH%" /u >nul 2>&1
"%REGASM32%" "%DLL_PATH%" /u >nul 2>&1

:: Limpieza profunda de claves de registro anteriores en las 4 ubicaciones
reg delete "HKCU\Software\Microsoft\VBA\VBE\6.0\Addins\VBASinc.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64\VBASinc.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\VBA\VBE\7.1\Addins\VBASinc.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\VBA\VBE\7.1\Addins64\VBASinc.Connect" /f >nul 2>&1
echo !GREEN!OK: Limpieza completada.!RESET!
echo(

echo !WHITE![3/4] Registrando COM (32 y 64 bits)...!RESET!
echo Registrando 64-bit...
"%REGASM64%" "%DLL_PATH%" /codebase /tlb >nul
if %ERRORLEVEL% neq 0 (
    echo !RED!ERROR: Fallo el registro COM 64-bit!RESET!
    pause
    exit /b 1
)
echo Registrando 32-bit...
"%REGASM32%" "%DLL_PATH%" /codebase /tlb >nul
if %ERRORLEVEL% neq 0 (
    echo !RED!ERROR: Fallo el registro COM 32-bit!RESET!
    pause
    exit /b 1
)
echo !GREEN!OK: DLL registrada en ambas arquitecturas.!RESET!
echo(

echo !WHITE![4/4] Configurando claves del Add-in (LAS 4 UBICACIONES)...!RESET!

:: 1. VBE 6.0 Addins (32-bit)
reg add "HKCU\Software\Microsoft\VBA\VBE\6.0\Addins\VBASinc.Connect" /v "FriendlyName" /t REG_SZ /d "VBASinc" /f >nul
reg add "HKCU\Software\Microsoft\VBA\VBE\6.0\Addins\VBASinc.Connect" /v "Description" /t REG_SZ /d "Sincronizador VBA" /f >nul
reg add "HKCU\Software\Microsoft\VBA\VBE\6.0\Addins\VBASinc.Connect" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
echo !BLUE!   - VBE 6.0 (32-bit): OK!RESET!

:: 2. VBE 6.0 Addins64 (64-bit)
reg add "HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64\VBASinc.Connect" /v "FriendlyName" /t REG_SZ /d "VBASinc" /f >nul
reg add "HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64\VBASinc.Connect" /v "Description" /t REG_SZ /d "Sincronizador VBA" /f >nul
reg add "HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64\VBASinc.Connect" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
echo !BLUE!   - VBE 6.0 (64-bit): OK!RESET!

:: 3. VBE 7.1 Addins (32-bit)
reg add "HKCU\Software\Microsoft\VBA\VBE\7.1\Addins\VBASinc.Connect" /v "FriendlyName" /t REG_SZ /d "VBASinc" /f >nul
reg add "HKCU\Software\Microsoft\VBA\VBE\7.1\Addins\VBASinc.Connect" /v "Description" /t REG_SZ /d "Sincronizador VBA" /f >nul
reg add "HKCU\Software\Microsoft\VBA\VBE\7.1\Addins\VBASinc.Connect" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
echo !BLUE!   - VBE 7.1 (32-bit): OK!RESET!

:: 4. VBE 7.1 Addins64 (64-bit)
reg add "HKCU\Software\Microsoft\VBA\VBE\7.1\Addins64\VBASinc.Connect" /v "FriendlyName" /t REG_SZ /d "VBASinc" /f >nul
reg add "HKCU\Software\Microsoft\VBA\VBE\7.1\Addins64\VBASinc.Connect" /v "Description" /t REG_SZ /d "Sincronizador VBA" /f >nul
reg add "HKCU\Software\Microsoft\VBA\VBE\7.1\Addins64\VBASinc.Connect" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
echo !BLUE!   - VBE 7.1 (64-bit): OK!RESET!

echo(
echo !GREEN!======================================!RESET!
echo !GREEN!       REGISTRO EXITOSO              !RESET!
echo !GREEN!======================================!RESET!
echo(
pause
