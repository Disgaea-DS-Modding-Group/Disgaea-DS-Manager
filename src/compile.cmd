@echo off
setlocal enabledelayedexpansion

echo ================================
echo   Publishing Avalonia App
echo   .NET 10 Multi-Platform Build
echo ================================
echo.

REM Change this if you want a custom output folder
set OUTDIR=publish

REM Clean previous output
if exist %OUTDIR% (
    echo Cleaning previous publish output...
    rmdir /s /q %OUTDIR%
)

echo.
echo === Windows x86 (framework-dependent) ===
dotnet publish -c Release -r win-x86 ^
  --self-contained false ^
  -p:PublishSingleFile=true ^
  -o %OUTDIR%\win-x86

if errorlevel 1 goto :error

echo.
echo === Linux x64 (self-contained) ===
dotnet publish -c Release -r linux-x64 ^
  --self-contained true ^
  -p:PublishSingleFile=true ^
  -o %OUTDIR%\linux-x64

if errorlevel 1 goto :error

echo.
echo === Linux ARM64 (self-contained) ===
dotnet publish -c Release -r linux-arm64 ^
  --self-contained true ^
  -p:PublishSingleFile=true ^
  -o %OUTDIR%\linux-arm64

if errorlevel 1 goto :error

echo.
echo ================================
echo   ✅ Publish completed successfully
echo   Output folder: %OUTDIR%
echo ================================
goto :eof

:error
echo.
echo ❌ Publish failed!
echo Check the error output above.
exit /b 1