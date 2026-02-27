@echo off
chcp 65001 >nul
title CSV → Markdown 変換

echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo   📄 CSV → Markdown 変換ツール
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.

if "%~1"=="" (
    echo   ドラッグ＆ドロップで CSV ファイルを指定するか、
    echo   引数にファイルパスを渡してください。
    echo.
    echo   使い方: csv2md.bat [CsvFile] [OutputFile] [--no-pause]
    echo.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%~dp0..\scripts"
set "NO_PAUSE=0"
set "OUTPUT_FILE="

if /I "%~2"=="--no-pause" (
    set "NO_PAUSE=1"
) else (
    set "OUTPUT_FILE=%~2"
)

if /I "%~3"=="--no-pause" set "NO_PAUSE=1"

if "%OUTPUT_FILE%"=="" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Convert-CsvToMarkdown.ps1" -InputFile "%~1"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Convert-CsvToMarkdown.ps1" -InputFile "%~1" -OutputFile "%OUTPUT_FILE%"
)

set "EXIT_CODE=%ERRORLEVEL%"
echo.
if "%NO_PAUSE%"=="0" pause
exit /b %EXIT_CODE%
