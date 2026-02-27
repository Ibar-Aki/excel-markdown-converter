@echo off
chcp 65001 >nul
title Markdown → CSV 変換

echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo   📝 Markdown → CSV 変換ツール
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.

if "%~1"=="" (
    echo   ドラッグ＆ドロップで Markdown ファイルを指定するか、
    echo   引数にファイルパスを渡してください。
    echo.
    echo   使い方: md2csv.bat [MarkdownFile] [OutputFile] [--no-pause]
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
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Convert-MarkdownToCsv.ps1" -InputFile "%~1"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\Convert-MarkdownToCsv.ps1" -InputFile "%~1" -OutputFile "%OUTPUT_FILE%"
)

set "EXIT_CODE=%ERRORLEVEL%"
echo.
if "%NO_PAUSE%"=="0" pause
exit /b %EXIT_CODE%
