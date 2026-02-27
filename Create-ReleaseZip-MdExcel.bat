@echo off
chcp 65001 >nul
title 配布用ZIPの作成（MD↔Excel専用版）

echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo   📦 配布用ZIPファイル 自動作成ツール (MD↔Excel専用)
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.
echo md2excel / excel2md に必要な最小構成ファイルだけを取り出してZIPに圧縮します。
echo.

set "SCRIPT_DIR=%~dp0"
set "ZIP_FILE_NAME=md-excel-converter-v1.0.zip"
set "TARGET_DIR=%SCRIPT_DIR%Release_MdExcel"

echo 古いファイルがあれば削除します...
if exist "%SCRIPT_DIR%%ZIP_FILE_NAME%" del "%SCRIPT_DIR%%ZIP_FILE_NAME%"
if exist "%TARGET_DIR%" rmdir /s /q "%TARGET_DIR%"

echo 配布用の準備をしています...
mkdir "%TARGET_DIR%"
mkdir "%TARGET_DIR%\bat"
mkdir "%TARGET_DIR%\scripts"

echo 必要なバッチファイルをコピーしています...
copy /Y "%SCRIPT_DIR%bat\excel2md.bat" "%TARGET_DIR%\bat\" >nul
copy /Y "%SCRIPT_DIR%bat\md2excel.bat" "%TARGET_DIR%\bat\" >nul

echo 必要なPowerShellスクリプトをコピーしています...
copy /Y "%SCRIPT_DIR%scripts\Convert-ExcelToMarkdown.ps1" "%TARGET_DIR%\scripts\" >nul
copy /Y "%SCRIPT_DIR%scripts\Convert-MarkdownToExcel.ps1" "%TARGET_DIR%\scripts\" >nul

echo マニュアルをコピーしています...
copy /Y "%SCRIPT_DIR%配布用_使い方ガイド_MD_Excel.md" "%TARGET_DIR%\使い方ガイド.md" >nul

echo ZIPファイルに圧縮しています...
powershell -NoProfile -Command "Compress-Archive -Path '%TARGET_DIR%\*' -DestinationPath '%SCRIPT_DIR%%ZIP_FILE_NAME%' -Force"

echo 作業用フォルダを削除しています...
rmdir /s /q "%TARGET_DIR%"

echo.
echo ✅ 完成しました！
echo 同じフォルダに【 %ZIP_FILE_NAME% 】が作成されています。
echo これをそのまま相手に渡してください。
echo.
pause
