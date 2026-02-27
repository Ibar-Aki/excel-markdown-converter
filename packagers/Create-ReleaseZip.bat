@echo off
chcp 65001 >nul
title 配布用ZIPの作成

echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo   📦 配布用ZIPファイル 自動作成ツール
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.
echo 現在のフォルダから、他の方へ配るための「bat」と「scripts」フォルダ、そして
echo 「配布用_使い方ガイド.md」だけを取り出してZIPに圧縮します。
echo.

set "TOOL_DIR=%~dp0"
set "ROOT_DIR=%~dp0.."
set "ZIP_FILE_NAME=excel-markdown-converter-v1.0.zip"
set "OUTPUT_DIR=%ROOT_DIR%\releases"
set "TARGET_DIR=%TOOL_DIR%Release"

echo 古いファイルがあれば削除します...
if exist "%OUTPUT_DIR%\%ZIP_FILE_NAME%" del "%OUTPUT_DIR%\%ZIP_FILE_NAME%"
if exist "%TARGET_DIR%" rmdir /s /q "%TARGET_DIR%"

echo 配布用の準備をしています...
mkdir "%TARGET_DIR%"
xcopy /E /I /Y "%ROOT_DIR%\bat" "%TARGET_DIR%\bat" >nul
xcopy /E /I /Y "%ROOT_DIR%\scripts" "%TARGET_DIR%\scripts" >nul
copy /Y "%TOOL_DIR%配布用_使い方ガイド.md" "%TARGET_DIR%\" >nul

echo ZIPファイルに圧縮しています...
powershell -NoProfile -Command "Compress-Archive -Path '%TARGET_DIR%\*' -DestinationPath '%OUTPUT_DIR%\%ZIP_FILE_NAME%' -Force"

echo 作業用フォルダを削除しています...
rmdir /s /q "%TARGET_DIR%"

echo.
echo ✅ 完成しました！
echo releases フォルダに【 %ZIP_FILE_NAME% 】が作成されています。
echo これをそのまま相手に渡してください。
echo.
pause
