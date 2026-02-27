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

set "SCRIPT_DIR=%~dp0"
set "ZIP_FILE_NAME=excel-markdown-converter-v1.0.zip"
set "TARGET_DIR=%SCRIPT_DIR%Release"

echo 古いファイルがあれば削除します...
if exist "%SCRIPT_DIR%%ZIP_FILE_NAME%" del "%SCRIPT_DIR%%ZIP_FILE_NAME%"
if exist "%TARGET_DIR%" rmdir /s /q "%TARGET_DIR%"

echo 配布用の準備をしています...
mkdir "%TARGET_DIR%"
xcopy /E /I /Y "%SCRIPT_DIR%bat" "%TARGET_DIR%\bat" >nul
xcopy /E /I /Y "%SCRIPT_DIR%scripts" "%TARGET_DIR%\scripts" >nul
copy /Y "%SCRIPT_DIR%配布用_使い方ガイド.md" "%TARGET_DIR%\" >nul

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
