@echo off
chcp 65001 >nul 2>&1
title 自動でつくる君 - EXE作成中...

echo.
echo  ╔══════════════════════════════════════╗
echo  ║   自動でつくる君 - EXE 作成          ║
echo  ╚══════════════════════════════════════╝
echo.

REM --- Node.js チェック ---
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] Node.js が見つかりません。
    echo      https://nodejs.org/ からインストールしてください。
    start https://nodejs.org/
    pause
    exit /b 1
)

REM --- 依存パッケージ ---
if not exist "node_modules\electron" (
    echo  [*] 依存パッケージをインストール中...
    call npm install --no-fund --no-audit
)

REM --- EXE ビルド（署名スキップ）---
echo  [*] ポータブル版 EXE を作成しています...
echo      ※ 初回は数分かかります。
echo.

set CSC_IDENTITY_AUTO_DISCOVERY=false
call npx electron-builder --win --dir 2>&1

if exist "dist\win-unpacked\自動でつくる君.exe" (
    echo.
    echo  ╔══════════════════════════════════════╗
    echo  ║   EXE 作成完了！                     ║
    echo  ╚══════════════════════════════════════╝
    echo.
    echo  dist\win-unpacked フォルダに
    echo  「自動でつくる君.exe」が作成されました。
    echo.
    echo  このフォルダごとコピーして配布できます。
    echo  （USBメモリ等で持ち運び可能）
    echo.
    explorer "dist\win-unpacked"
) else (
    echo.
    echo  [!] ビルドに失敗しました。
    echo      エラーメッセージを確認してください。
)

pause
