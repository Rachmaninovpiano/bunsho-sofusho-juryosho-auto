@echo off
chcp 65001 >nul 2>&1
title 自動でつくる君 - 起動中...

echo.
echo  ╔══════════════════════════════════════╗
echo  ║   自動でつくる君 - デスクトップ版    ║
echo  ╚══════════════════════════════════════╝
echo.

REM --- Node.js チェック ---
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] Node.js が見つかりません。
    echo.
    echo  Node.js を以下からインストールしてください:
    echo  https://nodejs.org/
    echo.
    echo  インストール後、このファイルをもう一度ダブルクリックしてください。
    echo.
    start https://nodejs.org/
    pause
    exit /b 1
)

echo  [OK] Node.js 検出
for /f "tokens=*" %%v in ('node -v') do echo       バージョン: %%v

REM --- 依存パッケージ自動インストール ---
if not exist "node_modules\electron" (
    echo.
    echo  [*] 初回セットアップ中（Electron をインストール）...
    echo      ※ 初回のみ数分かかります。次回からは即起動します。
    echo.
    call npm install --no-fund --no-audit
    if %errorlevel% neq 0 (
        echo.
        echo  [!] インストールに失敗しました。
        echo      ネットワーク接続を確認してください。
        pause
        exit /b 1
    )
    echo.
    echo  [OK] インストール完了
)

REM --- アプリ起動 ---
echo.
echo  [*] アプリを起動しています...
echo      （このウィンドウは閉じないでください）
echo.

call npx electron . 2>nul

echo.
echo  アプリを終了しました。
timeout /t 3 >nul
