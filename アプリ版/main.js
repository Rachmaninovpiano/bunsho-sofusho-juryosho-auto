/**
 * 自動でつくる君 - Electron メインプロセス
 *
 * ブラウザ版と同じWeb資産（app/フォルダ）をElectronウィンドウで表示。
 * ファイルはすべてローカルから読み込まれるため、完全オフラインで動作。
 */

const { app, BrowserWindow, Menu, shell, dialog } = require('electron');
const path = require('path');

// セキュリティ: 開発ツールは本番では無効化
const isDev = process.argv.includes('--dev');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1100,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    title: '自動でつくる君',
    icon: path.join(__dirname, 'app', 'icons', 'icon.svg'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      // ローカルファイルを安全に読み込む
      webSecurity: true,
    },
    // macOS: タイトルバー統合
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    backgroundColor: '#f1f5f9',
    show: false,
  });

  // ウィンドウ準備完了後に表示（白画面チラツキ防止）
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  // ローカルのindex.htmlを読み込み
  mainWindow.loadFile(path.join(__dirname, 'app', 'index.html'));

  // 外部リンクはデフォルトブラウザで開く
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  // 開発モード時のみDevToolsを開く
  if (isDev) {
    mainWindow.webContents.openDevTools();
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// ===== メニューバー =====
function createMenu() {
  const template = [
    {
      label: 'ファイル',
      submenu: [
        {
          label: '新しいPDFを開く',
          accelerator: 'CmdOrCtrl+O',
          click: async () => {
            const result = await dialog.showOpenDialog(mainWindow, {
              properties: ['openFile', 'multiSelections'],
              filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
            });
            if (!result.canceled && result.filePaths.length > 0) {
              mainWindow.webContents.send('open-files', result.filePaths);
            }
          },
        },
        { type: 'separator' },
        { role: 'quit', label: '終了' },
      ],
    },
    {
      label: '表示',
      submenu: [
        { role: 'reload', label: '再読み込み' },
        { role: 'togglefullscreen', label: 'フルスクリーン' },
        ...(isDev ? [
          { type: 'separator' },
          { role: 'toggleDevTools', label: '開発者ツール' },
        ] : []),
      ],
    },
    {
      label: 'ヘルプ',
      submenu: [
        {
          label: 'バージョン情報',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: '自動でつくる君',
              message: '自動でつくる君 v1.0.0',
              detail: '文書送付書・受領書・証拠番号を自動生成するデスクトップアプリ\n\nElectron: ' + process.versions.electron + '\nChromium: ' + process.versions.chrome + '\nNode.js: ' + process.versions.node,
            });
          },
        },
      ],
    },
  ];

  // macOS: アプリメニューを追加
  if (process.platform === 'darwin') {
    template.unshift({
      label: app.name,
      submenu: [
        { role: 'about', label: 'つくる君について' },
        { type: 'separator' },
        { role: 'hide', label: '非表示' },
        { role: 'hideOthers', label: '他を非表示' },
        { role: 'unhide', label: 'すべて表示' },
        { type: 'separator' },
        { role: 'quit', label: '終了' },
      ],
    });
  }

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

// ===== アプリ起動 =====
app.whenReady().then(() => {
  createMenu();
  createWindow();

  // macOS: Dock クリックでウィンドウ再作成
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

// ===== 全ウィンドウ閉じた時 =====
app.on('window-all-closed', () => {
  // macOS以外ではアプリ終了
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
