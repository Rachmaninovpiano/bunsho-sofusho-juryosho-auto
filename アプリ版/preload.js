/**
 * Preload スクリプト - Electron
 *
 * メインプロセスとレンダラプロセスの橋渡し。
 * contextIsolation: true の環境で安全にAPIを公開。
 */

const { contextBridge, ipcRenderer } = require('electron');

// レンダラに公開するAPI
contextBridge.exposeInMainWorld('electronAPI', {
  // メインプロセスからのファイルオープン通知を受け取る
  onOpenFiles: (callback) => {
    ipcRenderer.on('open-files', (event, filePaths) => callback(filePaths));
  },
  // アプリ情報
  isElectron: true,
  platform: process.platform,
});
