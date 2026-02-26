/**
 * 文書送付書 自動生成 Web サーバー
 *
 * 起動: node web/server.js
 * URL: http://localhost:3000
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

// generate.js からモジュールインポート
const generator = require('../system/generate.js');
// receipt.js からモジュールインポート
const receipt = require('../system/receipt.js');

// formidable (ESM対応: v3はdynamic importが必要)
let formidable;

const BASE_DIR = path.resolve(__dirname, '..');
const PUBLIC_DIR = path.join(__dirname, 'public');
const UPLOAD_DIR = path.join(BASE_DIR, 'temp_uploads');
const OUTPUT_DIR = path.join(BASE_DIR, 'output');

// config.json からポート番号を読み込み
let PORT = 3000;
try {
  const cfg = JSON.parse(fs.readFileSync(path.join(BASE_DIR, 'config.json'), 'utf-8'));
  if (cfg.port) PORT = cfg.port;
} catch (e) { /* デフォルト3000 */ }

// セッション管理（メモリ内）
const sessions = new Map();

// セッション自動クリーンアップ（30分経過で削除）
setInterval(() => {
  const now = Date.now();
  for (const [id, session] of sessions) {
    if (now - session.createdAt > 30 * 60 * 1000) {
      try {
        if (session.uploadPath && fs.existsSync(session.uploadPath)) {
          fs.unlinkSync(session.uploadPath);
        }
      } catch (e) { /* ignore */ }
      sessions.delete(id);
    }
  }
}, 5 * 60 * 1000);

// MIMEタイプ
const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.svg': 'image/svg+xml',
  '.woff2': 'font/woff2',
  '.woff': 'font/woff',
};

// ===== 静的ファイル配信 =====
function serveStatic(req, res) {
  let filePath = req.url === '/' ? '/index.html' : req.url;
  filePath = filePath.split('?')[0];
  const fullPath = path.join(PUBLIC_DIR, filePath);

  if (!fullPath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  if (!fs.existsSync(fullPath)) {
    res.writeHead(404);
    res.end('Not Found');
    return;
  }

  const ext = path.extname(fullPath);
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';
  const data = fs.readFileSync(fullPath);
  res.writeHead(200, {
    'Content-Type': contentType,
    'Cache-Control': 'no-cache, no-store, must-revalidate',
  });
  res.end(data);
}

// ===== JSON レスポンス =====
function sendJSON(res, statusCode, data) {
  res.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(data));
}

// ===== API: PDFアップロード＆解析 =====
async function handleUpload(req, res) {
  try {
    if (!formidable) {
      const mod = await import('formidable');
      formidable = mod.formidable || mod.default || mod;
    }

    if (!fs.existsSync(UPLOAD_DIR)) {
      fs.mkdirSync(UPLOAD_DIR, { recursive: true });
    }

    const form = formidable({
      uploadDir: UPLOAD_DIR,
      keepExtensions: true,
      maxFileSize: 50 * 1024 * 1024, // 50MB
      // filterを削除: D&D時にmimetypeがoctet-streamになるブラウザがある
    });

    const [fields, files] = await form.parse(req);

    // formidable v3: files は { fieldName: [File, ...] } の形式
    const uploadedFile = files.pdf?.[0];
    if (!uploadedFile) {
      sendJSON(res, 400, { error: 'PDFファイルが見つかりません。pdf フィールドにファイルを添付してください。' });
      return;
    }

    const pdfPath = uploadedFile.filepath;
    const originalName = uploadedFile.originalFilename || 'unknown.pdf';

    // 拡張子チェック（filterの代替）
    if (!originalName.toLowerCase().endsWith('.pdf')) {
      // アップロードされたファイルを削除
      try { fs.unlinkSync(pdfPath); } catch (e) { /* ignore */ }
      sendJSON(res, 400, { error: 'PDFファイルのみアップロードできます。' });
      return;
    }

    console.log(`[Upload] ${originalName} -> ${pdfPath}`);

    // PDFからテキスト抽出
    console.log('[OCR] テキスト抽出開始...');
    const pdfText = await generator.extractTextFromPDF(pdfPath);

    // 情報抽出
    console.log('[Extract] 情報抽出開始...');
    const info = generator.extractInfoFromText(pdfText);

    // 送付書類名（ファイル名から自動取得）
    const documentTitle = generator.getDocumentTitleFromFilename(originalName);

    // セッション作成
    const sessionId = crypto.randomUUID();
    sessions.set(sessionId, {
      createdAt: Date.now(),
      uploadPath: pdfPath,
      originalName,
      info,
      documentTitle,
    });

    console.log(`[Session] ${sessionId} 作成完了`);
    console.log('[Result]', JSON.stringify(info, null, 2));

    sendJSON(res, 200, {
      sessionId,
      info,
      documentTitle,
      originalName,
    });

  } catch (err) {
    console.error('[Upload Error]', err);
    sendJSON(res, 500, { error: `PDF処理中にエラーが発生しました: ${err.message}` });
  }
}

// ===== API: Word文書生成 =====
async function handleGenerate(req, res) {
  try {
    const body = await new Promise((resolve, reject) => {
      let data = '';
      req.on('data', chunk => { data += chunk; });
      req.on('end', () => {
        try { resolve(JSON.parse(data)); } catch (e) { reject(new Error('Invalid JSON')); }
      });
      req.on('error', reject);
    });

    const { sessionId, info, documentTitle } = body;

    if (!sessionId || !sessions.has(sessionId)) {
      sendJSON(res, 400, { error: 'セッションが見つかりません。もう一度アップロードしてください。' });
      return;
    }

    const session = sessions.get(sessionId);

    console.log(`[Generate] Session ${sessionId}`);
    console.log('[Generate Info]', JSON.stringify(info, null, 2));

    const result = await generator.generateDocumentFromInfo(
      info,
      documentTitle || session.documentTitle,
      OUTPUT_DIR
    );

    session.outputPath = result.outputPath;
    session.outputFileName = result.outputFileName;

    console.log(`[Generate] 完了: ${result.outputPath}`);

    sendJSON(res, 200, {
      sessionId,
      fileName: result.outputFileName,
      downloadUrl: `/api/download/${sessionId}`,
    });

  } catch (err) {
    console.error('[Generate Error]', err);
    sendJSON(res, 500, { error: `文書生成中にエラーが発生しました: ${err.message}` });
  }
}

// ===== API: 受領書生成 =====
async function handleReceiptGenerate(req, res) {
  try {
    if (!formidable) {
      const mod = await import('formidable');
      formidable = mod.formidable || mod.default || mod;
    }

    if (!fs.existsSync(UPLOAD_DIR)) {
      fs.mkdirSync(UPLOAD_DIR, { recursive: true });
    }

    const form = formidable({
      uploadDir: UPLOAD_DIR,
      keepExtensions: true,
      maxFileSize: 50 * 1024 * 1024,
    });

    const [fields, files] = await form.parse(req);

    const uploadedFile = files.pdf?.[0];
    if (!uploadedFile) {
      sendJSON(res, 400, { error: 'PDFファイルが見つかりません。' });
      return;
    }

    const pdfPath = uploadedFile.filepath;
    const originalName = uploadedFile.originalFilename || 'unknown.pdf';

    if (!originalName.toLowerCase().endsWith('.pdf')) {
      try { fs.unlinkSync(pdfPath); } catch (e) { /* ignore */ }
      sendJSON(res, 400, { error: 'PDFファイルのみアップロードできます。' });
      return;
    }

    console.log(`[Receipt Upload] ${originalName} -> ${pdfPath}`);

    // フォームデータからオプションを取得
    const signerTitle = fields.signerTitle?.[0] || '被告訴訟代理人';
    const signerName = fields.signerName?.[0] || '大元和貴';
    const receiptDate = fields.receiptDate?.[0] || undefined; // undefinedなら今日

    console.log(`[Receipt] 署名: ${signerTitle} ${signerName}`);

    // 受領書生成
    const result = await receipt.generateReceipt(pdfPath, {
      receiptDate,
      signerTitle,
      signerName,
      outputDir: OUTPUT_DIR,
    });

    // セッション作成（ダウンロード用）
    const sessionId = crypto.randomUUID();
    sessions.set(sessionId, {
      createdAt: Date.now(),
      uploadPath: pdfPath,
      originalName,
      outputPath: result.outputPath,
      outputFileName: result.outputFileName,
    });

    console.log(`[Receipt] 完了: ${result.outputPath}`);

    sendJSON(res, 200, {
      sessionId,
      fileName: result.outputFileName,
      downloadUrl: `/api/download/${sessionId}`,
    });

  } catch (err) {
    console.error('[Receipt Error]', err);
    sendJSON(res, 500, { error: `受領書生成中にエラーが発生しました: ${err.message}` });
  }
}

// ===== API: ファイルダウンロード =====
function handleDownload(req, res, sessionId) {
  const session = sessions.get(sessionId);
  if (!session || !session.outputPath) {
    res.writeHead(404);
    res.end('File not found');
    return;
  }

  if (!fs.existsSync(session.outputPath)) {
    res.writeHead(404);
    res.end('File not found on disk');
    return;
  }

  const fileData = fs.readFileSync(session.outputPath);
  const fileName = session.outputFileName;
  const encodedName = encodeURIComponent(fileName).replace(/'/g, '%27');

  // ファイル拡張子でContent-Typeを決定
  const ext = path.extname(fileName).toLowerCase();
  const contentType = ext === '.pdf'
    ? 'application/pdf'
    : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  res.writeHead(200, {
    'Content-Type': contentType,
    'Content-Disposition': `attachment; filename*=UTF-8''${encodedName}`,
    'Content-Length': fileData.length,
  });
  res.end(fileData);
}

// ===== リクエストルーター =====
async function handleRequest(req, res) {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const pathname = url.pathname;

  console.log(`[${new Date().toISOString()}] ${req.method} ${pathname}`);

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }

  try {
    if (req.method === 'POST' && pathname === '/api/upload') {
      await handleUpload(req, res);
    } else if (req.method === 'POST' && pathname === '/api/generate') {
      await handleGenerate(req, res);
    } else if (req.method === 'POST' && pathname === '/api/receipt') {
      await handleReceiptGenerate(req, res);
    } else if (req.method === 'GET' && pathname.startsWith('/api/download/')) {
      const sessionId = pathname.replace('/api/download/', '');
      handleDownload(req, res, sessionId);
    } else if (req.method === 'GET') {
      serveStatic(req, res);
    } else {
      res.writeHead(405);
      res.end('Method Not Allowed');
    }
  } catch (err) {
    console.error('[Server Error]', err);
    sendJSON(res, 500, { error: 'Internal Server Error' });
  }
}

// ===== サーバー起動 =====
const server = http.createServer(handleRequest);

// OCR処理は時間がかかるため、タイムアウトを延長（5分）
server.timeout = 300000;
server.requestTimeout = 300000;

server.listen(PORT, () => {
  console.log('');
  console.log('====================================================');
  console.log('  文書送付書自動でつくる君');
  console.log('====================================================');
  console.log('');
  console.log(`  ブラウザでアクセス: http://localhost:${PORT}`);
  console.log('');
  console.log('  終了するには Ctrl+C を押してください');
  console.log('====================================================');
  console.log('');
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`ポート ${PORT} は既に使用されています。`);
    console.error(`他のプログラムがポート${PORT}を使用していないか確認してください。`);
    process.exit(1);
  }
  throw err;
});
