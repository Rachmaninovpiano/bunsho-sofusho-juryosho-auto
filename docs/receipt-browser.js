/**
 * 受領書自動でつくる君 - ブラウザ版受領書生成モジュール
 *
 * pdf.js + Tesseract.js + pdf-lib で全てブラウザ内で処理
 * サーバー不要・インストール不要
 */

// ===== グローバル変数（CDNから読み込んだライブラリ）=====
// PDFLib, fontkit, pdfjsLib, Tesseract は index.html の <script> で読み込み済み

// ===== 設定（localStorageから取得）=====
function getConfig() {
  try {
    return JSON.parse(localStorage.getItem('tsukurukun_config') || '{}');
  } catch (e) { return {}; }
}

// ===== CMap URL（日本語PDFのテキスト抽出に必須）=====
const RECEIPT_CMAP_URL = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/';

// ===== pdf.js でPDFページをcanvasに描画 → Tesseract.jsでOCR =====

/**
 * PDFファイル(ArrayBuffer)の指定ページをOCR
 * @param {ArrayBuffer} pdfArrayBuffer
 * @param {number} pageNum - 1始まり
 * @param {Function} onProgress - 進捗コールバック (msg)
 * @returns {Promise<{ words, imgWidth, imgHeight }>}
 */
async function runOcrBrowser(pdfArrayBuffer, pageNum, onProgress) {
  onProgress && onProgress(`ページ${pageNum}を描画中...`);

  // pdf.js でページをcanvasに描画（CMap設定で日本語対応）
  const pdfDoc = await pdfjsLib.getDocument({
    data: pdfArrayBuffer,
    cMapUrl: RECEIPT_CMAP_URL,
    cMapPacked: true,
  }).promise;
  const page = await pdfDoc.getPage(pageNum);

  // 300dpi相当のスケール
  const viewport = page.getViewport({ scale: 300 / 72 });
  const canvas = document.createElement('canvas');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  const ctx = canvas.getContext('2d');

  await page.render({ canvasContext: ctx, viewport }).promise;

  const imgWidth = canvas.width;
  const imgHeight = canvas.height;

  onProgress && onProgress(`ページ${pageNum}をOCR中...`);

  // Tesseract.js でOCR（HOCR出力）
  const worker = await Tesseract.createWorker('jpn', 1, {
    logger: m => {
      if (m.status === 'recognizing text' && onProgress) {
        onProgress(`OCR処理中... ${Math.round((m.progress || 0) * 100)}%`);
      }
    }
  });

  const { data } = await worker.recognize(canvas);
  await worker.terminate();

  // Tesseract.js の結果からword座標を抽出
  const words = [];
  if (data && data.words) {
    for (const w of data.words) {
      const text = w.text.trim();
      if (!text) continue;
      const bbox = w.bbox;
      words.push({
        x1: bbox.x0,
        y1: bbox.y0,
        x2: bbox.x1,
        y2: bbox.y1,
        text
      });
    }
  }

  console.log(`  [OCR] ページ${pageNum}: ${words.length}語検出`);
  return { words, imgWidth, imgHeight };
}

// ===== ピクセル座標 → PDF座標（左下原点）変換 =====
function px2pdf(px, py, imgW, imgH, pgW, pgH) {
  return {
    x: px * pgW / imgW,
    y: pgH - (py * pgH / imgH),
  };
}

// ===== 受領書ページ特定ロジック（サーバー版と同一）=====

function findReceiptLabel(words) {
  const direct = words.find(w =>
    w.text.includes('受領書') || w.text.includes('受領')
  );
  if (direct) return { found: true, y: direct.y1 };

  const juWords = words.filter(w => w.text === '受');
  for (const ju of juWords) {
    const ryou = words.find(w =>
      w.text === '領' &&
      Math.abs(w.y1 - ju.y1) < 50 &&
      w.x1 > ju.x1 && w.x1 < ju.x1 + 300
    );
    if (ryou) return { found: true, y: ju.y1 };
  }

  return { found: false, y: null };
}

function scoreReceiptPage(words) {
  const allText = words.map(w => w.text).join('');
  let score = 0;

  const receiptLabel = findReceiptLabel(words);
  if (receiptLabel.found) score += 50;
  if (/令和/.test(allText))  score += 10;
  if (/代理人/.test(allText)) score += 10;
  if (words.length > 200) score -= 20;
  if (words.length > 300) score -= 20;

  return score;
}

async function findReceiptPage(pdfArrayBuffer, totalPages, onProgress) {
  console.log(`  🔍 受領書ページ特定中... (全${totalPages}ページ)`);

  if (totalPages === 1) {
    const ocr = await runOcrBrowser(pdfArrayBuffer, 1, onProgress);
    return { pageNum: 1, ocr };
  }

  const scanOrder = [1, totalPages];
  for (let p = 2; p < totalPages; p++) scanOrder.push(p);

  let bestPageNum = 1;
  let bestOcr = null;
  let bestScore = -Infinity;

  for (const p of scanOrder) {
    onProgress && onProgress(`ページ ${p}/${totalPages} スキャン中...`);
    const ocr = await runOcrBrowser(pdfArrayBuffer, p, onProgress);
    const score = scoreReceiptPage(ocr.words);
    console.log(`    → ページ${p} スコア: ${score}`);

    if (score > bestScore) {
      bestScore = score;
      bestPageNum = p;
      bestOcr = ocr;
    }

    if (score >= 50) {
      console.log(`  ✅ 受領書ページ確定: ページ${p}`);
      return { pageNum: p, ocr };
    }
  }

  return { pageNum: bestPageNum, ocr: bestOcr };
}

// ===== 受領書セクション検出（サーバー版と同一）=====

function findReceiptSectionStart(words, imgH) {
  const receiptLabel = findReceiptLabel(words);
  if (receiptLabel.found && receiptLabel.y !== null) {
    return Math.max(0, receiptLabel.y - 50);
  }
  return 0;
}

function detectPositions(words, imgW, imgH, pgW, pgH) {
  const receiptStartY = findReceiptSectionStart(words, imgH);
  const rw = words.filter(w => w.y1 >= receiptStartY);

  // 1. 「行」の検出（OCR完全一致のみ）
  let gyouWord = null;
  const allGyouWords = rw.filter(w => w.text === '行');

  if (allGyouWords.length > 0) {
    const bengoWordInReceipt = rw.find(w =>
      w.text.includes('弁護') || w.text.includes('護士')
    );
    if (bengoWordInReceipt) {
      const bengoY = bengoWordInReceipt.y1;
      const gyouInBengoLine = allGyouWords.filter(w => Math.abs(w.y1 - bengoY) < 60);
      if (gyouInBengoLine.length > 0) {
        gyouWord = gyouInBengoLine.reduce((a, b) => a.x1 > b.x1 ? a : b);
      }
    }
    if (!gyouWord) {
      gyouWord = allGyouWords.reduce((a, b) => a.y1 < b.y1 ? a : b);
    }
  }

  // 2. 署名行の検出（先に検出 → 日付の検索範囲を限定）
  let agentWord = null;
  const agentCandidates = rw.filter(w =>
    w.text.includes('代理人') || w.text.includes('代理')
  );
  if (agentCandidates.length > 0) {
    agentWord = agentCandidates.reduce((a, b) => a.y1 > b.y1 ? a : b);
  } else {
    const midY = receiptStartY + (imgH - receiptStartY) * 0.5;
    const agentCandidates2 = rw.filter(w =>
      w.y1 > midY && (w.text.includes('被告') || w.text.includes('原告'))
    );
    if (agentCandidates2.length > 0) {
      agentWord = agentCandidates2.reduce((a, b) => a.y1 > b.y1 ? a : b);
    }
  }

  if (agentWord) {
    const agentY = agentWord.y1;
    const agentRowWords = rw.filter(w => Math.abs(w.y1 - agentY) < 60).sort((a, b) => a.x1 - b.x1);
    const ninWord = agentRowWords.find(w => w.text === '人' || w.text.endsWith('人'));
    if (ninWord) agentWord._titleEndX = ninWord.x2;
    const leftMost = agentRowWords[0];
    if (leftMost && leftMost.x1 < agentWord.x1) agentWord._lineStartX = leftMost.x1;
  }

  // 3. 日付欄の検出
  const searchTopY = gyouWord ? gyouWord.y1 + 20 : receiptStartY;
  const searchBottomY = agentWord ? agentWord.y1 - 10 : imgH;

  let dateWord = null;
  const reiwaWords = rw.filter(w =>
    w.y1 > searchTopY && w.y1 < searchBottomY &&
    (w.text === '令' || w.text === '令和' || w.text.startsWith('令'))
  );
  if (reiwaWords.length > 0) {
    dateWord = reiwaWords.reduce((a, b) => a.y1 < b.y1 ? a : b);
  } else {
    const dateish = rw.filter(w =>
      w.y1 > searchTopY && w.y1 < searchBottomY &&
      (w.text.includes('年') || w.text.includes('月'))
    );
    if (dateish.length > 0) {
      dateWord = dateish.reduce((a, b) => a.y1 < b.y1 ? a : b);
    }
  }

  // フォールバック推定
  if (!dateWord && agentWord) {
    const estimatedY = Math.round(agentWord.y1 - (imgH * 0.06));
    const estimatedX = Math.round(imgW * 0.05);
    dateWord = { x1: estimatedX, y1: estimatedY, x2: estimatedX + 200, y2: estimatedY + 40, text: '令和（推定）', estimated: true };
  } else if (!dateWord) {
    const estimatedY = Math.round(receiptStartY + (imgH - receiptStartY) * 0.5);
    const estimatedX = Math.round(imgW * 0.05);
    dateWord = { x1: estimatedX, y1: estimatedY, x2: estimatedX + 200, y2: estimatedY + 40, text: '令和（推定）', estimated: true };
  }
  if (!agentWord) {
    const estimatedY = Math.round(receiptStartY + (imgH - receiptStartY) * 0.85);
    const estimatedX = Math.round(imgW * 0.20);
    agentWord = { x1: estimatedX, y1: estimatedY, x2: estimatedX + 300, y2: estimatedY + 40, text: '代理人（推定）', estimated: true };
  }

  // PDF座標に変換
  const gyouPdf = gyouWord ? {
    left:   px2pdf(gyouWord.x1, gyouWord.y2, imgW, imgH, pgW, pgH),
    right:  px2pdf(gyouWord.x2, gyouWord.y2, imgW, imgH, pgW, pgH),
    top:    px2pdf(gyouWord.x1, gyouWord.y1, imgW, imgH, pgW, pgH).y,
    width:  (gyouWord.x2 - gyouWord.x1) * pgW / imgW,
    height: (gyouWord.y2 - gyouWord.y1) * pgH / imgH,
    pxY1:   gyouWord.y1,
  } : null;

  const datePdfTop    = px2pdf(dateWord.x1, dateWord.y1, imgW, imgH, pgW, pgH);
  const datePdfBottom = px2pdf(dateWord.x1, dateWord.y2, imgW, imgH, pgW, pgH);
  const agentPdfTop    = px2pdf(agentWord.x1, agentWord.y1, imgW, imgH, pgW, pgH);
  const agentPdfBottom = px2pdf(agentWord.x1, agentWord.y2, imgW, imgH, pgW, pgH);
  const agentLinePdf = agentPdfBottom;

  const agentTitleEndX = agentWord && agentWord._titleEndX
    ? agentWord._titleEndX * pgW / imgW : null;

  return {
    gyou: gyouPdf,
    date: { x: datePdfBottom.x, yTop: datePdfTop.y, yBase: datePdfBottom.y },
    agent: { x: agentLinePdf.x, yTop: agentPdfTop.y, yBase: agentPdfBottom.y },
    agentTitleEndX,
    agentRowY: agentPdfBottom.y,
    imgW, imgH,
  };
}

// ===== メイン: 受領書PDF生成（ブラウザ版）=====

/**
 * ブラウザ内で受領書PDFを生成
 * @param {File} file - PDFファイル（File API）
 * @param {Object} options - { receiptDate, signerTitle, signerName }
 * @param {Function} onProgress - 進捗コールバック (msg)
 * @returns {Promise<{ blob: Blob, fileName: string }>}
 */
async function generateReceiptBrowser(file, options = {}, onProgress) {
  console.log('\n📄 受領書生成開始:', file.name);
  onProgress && onProgress('PDFを読み込み中...');

  // デフォルト値
  const today = new Date();
  const reiwaYear = today.getFullYear() - 2018;
  const defaultDate = `令和${reiwaYear}年${today.getMonth() + 1}月${today.getDate()}日`;

  const config = getConfig();
  const receiptDate = options.receiptDate || defaultDate;
  const signerTitle = options.signerTitle || '被告訴訟代理人';
  const signerName  = options.signerName  || config.signerName || '山田太郎';

  // PDF読み込み
  const pdfArrayBuffer = await file.arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(pdfArrayBuffer);
  pdfDoc.registerFontkit(fontkit);

  const totalPages = pdfDoc.getPageCount();
  console.log(`  📑 総ページ数: ${totalPages}`);

  // 受領書ページ特定
  onProgress && onProgress('受領書ページを探しています...');
  const { pageNum: receiptPageNum, ocr } = await findReceiptPage(pdfArrayBuffer, totalPages, onProgress);
  const receiptPageIndex = receiptPageNum - 1;

  const words = ocr.words;
  const imgWidth = ocr.imgWidth;
  const imgHeight = ocr.imgHeight;

  const page = pdfDoc.getPage(receiptPageIndex);
  const { width: pgW, height: pgH } = page.getSize();

  // フォント読み込み（Noto Serif JP）
  onProgress && onProgress('フォントを読み込み中...');
  const fontResp = await fetch('fonts/NotoSerifJP.ttf');
  const fontBytes = await fontResp.arrayBuffer();
  const font = await pdfDoc.embedFont(fontBytes, { subset: false });

  // グリフ事前登録
  const allChars = `行先生${receiptDate}${signerTitle}　${signerName}㊞`;
  try { font.encodeText(allChars); } catch (e) { /* ignore */ }

  // 位置検出
  onProgress && onProgress('書き込み位置を検出中...');
  const pos = detectPositions(words, imgWidth, imgHeight, pgW, pgH);

  const fs_ = 10.5;
  const { rgb } = PDFLib;

  // 1. 「行」→ 二重打消し線 + 「先生」
  if (pos.gyou) {
    const g = pos.gyou;
    const gyouOcrW = g.width;
    const gyouCharW = font.widthOfTextAtSize('行', fs_);
    const strikeW = Math.min(gyouOcrW, gyouCharW);
    const midY = g.left.y + fs_ * 0.40;
    const lx1 = g.left.x;
    const lx2 = g.left.x + strikeW;
    page.drawLine({ start: { x: lx1, y: midY + 1.5 }, end: { x: lx2, y: midY + 1.5 }, thickness: 0.8, color: rgb(0,0,0) });
    page.drawLine({ start: { x: lx1, y: midY - 1.5 }, end: { x: lx2, y: midY - 1.5 }, thickness: 0.8, color: rgb(0,0,0) });
    const senseiX = g.right.x + 2;
    page.drawText('先生', { x: senseiX, y: g.left.y, size: fs_, font, color: rgb(0, 0, 0) });
  }

  // 2. 受領日記入
  {
    const d = pos.date;
    const textW = font.widthOfTextAtSize(receiptDate, fs_);
    const whiteWidth = Math.max(textW + 40, pgW * 0.50);
    const margin = 3;
    const rectBottom = d.yBase - margin;
    const rectTop    = d.yTop  + margin;
    const rectHeight = rectTop - rectBottom;
    page.drawRectangle({ x: d.x - 4, y: rectBottom, width: whiteWidth, height: rectHeight, color: rgb(1, 1, 1) });
    page.drawText(receiptDate, { x: d.x, y: d.yBase, size: fs_, font, color: rgb(0, 0, 0) });
  }

  // 3. 署名記入
  {
    const a = pos.agent;
    let nameX;
    if (pos.agentTitleEndX) {
      nameX = pos.agentTitleEndX + 4;
    } else {
      const titleWidth = font.widthOfTextAtSize(signerTitle, fs_);
      nameX = a.x + titleWidth + 4;
    }
    const nameText = `　${signerName}`;
    const nameW = font.widthOfTextAtSize(nameText, fs_);

    const sigMargin = 3;
    const sigRectBottom = a.yBase - sigMargin;
    const sigRectTop    = a.yTop  + sigMargin;
    const sigRectHeight = sigRectTop - sigRectBottom;
    page.drawRectangle({ x: nameX - 2, y: sigRectBottom, width: nameW + 20, height: sigRectHeight, color: rgb(1, 1, 1) });
    page.drawText(nameText, { x: nameX, y: a.yBase, size: fs_, font, color: rgb(0, 0, 0) });

    // 4. 印鑑画像
    const sealBase64 = localStorage.getItem('tsukurukun_seal');
    if (sealBase64) {
      try {
        const sealData = Uint8Array.from(atob(sealBase64.replace(/^data:image\/\w+;base64,/, '')), c => c.charCodeAt(0));
        let sealImage;
        if (sealBase64.includes('image/png')) {
          sealImage = await pdfDoc.embedPng(sealData);
        } else {
          sealImage = await pdfDoc.embedJpg(sealData);
        }
        const sealSize = 36;
        const sealX = nameX + nameW + 2;
        const sealY = a.yBase - sealSize * 0.5 + fs_ * 0.3;
        page.drawImage(sealImage, { x: sealX, y: sealY, width: sealSize, height: sealSize });
      } catch (e) {
        console.warn('印鑑画像の読み込みに失敗:', e);
        page.drawText('㊞', { x: nameX + nameW + 4, y: a.yBase, size: fs_, font, color: rgb(0, 0, 0) });
      }
    } else {
      page.drawText('㊞', { x: nameX + nameW + 4, y: a.yBase, size: fs_, font, color: rgb(0, 0, 0) });
    }
  }

  // 出力（受領書ページのみ抽出）
  onProgress && onProgress('PDFを生成中...');
  const outDoc = await PDFLib.PDFDocument.create();
  outDoc.registerFontkit(fontkit);
  const [copiedPage] = await outDoc.copyPages(pdfDoc, [receiptPageIndex]);
  outDoc.addPage(copiedPage);
  const savedBytes = await outDoc.save();

  const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const baseName = file.name.replace(/\.pdf$/i, '');
  const outFileName = `受領書_${baseName}_${ts}.pdf`;

  const blob = new Blob([savedBytes], { type: 'application/pdf' });
  console.log(`✅ 受領書生成完了: ${outFileName}`);

  return { blob, fileName: outFileName };
}

// グローバルに公開
window.generateReceiptBrowser = generateReceiptBrowser;
