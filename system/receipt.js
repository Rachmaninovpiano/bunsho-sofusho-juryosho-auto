/**
 * 受領書自動でつくる君 - 受領書生成モジュール
 *
 * 相手方から送られてきた文書送付書PDFの受領書部分に
 * 受領日・記名・押印を自動で書き込み、PDFとして出力する。
 *
 * 対応する文書パターン:
 *   A) 1ページ目が「送付書（上）+受領書（下）」の複合ページ + 以降が書面本文
 *   B) 1ページ全体が受領書（独立型）
 *   C) 最終ページが受領書
 *
 * 処理内容:
 *   1. 「受領書」セクションを特定（OCRで「受領」「受　領　書」を検出）
 *   2. 「行」を見つけて二重打消し線を引き、横に「先生」を追加
 *   3. 「令和　年　月　日」に受領日を記入
 *   4. 「被告（ら）訴訟代理人」等の下に弁護士名を記入
 *   5. 弁護士名の横に印鑑画像を配置
 *   6. 受領書ページ（1ページ分）のみを出力
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { PDFDocument, rgb } = require('pdf-lib');
const fontkit = require('@pdf-lib/fontkit');
const { execSync } = require('child_process');

// ===== 設定ファイル読み込み =====
const BASE_DIR = path.resolve(__dirname, '..');
const CONFIG_PATH = path.join(BASE_DIR, 'config.json');
let CONFIG = { officeName: '', lawyerNames: [], faxNumbers: [], port: 3000 };
try {
  if (fs.existsSync(CONFIG_PATH)) {
    CONFIG = { ...CONFIG, ...JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8')) };
  }
} catch (e) { /* ignore */ }

// 印鑑画像フォルダ
const SEAL_DIR = path.join(BASE_DIR, 'seal');

// Tesseract OCRのパス候補
const TESSERACT_PATHS = [
  'C:/Program Files/Tesseract-OCR/tesseract.exe',
  'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe',
  'tesseract',
];

// Ghostscriptのパス候補
const GS_PATHS = [
  'C:/Program Files/gs/gs10.04.0/bin/gswin64c.exe',
  'C:/Program Files/gs/gs10.03.1/bin/gswin64c.exe',
  'C:/Program Files (x86)/gs/gs10.04.0/bin/gswin64c.exe',
  'C:/Program Files (x86)/gs/gs10.03.1/bin/gswin64c.exe',
];

// 日本語フォントのパス候補（TTFのみ、TTCは直接使えない）
const FONT_PATHS = [
  'C:/Windows/Fonts/yumin.ttf',         // 游明朝
  'C:/Windows/Fonts/yumindb.ttf',       // 游明朝 Demibold
  'C:/Windows/Fonts/YuMincho-Regular.ttf',
];

function findTesseract() {
  for (const tp of TESSERACT_PATHS) {
    try {
      if (tp.includes('/') || tp.includes('\\')) {
        if (fs.existsSync(tp)) return tp;
      } else {
        execSync(`"${tp}" --version`, { stdio: 'pipe' });
        return tp;
      }
    } catch (e) { /* next */ }
  }
  return null;
}

function findGhostscript() {
  for (const gp of GS_PATHS) {
    if (fs.existsSync(gp)) return `"${gp}"`;
  }
  return 'gswin64c';
}

function findFont() {
  for (const fp of FONT_PATHS) {
    if (fs.existsSync(fp)) return fp;
  }
  return null;
}

/**
 * 印鑑画像を探す（PNG/JPG）
 */
function findSealImage() {
  if (!fs.existsSync(SEAL_DIR)) return null;
  const files = fs.readdirSync(SEAL_DIR);
  const found = files.find(f => /\.(png|jpg|jpeg)$/i.test(f));
  return found ? path.join(SEAL_DIR, found) : null;
}

/**
 * PDFの指定ページをGhostscriptでPNGに変換してOCR実行
 * @param {string} pdfPath
 * @param {number} pageNum - 1始まり
 * @returns {{ words, imgWidth, imgHeight }}
 */
function runOcr(pdfPath, pageNum = 1) {
  const tesseract = findTesseract();
  if (!tesseract) throw new Error('Tesseract OCRがインストールされていません。');

  const tempDir = path.join(BASE_DIR, 'temp_receipt_ocr');
  fs.mkdirSync(tempDir, { recursive: true });

  try {
    // PDF → PNG（300dpi）
    const pngPath = path.join(tempDir, 'page.png');
    const gsCmd = findGhostscript();
    execSync(
      `${gsCmd} -dNOPAUSE -dBATCH -sDEVICE=png16m -r300 -dFirstPage=${pageNum} -dLastPage=${pageNum} -sOutputFile="${pngPath}" "${pdfPath}"`,
      { stdio: 'pipe', timeout: 60000 }
    );
    if (!fs.existsSync(pngPath)) throw new Error('PDF→PNG変換失敗');

    // PNGサイズ取得
    const pngBuf = fs.readFileSync(pngPath);
    const imgWidth = pngBuf.readUInt32BE(16);
    const imgHeight = pngBuf.readUInt32BE(20);

    // Tesseract HOCR（日本語）
    let lang = 'jpn';
    try {
      const langs = execSync(`"${tesseract}" --list-langs`, { stdio: 'pipe' }).toString();
      if (!langs.includes('jpn')) lang = 'eng';
    } catch (e) { lang = 'eng'; }

    const hocrBase = path.join(tempDir, 'ocr');
    execSync(`"${tesseract}" "${pngPath}" "${hocrBase}" -l ${lang} --psm 6 hocr`, {
      stdio: 'pipe', timeout: 120000
    });

    const hocrPath = hocrBase + '.hocr';
    if (!fs.existsSync(hocrPath)) throw new Error('HOCR出力が生成されませんでした');
    const hocr = fs.readFileSync(hocrPath, 'utf-8');

    // 単語を抽出
    const words = [];
    const re = /<span[^>]+class='ocrx_word'[^>]+title='bbox (\d+) (\d+) (\d+) (\d+)[^>]*>([\s\S]*?)<\/span>/g;
    let m;
    while ((m = re.exec(hocr)) !== null) {
      const text = m[5].replace(/<[^>]+>/g, '').trim();
      if (!text) continue;
      words.push({
        x1: +m[1], y1: +m[2], x2: +m[3], y2: +m[4], text
      });
    }

    console.log(`  [OCR] ページ${pageNum}: ${words.length}語検出`);
    return { words, imgWidth, imgHeight };

  } finally {
    try {
      fs.readdirSync(tempDir).forEach(f => fs.unlinkSync(path.join(tempDir, f)));
      fs.rmdirSync(tempDir);
    } catch (e) { /* ignore */ }
  }
}

/**
 * ピクセル座標 → PDF座標（左下原点）変換
 */
function px2pdf(px, py, imgW, imgH, pgW, pgH) {
  return {
    x: px * pgW / imgW,
    y: pgH - (py * pgH / imgH),
  };
}

// ===================================================================
// 受領書ページ特定（効率的: 1ページ目→最終ページ→残りの順に探す）
// ===================================================================

/**
 * OCR単語リストから「受領書」ラベルを探す
 * 「受領書」「受　領　書」「受領」などのパターンに対応
 * @returns {{ found: boolean, y: number|null }} - found=true なら y は「受領書」ラベルのピクセルY座標
 */
function findReceiptLabel(words) {
  // 1) 「受領書」「受領」がそのまま含まれる単語
  const direct = words.find(w =>
    w.text.includes('受領書') || w.text.includes('受領')
  );
  if (direct) return { found: true, y: direct.y1 };

  // 2) 「受」「領」「書」がバラバラに並ぶパターン（受　領　書）
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

/**
 * ページが受領書を含むかスコアリング
 * 「受領」ラベルの存在を最重要視する
 */
function scoreReceiptPage(words) {
  const allText = words.map(w => w.text).join('');
  let score = 0;

  // ★最重要: 「受領」ラベルの存在
  const receiptLabel = findReceiptLabel(words);
  if (receiptLabel.found) score += 50;

  // 受領書特有の要素
  if (/令和/.test(allText))  score += 10;
  if (/代理人/.test(allText)) score += 10;

  // ペナルティ: 長文の本文ページ
  if (words.length > 200) score -= 20;
  if (words.length > 300) score -= 20;

  return score;
}

/**
 * 受領書ページを効率的に特定する
 * 戦略: 1ページ目 → 最終ページ → 残りの順にスキャン（高スコアが見つかれば即返す）
 */
function findReceiptPage(pdfPath, totalPages) {
  console.log(`  🔍 受領書ページ特定中... (全${totalPages}ページ)`);

  if (totalPages === 1) {
    console.log('  📄 1ページのみ');
    const ocr = runOcr(pdfPath, 1);
    return { pageNum: 1, ocr };
  }

  // スキャン順序: 1ページ目（最も可能性が高い）→ 最終ページ → 残り
  const scanOrder = [1, totalPages];
  for (let p = 2; p < totalPages; p++) {
    scanOrder.push(p);
  }

  let bestPageNum = 1;
  let bestOcr = null;
  let bestScore = -Infinity;

  for (const p of scanOrder) {
    console.log(`    ページ ${p}/${totalPages} スキャン中...`);
    const ocr = runOcr(pdfPath, p);
    const score = scoreReceiptPage(ocr.words);
    console.log(`    → スコア: ${score} (単語数: ${ocr.words.length})`);

    if (score > bestScore) {
      bestScore = score;
      bestPageNum = p;
      bestOcr = ocr;
    }

    // 「受領」ラベルが見つかった（スコア50以上）→ 即決定
    if (score >= 50) {
      console.log(`  ✅ 受領書ページ確定: ページ${p} (スコア: ${score})`);
      return { pageNum: p, ocr };
    }
  }

  console.log(`  ✅ 受領書ページ推定: ページ${bestPageNum} (スコア: ${bestScore})`);
  return { pageNum: bestPageNum, ocr: bestOcr };
}

// ===================================================================
// 受領書セクション内の書き込み位置検出
// ===================================================================

/**
 * 受領書セクションの開始Y座標を特定
 *
 * パターンA: 送付書（上）+受領書（下）の複合ページ
 *   → 「受領書」ラベルのY座標を境界とする
 * パターンB: ページ全体が受領書
 *   → Y=0 からスタート
 */
function findReceiptSectionStart(words, imgH) {
  const receiptLabel = findReceiptLabel(words);

  if (receiptLabel.found && receiptLabel.y !== null) {
    // 「受領書」ラベルが見つかった → その少し上を境界にする
    const borderY = Math.max(0, receiptLabel.y - 50);
    console.log(`  📊 「受領書」ラベル検出 y=${receiptLabel.y}px → 受領書セクション開始: ${borderY}px`);
    return borderY;
  }

  // 「受領書」ラベルがない場合はページ全体を対象
  console.log('  📊 「受領書」ラベル未検出 → ページ全体を受領書として処理');
  return 0;
}

/**
 * OCR結果から受領書セクション内の書き込み位置を検出
 */
function detectPositions(words, imgW, imgH, pgW, pgH) {
  // 受領書セクションの開始Y座標
  const receiptStartY = findReceiptSectionStart(words, imgH);

  // 受領書セクションの単語のみ対象
  const rw = words.filter(w => w.y1 >= receiptStartY);
  console.log(`  📊 受領書エリア単語数: ${rw.length}`);

  // DEBUG（必要時のみ有効化）:
  // rw.forEach(w => { console.log(`    [rw] "${w.text}" px(${w.x1},${w.y1})-(${w.x2},${w.y2})`); });

  // =========================================
  // 1. 「行」の検出
  //
  // ★重要: OCRテキストが明確に「行」と読めた場合のみ処理する
  //   推定・フォールバックは行わない（「殿」「宛」等の書式で誤検出するため）
  //   「行」がない書式ではこのステップはスキップされる
  // =========================================
  let gyouWord = null;

  // 受領書セクション内でOCRテキストが「行」と完全一致する単語を検索
  const allGyouWords = rw.filter(w => w.text === '行');

  if (allGyouWords.length > 0) {
    // 「弁護士」と同じ行にある「行」を優先
    const bengoWordInReceipt = rw.find(w =>
      w.text.includes('弁護') || w.text.includes('護士')
    );

    if (bengoWordInReceipt) {
      const bengoY = bengoWordInReceipt.y1;
      console.log(`  📍 受領書内「弁護士」: px(${bengoWordInReceipt.x1},${bengoY}) text="${bengoWordInReceipt.text}"`);

      // 弁護士行（±60px）にある「行」
      const gyouInBengoLine = allGyouWords.filter(w =>
        Math.abs(w.y1 - bengoY) < 60
      );
      if (gyouInBengoLine.length > 0) {
        // 最右端の「行」（弁護士名の末尾）
        gyouWord = gyouInBengoLine.reduce((a, b) => a.x1 > b.x1 ? a : b);
        console.log(`  📍 「行」検出(弁護士行): px(${gyouWord.x1},${gyouWord.y1})`);
      }
    }

    // 弁護士行で見つからなかった場合: 受領書セクション最上部の「行」
    if (!gyouWord) {
      gyouWord = allGyouWords.reduce((a, b) => a.y1 < b.y1 ? a : b);
      console.log(`  📍 「行」検出(受領書内): px(${gyouWord.x1},${gyouWord.y1})`);
    }
  } else {
    console.log('  ℹ️ 「行」なし（「殿」「宛」等の書式）→ スキップ');
  }

  // =========================================
  // 2. 署名行の検出（先に検出 → 日付の検索範囲を限定するため）
  //
  // 「被告訴訟代理人」「被告ら訴訟代理人」「原告訴訟代理人」「相手方代理人」
  // 受領書セクション最下部にある
  // =========================================
  let agentWord = null;
  // 受領書セクション内で「代理人」を含む単語を検索
  const agentCandidates = rw.filter(w =>
    w.text.includes('代理人') || w.text.includes('代理')
  );

  if (agentCandidates.length > 0) {
    // 最も下にある「代理人」（受領書の署名行は最下部）
    agentWord = agentCandidates.reduce((a, b) => a.y1 > b.y1 ? a : b);
    console.log(`  📍 「代理人」検出: px(${agentWord.x1},${agentWord.y1}) text="${agentWord.text}"`);
  } else {
    // 「被告」「原告」でも試す（受領書セクション下半分）
    const midY = receiptStartY + (imgH - receiptStartY) * 0.5;
    const agentCandidates2 = rw.filter(w =>
      w.y1 > midY &&
      (w.text.includes('被告') || w.text.includes('原告'))
    );
    if (agentCandidates2.length > 0) {
      agentWord = agentCandidates2.reduce((a, b) => a.y1 > b.y1 ? a : b);
      console.log(`  📍 「被告/原告」検出: px(${agentWord.x1},${agentWord.y1}) text="${agentWord.text}"`);
    }
  }

  // 代理人行の全単語を収集して「人」の右端を見つける
  if (agentWord) {
    const agentY = agentWord.y1;
    const agentRowWords = rw.filter(w => Math.abs(w.y1 - agentY) < 60)
      .sort((a, b) => a.x1 - b.x1);
    const ninWord = agentRowWords.find(w => w.text === '人' || w.text.endsWith('人'));
    if (ninWord) {
      agentWord._titleEndX = ninWord.x2;
      console.log(`  📍 「人」右端: px(${ninWord.x2})`);
    }
    const leftMost = agentRowWords[0];
    if (leftMost && leftMost.x1 < agentWord.x1) {
      agentWord._lineStartX = leftMost.x1;
    }
  }

  // =========================================
  // 3. 受領書内の日付欄の検出
  //
  // 書式パターン:
  //   A) 「令和　年　月　日」（空欄タイプ）→ 「令和」の位置に書き込む
  //   B) 日付記入欄がない → 「受領書」ラベルと「署名行」の間で「年」「月」を含む行
  //
  // 検索範囲: 「行」位置（or受領書ラベル）〜 署名行の間
  // =========================================
  const searchTopY = gyouWord ? gyouWord.y1 + 20 : receiptStartY;
  const searchBottomY = agentWord ? agentWord.y1 - 10 : imgH;

  let dateWord = null;
  // 「令」or「令和」を検索範囲内で探す
  const reiwaWords = rw.filter(w =>
    w.y1 > searchTopY && w.y1 < searchBottomY &&
    (w.text === '令' || w.text === '令和' || w.text.startsWith('令'))
  );

  if (reiwaWords.length > 0) {
    dateWord = reiwaWords.reduce((a, b) => a.y1 < b.y1 ? a : b);
    console.log(`  📍 「令和」検出: px(${dateWord.x1},${dateWord.y1}) text="${dateWord.text}"`);
  } else {
    // 「年」「月」を含む行で代用
    const dateish = rw.filter(w =>
      w.y1 > searchTopY && w.y1 < searchBottomY &&
      (w.text.includes('年') || w.text.includes('月'))
    );
    if (dateish.length > 0) {
      dateWord = dateish.reduce((a, b) => a.y1 < b.y1 ? a : b);
      console.log(`  📍 日付行検出: px(${dateWord.x1},${dateWord.y1}) text="${dateWord.text}"`);
    }
  }

  // =========================================
  // フォールバック推定
  // 検出できなかった要素はagentWordを基準に推定
  // =========================================
  if (!dateWord && agentWord) {
    console.log('  ⚠️ 日付未検出 → 署名行基準で推定');
    // 署名行の少し上（署名行と「行」/ラベルの中間あたり）
    const estimatedY = Math.round(agentWord.y1 - (imgH * 0.06));
    const estimatedX = Math.round(imgW * 0.05);
    dateWord = {
      x1: estimatedX, y1: estimatedY,
      x2: estimatedX + 200, y2: estimatedY + 40,
      text: '令和（推定）', estimated: true
    };
  } else if (!dateWord) {
    console.log('  ⚠️ 日付未検出 → 受領書エリア中間で推定');
    const estimatedY = Math.round(receiptStartY + (imgH - receiptStartY) * 0.5);
    const estimatedX = Math.round(imgW * 0.05);
    dateWord = {
      x1: estimatedX, y1: estimatedY,
      x2: estimatedX + 200, y2: estimatedY + 40,
      text: '令和（推定）', estimated: true
    };
  }

  if (!agentWord) {
    console.log('  ⚠️ 「代理人」未検出 → 推定');
    // 受領書セクション最下部
    const estimatedY = Math.round(receiptStartY + (imgH - receiptStartY) * 0.85);
    const estimatedX = Math.round(imgW * 0.20);
    agentWord = {
      x1: estimatedX, y1: estimatedY,
      x2: estimatedX + 300, y2: estimatedY + 40,
      text: '代理人（推定）', estimated: true
    };
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
    ? agentWord._titleEndX * pgW / imgW
    : null;

  return {
    gyou: gyouPdf,
    date: {
      x:      datePdfBottom.x,
      yTop:   datePdfTop.y,
      yBase:  datePdfBottom.y,
    },
    agent: {
      x:      agentLinePdf.x,
      yTop:   agentPdfTop.y,
      yBase:  agentPdfBottom.y,
    },
    agentTitleEndX,
    agentRowY: agentPdfBottom.y,
    imgW, imgH,
  };
}

// ===================================================================
// メイン: 受領書PDF生成
// ===================================================================

/**
 * 受領書PDFを生成
 */
async function generateReceipt(pdfPath, options = {}) {
  console.log('\n📄 受領書生成開始:', path.basename(pdfPath));

  // デフォルト値
  const today = new Date();
  const reiwaYear = today.getFullYear() - 2018;
  const defaultDate = `令和${reiwaYear}年${today.getMonth() + 1}月${today.getDate()}日`;

  const receiptDate  = options.receiptDate  || defaultDate;
  const signerTitle  = options.signerTitle  || '被告訴訟代理人';
  const defaultName  = (CONFIG.lawyerNames && CONFIG.lawyerNames.length > 0)
    ? CONFIG.lawyerNames[CONFIG.lawyerNames.length - 1]
    : '山田太郎';
  const signerName   = options.signerName   || defaultName;
  const sealImgPath  = options.sealImagePath || findSealImage();
  const outputDir    = options.outputDir    || path.join(BASE_DIR, 'output');

  console.log(`  📅 受領日: ${receiptDate}`);
  console.log(`  ✍️  署名: ${signerTitle}　${signerName}`);
  console.log(`  🔏 印鑑: ${sealImgPath || '(なし → ㊞で代替)'}`);

  // ===== PDF読み込み =====
  const pdfBytes = fs.readFileSync(pdfPath);
  const pdfDoc = await PDFDocument.load(pdfBytes);
  pdfDoc.registerFontkit(fontkit);

  const totalPages = pdfDoc.getPageCount();
  console.log(`  📑 総ページ数: ${totalPages}`);

  // ===== 受領書ページ特定（効率的スキャン）=====
  const { pageNum: receiptPageNum, ocr } = findReceiptPage(pdfPath, totalPages);
  const receiptPageIndex = receiptPageNum - 1; // 0始まり

  const words    = ocr.words;
  const imgWidth  = ocr.imgWidth;
  const imgHeight = ocr.imgHeight;

  const page = pdfDoc.getPage(receiptPageIndex);
  const { width: pgW, height: pgH } = page.getSize();
  console.log(`  📐 受領書ページ(${receiptPageNum}): ${pgW.toFixed(1)} × ${pgH.toFixed(1)} pt`);

  // ===== フォント読み込み =====
  const fontPath = findFont();
  if (!fontPath) {
    throw new Error('游明朝フォント(yumin.ttf)が見つかりません。C:/Windows/Fonts/yumin.ttf を確認してください。');
  }
  console.log(`  🔤 フォント: ${path.basename(fontPath)}`);

  const fontBytes = fs.readFileSync(fontPath);
  const font = await pdfDoc.embedFont(fontBytes, { subset: false });

  // グリフ事前登録
  const allChars = `行先生${receiptDate}${signerTitle}　${signerName}㊞`;
  try { font.encodeText(allChars); } catch (e) { /* ignore */ }

  // ===== 位置検出 =====
  const pos = detectPositions(words, imgWidth, imgHeight, pgW, pgH);

  // フォントサイズ（元のFAX文書に合わせる: 約10.5pt）
  const fs_ = 10.5;

  // =========================================
  // 1. 「行」→ 二重打消し線 + 「先生」
  // =========================================
  if (pos.gyou) {
    const g = pos.gyou;
    console.log(`  ✏️ 「行」処理: PDF(${g.left.x.toFixed(1)}, ${g.left.y.toFixed(1)})`);

    // 元の「行」文字はそのまま残す（白矩形で消さない＝隣の文字を潰さない）
    // 上から二重打消し線を引く + 右に「先生」を追記するだけ

    // 打消し線の幅はOCR bboxの幅を使う（元の「行」文字の実際の幅）
    const gyouOcrW = g.width; // OCR bboxの幅（ピクセル→PDF変換済み）
    const gyouCharW = font.widthOfTextAtSize('行', fs_);
    // 打消し線の幅: OCR幅とフォント幅の小さい方を使い、はみ出しを防ぐ
    const strikeW = Math.min(gyouOcrW, gyouCharW);

    // 二重打消し線（「行」文字の上に2本の横線）
    const midY = g.left.y + fs_ * 0.40;
    const lx1 = g.left.x;
    const lx2 = g.left.x + strikeW;
    page.drawLine({ start: { x: lx1, y: midY + 1.5 }, end: { x: lx2, y: midY + 1.5 }, thickness: 0.8, color: rgb(0,0,0) });
    page.drawLine({ start: { x: lx1, y: midY - 1.5 }, end: { x: lx2, y: midY - 1.5 }, thickness: 0.8, color: rgb(0,0,0) });

    // 「先生」を「行」の右横に追記
    // 追記位置: OCR bbox右端から少し離す
    const senseiX = g.right.x + 2;
    page.drawText('先生', {
      x: senseiX,
      y: g.left.y,
      size: fs_,
      font,
      color: rgb(0, 0, 0),
    });

    console.log('    ✅ 「行」→ 二重打消し線 + 「先生」完了');
  } else {
    console.log('  ⚠️ 「行」が見つかりませんでした');
  }

  // =========================================
  // 2. 受領日記入
  // =========================================
  {
    const d = pos.date;
    console.log(`  ✏️ 受領日記入: PDF(${d.x.toFixed(1)}, yBase=${d.yBase.toFixed(1)}, yTop=${d.yTop.toFixed(1)})`);

    const textW = font.widthOfTextAtSize(receiptDate, fs_);
    const whiteWidth = Math.max(textW + 40, pgW * 0.50);
    const margin = 3;

    const rectBottom = d.yBase - margin;
    const rectTop    = d.yTop  + margin;
    const rectHeight = rectTop - rectBottom;

    page.drawRectangle({
      x: d.x - 4,
      y: rectBottom,
      width:  whiteWidth,
      height: rectHeight,
      color: rgb(1, 1, 1),
    });

    page.drawText(receiptDate, {
      x: d.x,
      y: d.yBase,
      size: fs_,
      font,
      color: rgb(0, 0, 0),
    });

    console.log(`    ✅ 受領日「${receiptDate}」記入完了`);
  }

  // =========================================
  // 3. 署名記入
  // =========================================
  {
    const a = pos.agent;
    console.log(`  ✏️ 署名記入: PDF(${a.x.toFixed(1)}, yBase=${a.yBase.toFixed(1)})`);

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

    page.drawRectangle({
      x: nameX - 2,
      y: sigRectBottom,
      width:  nameW + 20,
      height: sigRectHeight,
      color: rgb(1, 1, 1),
    });

    page.drawText(nameText, {
      x: nameX,
      y: a.yBase,
      size: fs_,
      font,
      color: rgb(0, 0, 0),
    });

    console.log(`    ✅ 名前「${nameText}」記入完了`);

    // =========================================
    // 4. 印鑑画像配置
    // =========================================
    if (sealImgPath && fs.existsSync(sealImgPath)) {
      console.log('  ✏️ 印鑑画像配置...');
      const sealBytes = fs.readFileSync(sealImgPath);
      let sealImage;
      if (/\.png$/i.test(sealImgPath)) {
        sealImage = await pdfDoc.embedPng(sealBytes);
      } else {
        sealImage = await pdfDoc.embedJpg(sealBytes);
      }

      const sealSize = 36;
      const sealX = nameX + nameW + 2;
      const sealY = a.yBase - sealSize * 0.5 + fs_ * 0.3;

      page.drawImage(sealImage, {
        x: sealX,
        y: sealY,
        width:  sealSize,
        height: sealSize,
      });

      console.log(`    ✅ 印鑑配置完了 (${sealSize}pt角)`);
    } else {
      page.drawText('㊞', {
        x: nameX + nameW + 4,
        y: a.yBase,
        size: fs_,
        font,
        color: rgb(0, 0, 0),
      });
      console.log('    ℹ️ 印鑑画像なし → 「㊞」で代替');
    }
  }

  // ===== 出力（受領書ページのみ抽出）=====
  fs.mkdirSync(outputDir, { recursive: true });
  const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const baseName = path.basename(pdfPath, path.extname(pdfPath));
  const outFileName = `受領書_${baseName}_${ts}.pdf`;
  const outPath = path.join(outputDir, outFileName);

  // ★ 常に受領書ページ（1ページ）のみを出力
  const outDoc = await PDFDocument.create();
  outDoc.registerFontkit(fontkit);
  const [copiedPage] = await outDoc.copyPages(pdfDoc, [receiptPageIndex]);
  outDoc.addPage(copiedPage);
  const saved = await outDoc.save();
  fs.writeFileSync(outPath, saved);

  if (totalPages > 1) {
    console.log(`  ✂️ ${totalPages}ページ中のページ${receiptPageNum}のみを抽出`);
  }

  console.log(`\n✅ 受領書生成完了`);
  console.log(`📁 出力: ${outPath}`);

  return { outputPath: outPath, outputFileName: outFileName };
}

// ===== エクスポート =====
module.exports = { generateReceipt, findSealImage };

// ===== CLI実行 =====
if (require.main === module) {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.log('使い方: node system/receipt.js <PDFファイルパス> [受領日] [肩書] [弁護士名]');
    console.log('例: node system/receipt.js FAX転送_sample.pdf "令和8年2月19日" "被告訴訟代理人" "山田太郎"');
    process.exit(0);
  }

  const opts = {};
  if (args[1]) opts.receiptDate  = args[1];
  if (args[2]) opts.signerTitle  = args[2];
  if (args[3]) opts.signerName   = args[3];

  generateReceipt(args[0], opts).catch(err => {
    console.error('❌ エラー:', err.message);
    process.exit(1);
  });
}
