/**
 * 文書送付書 自動生成システム
 *
 * 使い方:
 *   node system/generate.js <PDFファイルパス> [送付書類名]
 *
 * 例:
 *   node system/generate.js 準備書面.pdf "被告第10準備書面"
 *   node system/generate.js 準備書面.pdf
 *
 * PDFから以下の情報を自動抽出:
 *   - 裁判所名（係属部を含む）
 *   - 裁判所FAX番号
 *   - 事件番号
 *   - 事件名
 *   - 原告名・被告名
 *   - 原告代理人弁護士名
 *   - 原告代理人FAX番号
 *
 * 出力: デスクトップ/書式/output/ に Word ファイルを生成
 */

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const PDFParser = require('pdf2json');
const { execSync } = require('child_process');

// ===== 設定ファイル読み込み =====
const CONFIG_PATH = path.join(path.resolve(__dirname, '..'), 'config.json');
let CONFIG = { officeName: '', lawyerNames: [], faxNumbers: [], port: 3000 };
try {
  if (fs.existsSync(CONFIG_PATH)) {
    CONFIG = { ...CONFIG, ...JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8')) };
  }
} catch (e) { /* config読み込み失敗時はデフォルト値を使用 */ }

// ===== ログ出力（デバッグ用）=====
const LOG_PATH = path.join(__dirname, 'last_run.log');
const logLines = [];
const origLog = console.log;
const origErr = console.error;
console.log = function (...args) {
  const msg = args.map(a => typeof a === 'string' ? a : JSON.stringify(a, null, 2)).join(' ');
  logLines.push(msg);
  origLog.apply(console, args);
};
console.error = function (...args) {
  const msg = args.map(a => typeof a === 'string' ? a : (a && a.stack ? a.stack : JSON.stringify(a))).join(' ');
  logLines.push('[ERROR] ' + msg);
  origErr.apply(console, args);
};
process.on('exit', () => {
  try { fs.writeFileSync(LOG_PATH, logLines.join('\n'), 'utf-8'); } catch (e) { /* ignore */ }
});

// ===== 設定 =====
const BASE_DIR = path.resolve(__dirname, '..');
const TEMPLATE_PATH = path.join(BASE_DIR, 'template', '文書送付書.doc.docx');
const OUTPUT_DIR = path.join(BASE_DIR, 'output');

// Tesseract OCRのパス候補
const TESSERACT_PATHS = [
  'C:/Program Files/Tesseract-OCR/tesseract.exe',
  'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe',
  'tesseract', // PATHに通っている場合
];

// ===== 裁判所FAX番号辞書 =====
const COURT_FAX_MAP = {
  '神戸地方裁判所尼崎支部': '06-6438-1710',
  '大阪地方裁判所': '06-6316-2804',
  '大阪高等裁判所': '06-6316-2804',
  '東京地方裁判所': '03-3580-5611',
  '東京高等裁判所': '03-3580-5611',
  '広島地方裁判所': '082-228-0197',
  '広島高等裁判所': '082-228-0197',
  '広島地方裁判所福山支部': '084-923-2897',
  '岡山地方裁判所': '086-222-6961',
  '福岡地方裁判所': '092-781-3141',
  '名古屋地方裁判所': '052-204-7780',
  '京都地方裁判所': '075-211-4226',
  '神戸地方裁判所': '078-367-1478',
  '横浜地方裁判所': '045-212-0947',
  'さいたま地方裁判所': '048-863-8761',
  '千葉地方裁判所': '043-227-5601',
  '仙台地方裁判所': '022-266-0091',
  '札幌地方裁判所': '011-271-1456',
  '山口地方裁判所': '083-922-1440',
};

// ===== Tesseract OCR実行パスを取得 =====
function findTesseract() {
  for (const tp of TESSERACT_PATHS) {
    try {
      if (tp.includes('/') || tp.includes('\\')) {
        if (fs.existsSync(tp)) return tp;
      } else {
        execSync(`${tp} --version`, { stdio: 'pipe' });
        return tp;
      }
    } catch (e) { /* next */ }
  }
  return null;
}

// ===== OCRでPDFからテキスト抽出 =====
function extractTextWithOCR(pdfPath) {
  const tesseract = findTesseract();
  if (!tesseract) {
    throw new Error(
      'Tesseract OCRがインストールされていません。\n' +
      '画像PDFを処理するにはTesseract OCRが必要です。\n' +
      'https://github.com/UB-Mannheim/tesseract/wiki からインストールしてください。'
    );
  }

  console.log('  📷 画像PDFを検出。OCRで文字認識を実行中...');

  // pdf2imageの代わりにGhostscriptまたはpdfimagesを使う
  // まずはTesseractのPDF直接読み取りを試行
  const tempDir = path.join(BASE_DIR, 'temp_ocr');
  if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });

  const tempOutput = path.join(tempDir, 'ocr_result');

  try {
    // Tesseractは画像ファイルを受け付けるので、まずPDFを画像に変換する必要がある
    // Windowsの場合、magick (ImageMagick) を試す
    let imagePath;

    // 方法1: ImageMagickがある場合
    try {
      const pngPath = path.join(tempDir, 'page.png');
      execSync(`magick -density 300 "${pdfPath}[0]" -quality 100 "${pngPath}"`, {
        stdio: 'pipe', timeout: 60000
      });
      imagePath = pngPath;
    } catch (e) {
      // 方法2: Ghostscriptがある場合
      try {
        const pngPath = path.join(tempDir, 'page.png');
        // gswin64cを試す
        // Ghostscriptのパスを探す
        let gsCmd = 'gswin64c';
        const gsPaths = [
          'C:/Program Files/gs/gs10.04.0/bin/gswin64c.exe',
          'C:/Program Files/gs/gs10.03.1/bin/gswin64c.exe',
          'C:/Program Files (x86)/gs/gs10.04.0/bin/gswin64c.exe',
        ];
        for (const gp of gsPaths) {
          if (fs.existsSync(gp)) { gsCmd = `"${gp}"`; break; }
        }
        execSync(`${gsCmd} -dNOPAUSE -dBATCH -sDEVICE=png16m -r300 -dFirstPage=1 -dLastPage=1 -sOutputFile="${pngPath}" "${pdfPath}"`, {
          stdio: 'pipe', timeout: 60000
        });
        imagePath = pngPath;
      } catch (e2) {
        // 方法3: pdftoppmがある場合
        try {
          execSync(`pdftoppm -r 300 -png -f 1 -l 1 "${pdfPath}" "${path.join(tempDir, 'page')}"`, {
            stdio: 'pipe', timeout: 60000
          });
          // pdftoppmはpage-1.pngのようなファイルを生成
          const pngFiles = fs.readdirSync(tempDir).filter(f => f.endsWith('.png'));
          if (pngFiles.length > 0) {
            imagePath = path.join(tempDir, pngFiles[0]);
          }
        } catch (e3) {
          throw new Error(
            'PDFを画像に変換するツールが見つかりません。\n' +
            '以下のいずれかをインストールしてください:\n' +
            '  - ImageMagick (https://imagemagick.org/)\n' +
            '  - Ghostscript (https://ghostscript.com/)\n' +
            '  - poppler-utils (pdftoppm)'
          );
        }
      }
    }

    if (!imagePath || !fs.existsSync(imagePath)) {
      throw new Error('PDFから画像への変換に失敗しました。');
    }

    // Tesseract OCR実行（日本語 + 英語）
    console.log('  🔤 OCR実行中...');
    let lang = 'jpn+eng';
    try {
      const langList = execSync(`"${tesseract}" --list-langs`, { stdio: 'pipe' }).toString();
      if (!langList.includes('jpn')) lang = 'eng';
    } catch (e) {
      lang = 'eng';
    }

    // ★ 精度改善: PSM 4（単一列）とPSM 6（単一ブロック）の両方でOCRし、
    //   結果をマージする。FAX送信書はレイアウトが複雑で、
    //   PSMモードによって認識できる部分が異なるため。
    let textPsm4 = '';
    let textPsm6 = '';

    // PSM 4（単一列テキスト → 名前など日本語の固有名詞に強い）
    try {
      execSync(`"${tesseract}" "${imagePath}" "${tempOutput}_psm4" -l ${lang} --psm 4`, {
        stdio: 'pipe', timeout: 120000
      });
      if (fs.existsSync(tempOutput + '_psm4.txt')) {
        textPsm4 = fs.readFileSync(tempOutput + '_psm4.txt', 'utf-8');
      }
    } catch (e) { /* ignore */ }

    // PSM 6（単一ブロック → 事件番号やレイアウト構造に強い）
    try {
      execSync(`"${tesseract}" "${imagePath}" "${tempOutput}_psm6" -l ${lang} --psm 6`, {
        stdio: 'pipe', timeout: 120000
      });
      if (fs.existsSync(tempOutput + '_psm6.txt')) {
        textPsm6 = fs.readFileSync(tempOutput + '_psm6.txt', 'utf-8');
      }
    } catch (e) { /* ignore */ }

    // 両方のテキストを区切り付きで結合（extractInfoFromTextが両方から抽出可能に）
    let allText = textPsm4;
    if (textPsm6.length > 0) {
      allText += '\n\n===OCR_PSM6===\n\n' + textPsm6;
    }

    // 複数ページ対応（2ページ目以降）
    // FAX送信書（1ページ目）に必要情報が集約されているため、追加ページは3ページ目まで
    try {
      for (let pageNum = 2; pageNum <= 3; pageNum++) {
        const pagePng = path.join(tempDir, `page_${pageNum}.png`);
        let pageConverted = false;

        // ImageMagick を試す
        try {
          execSync(`magick -density 300 "${pdfPath}[${pageNum - 1}]" -quality 100 "${pagePng}"`, {
            stdio: 'pipe', timeout: 60000
          });
          pageConverted = true;
        } catch (e) {
          // Ghostscript を試す
          try {
            let gsCmd = 'gswin64c';
            const gsPaths = [
              'C:/Program Files/gs/gs10.04.0/bin/gswin64c.exe',
              'C:/Program Files/gs/gs10.03.1/bin/gswin64c.exe',
              'C:/Program Files (x86)/gs/gs10.04.0/bin/gswin64c.exe',
            ];
            for (const gp of gsPaths) {
              if (fs.existsSync(gp)) { gsCmd = `"${gp}"`; break; }
            }
            execSync(`${gsCmd} -dNOPAUSE -dBATCH -sDEVICE=png16m -r300 -dFirstPage=${pageNum} -dLastPage=${pageNum} -sOutputFile="${pagePng}" "${pdfPath}"`, {
              stdio: 'pipe', timeout: 60000
            });
            pageConverted = true;
          } catch (e2) {
            break; // これ以上変換できない
          }
        }

        if (pageConverted && fs.existsSync(pagePng)) {
          const pageOutput = path.join(tempDir, `ocr_page_${pageNum}`);
          execSync(`"${tesseract}" "${pagePng}" "${pageOutput}" -l ${lang} --psm 6`, {
            stdio: 'pipe', timeout: 120000
          });
          if (fs.existsSync(pageOutput + '.txt')) {
            allText += '\n' + fs.readFileSync(pageOutput + '.txt', 'utf-8');
          }
        }
      }
    } catch (e) {
      // 複数ページ処理のエラーは無視（1ページ目のテキストで続行）
    }

    return allText;
  } finally {
    // 一時ファイル削除
    try {
      const tempFiles = fs.readdirSync(tempDir);
      for (const f of tempFiles) {
        fs.unlinkSync(path.join(tempDir, f));
      }
      fs.rmdirSync(tempDir);
    } catch (e) { /* cleanup error ignored */ }
  }
}

// ===== PDF読み取り（テキスト抽出 → 空ならOCR）=====
async function extractTextFromPDF(pdfPath) {
  // まずテキスト抽出を試みる
  const text = await new Promise((resolve, reject) => {
    const pdfParser = new PDFParser();
    pdfParser.on('pdfParser_dataError', errData => reject(new Error(errData.parserError)));
    pdfParser.on('pdfParser_dataReady', pdfData => {
      let text = '';
      for (const page of pdfData.Pages) {
        for (const textItem of page.Texts) {
          for (const run of textItem.R) {
            text += decodeURIComponent(run.T);
          }
          text += '\n';
        }
        text += '\n';
      }
      resolve(text);
    });
    pdfParser.loadPDF(pdfPath);
  });

  // テキストが実質空（空白・改行のみ）ならOCR
  const trimmed = text.replace(/[\s\n\r]/g, '');
  if (trimmed.length < 10) {
    console.log('  ⚠️ テキスト埋め込みなしPDF → OCRに切り替え');
    return extractTextWithOCR(pdfPath);
  }

  return text;
}

// ===== 情報抽出 =====
function extractInfoFromText(text) {
  const info = {};

  const normalizedText = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');

  // 都市名リスト（裁判所パターン用）
  const cityNames = '東京|大阪|名古屋|広島|福岡|仙台|札幌|高松|京都|神戸|横浜|さいたま|千葉|山口|岡山|福山|松山|高知|那覇|長崎|熊本|鹿児島|大分|宮崎|佐賀|秋田|青森|盛岡|山形|福島|水戸|宇都宮|前橋|甲府|長野|新潟|富山|金沢|福井|津|大津|奈良|和歌山|鳥取|松江|徳島|旭川|釧路|函館';

  // --- 裁判所名（全マッチから最も完全なものを選ぶ） ---
  const courtPattern = new RegExp(
    `((?:${cityNames})\\s*(?:地方|高等|家庭|簡易)\\s*裁判\\s*所(?:\\s*[\\u4e00-\\u9fff]+\\s*支部)?(?:\\s*民事\\s*第\\s*[０-９\\d]+\\s*部)?)`,
    'g'
  );
  let courtMatch;
  const courtCandidates = [];
  while ((courtMatch = courtPattern.exec(normalizedText)) !== null) {
    const cleaned = courtMatch[1].replace(/\s+/g, '');
    courtCandidates.push(cleaned);
  }
  if (courtCandidates.length > 0) {
    // 最も長い（＝最も完全な）マッチを採用
    info.courtName = courtCandidates.reduce((a, b) => a.length >= b.length ? a : b);
  }

  // --- 事件番号 ---
  // パターン1: 正常なOCR結果（スペース挿入に対応）
  // OCRでは「令 和 6 年 ( ワ ) 第 228 号」のように大量のスペースが入る
  const caseSymbols = 'ワヲネレモハノニナラ行わをねれもはのになら';
  const caseNumberPatterns = [
    // 厳密パターン
    new RegExp(`([令平]和\\d+年[（(][${caseSymbols}][）)]\\s*第?\\s*\\d+号)`),
    // OCRスペース挿入対応（全角/半角括弧混在）
    new RegExp(`([令平]\\s*和\\s*\\d+\\s*年\\s*[（(]\\s*[${caseSymbols}]\\s*[）)]\\s*第?\\s*\\d+\\s*号)`),
    // 「(ワ)」が半角の場合: (ワ) → ( ワ )
    new RegExp(`([令平]\\s*和\\s*\\d+\\s*年\\s*\\(\\s*[${caseSymbols}]\\s*\\)\\s*第?\\s*\\d+\\s*号)`),
    // 極端なスペース挿入: 「令 和 6 年 ( ワ ) 第 225 号」
    new RegExp(`(令\\s*和\\s*(\\d+)\\s*年\\s*[（(]\\s*([${caseSymbols}])\\s*[）)]\\s*第\\s*(\\d+)\\s*号)`),
  ];
  for (const pattern of caseNumberPatterns) {
    const match = normalizedText.match(pattern);
    if (match) {
      // スペースを除去して正規化
      let cn = match[1].replace(/\s+/g, '');
      // 全角括弧を半角に統一
      cn = cn.replace(/（/g, '(').replace(/）/g, ')');
      info.caseNumber = cn;
      break;
    }
  }

  // パターン2: 「事件の表示」セクションから号数+符号を抽出
  if (!info.caseNumber) {
    const displaySectionMatch = normalizedText.match(
      /事\s*件\s*の\s*表\s*示[】\]\s]*([^\n]{1,80})/
    );
    if (displaySectionMatch) {
      const sectionText = displaySectionMatch[1];
      // まず完全な事件番号を探す（スペース許容）
      const fullMatch = sectionText.match(
        new RegExp(`令?\\s*和?\\s*(\\d+)\\s*年?\\s*[（(]\\s*([${caseSymbols}])\\s*[）)]\\s*第\\s*(\\d+)\\s*号`)
      );
      if (fullMatch) {
        const year = fullMatch[1];
        const symbol = fullMatch[2];
        const num = fullMatch[3];
        info.caseNumber = `令和${year}年(${symbol})第${num}号`;
      } else {
        // 「第NNN号」だけ抽出
        const numMatch = sectionText.match(/第\s*(\d+)\s*号/);
        if (numMatch) {
          const caseNum = numMatch[1];
          // 符号も探す
          const symbolMatch = sectionText.match(new RegExp(`[（(]\\s*([${caseSymbols}])\\s*[）)]`));
          const symbol = symbolMatch ? symbolMatch[1] : 'ワ';
          const guessed = !symbolMatch;

          // 年を推定
          const yearMatches = [];
          const yearRegex = /令\s*和\s*(\d+)\s*年/g;
          let ym;
          while ((ym = yearRegex.exec(normalizedText)) !== null) {
            yearMatches.push(parseInt(ym[1]));
          }
          if (yearMatches.length > 0) {
            const minYear = Math.min(...yearMatches);
            info.caseNumber = `令和${minYear}年(${symbol})第${caseNum}号`;
            info.caseNumberGuessed = guessed;
          } else {
            info.caseNumber = `(${symbol})第${caseNum}号`;
            info.caseNumberGuessed = true;
          }
        }
      }
    }
  }

  // --- 事件名 ---
  const caseNamePatterns = [
    // OCR対応: スペースで分断された「損害 賠償」+改行+「請求 事件」を結合
    /損害\s*賠償[\s\S]{0,50}?請求\s*事件/,
    /(?:号\s*)([\u4e00-\u9fff]+(?:請求|確認|等?)\s*事件)/,
    /(損害賠償請求事件|貸金返還請求事件|建物明渡請求事件|不当利得返還請求事件|(?:[\u4e00-\u9fff]+請求事件))/,
    // OCR: 分断されたパターン
    /([\u4e00-\u9fff]+\s+(?:請求|確認)\s*事件)/,
  ];
  for (const pattern of caseNamePatterns) {
    const match = normalizedText.match(pattern);
    if (match) {
      // スペース・改行を除去して正規化
      let caseName = match[0];
      // グループ1があればそちらを使う
      if (match[1] && !match[0].startsWith('損害')) {
        caseName = match[1];
      }
      // 「号」以降の部分だけ取る（先頭に「号」が含まれる場合）
      caseName = caseName.replace(/^号\s*/, '');
      // スペース除去して漢字のみ抽出
      caseName = caseName.replace(/[\s\n\r]+/g, '');
      // 「損害賠償」で始まらないゴミを除去
      const cleaned = caseName.match(/([\u4e00-\u9fff]+請求事件|[\u4e00-\u9fff]+確認事件)/);
      if (cleaned) {
        info.caseName = cleaned[1];
      } else if (caseName.includes('事件')) {
        info.caseName = caseName;
      }
      break;
    }
  }

  // --- 原告名・被告名 ---
  // 「当事者」セクションから抽出（最も確実）
  // FAX送信書の「当事者」欄: 「原告　山田民子 外1名」「被告　国立大学法人広島大学」
  // OCRでは「原 &」「原 告」「原告」などの表記ゆれがある

  // 当事者セクション内から原告・被告を探す（「当事者」の後に原告/被告が来るパターン）
  const partySection = normalizedText.match(
    /当\s*事\s*者[\s\S]{0,200}/
  );

  if (partySection) {
    const partySectionText = partySection[0];

    // 原告名: 「原告」or「原 告」or「原 &」（OCR誤認識）の後
    const plaintiffInParty = partySectionText.match(
      /原\s*[告&]\s*[_\s]*([^\n原被]{1,40})/
    );
    if (plaintiffInParty) {
      let name = plaintiffInParty[1].trim();
      // 「外N名」の処理
      name = name.replace(/\s*(外\s*\d+\s*名)\s*$/, (_, suffix) => {
        return ' ' + suffix.replace(/\s+/g, '');
      });
      const parts = name.split(/ (外\d+名)$/);
      if (parts.length > 1) {
        info.plaintiffName = parts[0].replace(/\s+/g, '') + ' ' + parts[1];
      } else {
        info.plaintiffName = name.replace(/\s+/g, '');
      }
    }

    // 被告名: 「被告」or「被 告」の後
    const defendantInParty = partySectionText.match(
      /被\s*告\s*[_\s]*([^\n原被]{1,40})/
    );
    if (defendantInParty) {
      let name = defendantInParty[1].trim();
      name = name.replace(/\s*(外\s*\d+\s*名)\s*$/, (_, suffix) => {
        return ' ' + suffix.replace(/\s+/g, '');
      });
      const parts = name.split(/ (外\d+名)$/);
      if (parts.length > 1) {
        info.defendantName = parts[0].replace(/\s+/g, '') + ' ' + parts[1];
      } else {
        info.defendantName = name.replace(/\s+/g, '');
      }
    }
  }

  // 当事者セクションで見つからなかった場合のフォールバック
  if (!info.plaintiffName) {
    const plaintiffPatterns = [
      /[【\[]\s*原\s*告\s*[】\]]\s*([^\n【\[]{1,30})/,
      // 「原告」の直後に人名（訴訟代理人ではなく）
      /原\s*告\s+(?!.*(?:訴訟|代理))([^\n（(被代訴]{1,20})/,
    ];
    for (const pattern of plaintiffPatterns) {
      const match = normalizedText.match(pattern);
      if (match) {
        let name = match[1].trim();
        name = name.replace(/\s*(外\s*\d+\s*名)\s*$/, (_, suffix) => {
          return ' ' + suffix.replace(/\s+/g, '');
        });
        const parts = name.split(/ (外\d+名)$/);
        if (parts.length > 1) {
          info.plaintiffName = parts[0].replace(/\s+/g, '') + ' ' + parts[1];
        } else {
          info.plaintiffName = name.replace(/\s+/g, '');
        }
        break;
      }
    }
  }

  if (!info.defendantName) {
    const defendantPatterns = [
      /[【\[]\s*(?:被|a)\s*告\s*[】\]]\s*([^\n【\[]{1,30})/,
      // 「被告」の直後に組織名・人名（訴訟代理人ではなく）
      /被\s*告\s+(?!.*(?:訴訟|代理))([^\n（(原代訴]{1,30})/,
    ];
    for (const pattern of defendantPatterns) {
      const match = normalizedText.match(pattern);
      if (match) {
        let name = match[1].trim();
        name = name.replace(/\s*(外\s*\d+\s*名)\s*$/, (_, suffix) => {
          return ' ' + suffix.replace(/\s+/g, '');
        });
        const parts = name.split(/ (外\d+名)$/);
        if (parts.length > 1) {
          info.defendantName = parts[0].replace(/\s+/g, '') + ' ' + parts[1];
        } else {
          info.defendantName = name.replace(/\s+/g, '');
        }
        break;
      }
    }
  }

  // --- 原告代理人弁護士名（全マッチから最もクリーンなものを選ぶ） ---
  const lawyerCandidates = [];

  // ヘルパー: OCR結果から弁護士名をクリーンアップ
  // 「石 口 R "pet i」→ CJK文字のみ抽出 → 「石口」
  // 「石 口 俊 一」→ CJK文字のみ抽出 → 「石口俊一」
  function cleanLawyerName(rawName) {
    // まず空白除去
    let name = rawName.trim();
    // CJK文字（漢字・ひらがな・カタカナ）のみ抽出
    const cjkChars = name.match(/[\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]+/g);
    if (cjkChars) {
      name = cjkChars.join('');
    } else {
      return null;
    }
    // 不要な末尾文字を除去（「宛て」「宛」「殿」「様」「御中」等）
    name = name.replace(/宛て$/, '');
    name = name.replace(/[宛殿様御中方和]+$/, '');
    // 日本人名として典型的な長さは2-4文字。5文字以上はOCRノイズの可能性
    // 末尾の1文字がOCRノイズの場合がある（例: 「石口俊一知」→「石口俊一」）
    if (name.length === 5) {
      // 5文字の場合、末尾を除去したバージョンも候補に
      // ただし、5文字の名前も実在するので両方返す（呼び出し側で判断）
      name = name; // そのまま
    }
    // 弁護士名として短すぎる or 長すぎる場合は除外
    if (name.length < 2 || name.length > 6) return null;
    // 「弁護士法」等の誤抽出を除外
    if (/^[法会事件番号裁判]/.test(name)) return null;
    return name;
  }

  // パターン1: 「原告訴訟代理人弁護士 NAME」（差出人欄 - 最優先）
  // OCR: 「原告 訴訟 代理 人 弁護 士 石 口 R "pet i」
  const formalPattern = /原告\s*(?:ら)?\s*(?:訴\s*訟)?\s*代理\s*人\s*弁護\s*士\s*([^\n]{2,20})/g;
  let lm;
  while ((lm = formalPattern.exec(normalizedText)) !== null) {
    const name = cleanLawyerName(lm[1]);
    if (name) lawyerCandidates.push({ name, priority: 1 });
  }

  // パターン2: 「人 弁護士 NAME」（差出人欄のバリエーション）
  const senderPattern = /人\s*弁護\s*士\s*([^\n]{2,20})/g;
  while ((lm = senderPattern.exec(normalizedText)) !== null) {
    // 「被告」の近くにあるものは自分（被告代理人）なので除外
    const contextBefore = normalizedText.substring(Math.max(0, lm.index - 30), lm.index);
    if (contextBefore.includes('被告')) continue;
    const name = cleanLawyerName(lm[1]);
    if (name) lawyerCandidates.push({ name, priority: 2 });
  }

  // パターン3: 「弁護士 NAME 宛て」パターン（受領書形式）
  const atePattern = /弁護\s*士\s*([^\n]{2,15})\s*宛/g;
  while ((lm = atePattern.exec(normalizedText)) !== null) {
    const name = cleanLawyerName(lm[1]);
    if (name) lawyerCandidates.push({ name, priority: 3 });
  }

  // パターン4: 「弁護士 NAME」（一般パターン）
  // 自事務所の弁護士は除外（config.jsonで設定）
  const ownLawyerNames = CONFIG.lawyerNames;
  const generalPattern = /弁護\s*士\s*([^\n]{2,15})/g;
  while ((lm = generalPattern.exec(normalizedText)) !== null) {
    const contextBefore = normalizedText.substring(Math.max(0, lm.index - 50), lm.index);
    if (contextBefore.includes('被告')) continue;
    const name = cleanLawyerName(lm[1]);
    if (name && !ownLawyerNames.some(own => name.includes(own))) {
      lawyerCandidates.push({ name, priority: 4 });
    }
  }

  if (lawyerCandidates.length > 0) {
    // 重複を除去
    const uniqueNames = [...new Set(lawyerCandidates.map(c => c.name))];
    const uniqueCandidates = uniqueNames.map(name => {
      const best = lawyerCandidates.filter(c => c.name === name)
        .sort((a, b) => a.priority - b.priority)[0];
      return best;
    });

    // 最も優先度が高いもの（priority小さい）を採用
    // 同優先度の場合: 3-4文字の名前を優先（日本人名の一般的な長さ）
    uniqueCandidates.sort((a, b) => {
      if (a.priority !== b.priority) return a.priority - b.priority;
      // 3-4文字の名前にボーナス
      const aIdeal = (a.name.length >= 3 && a.name.length <= 4) ? 0 : 1;
      const bIdeal = (b.name.length >= 3 && b.name.length <= 4) ? 0 : 1;
      if (aIdeal !== bIdeal) return aIdeal - bIdeal;
      return b.name.length - a.name.length;
    });

    let bestName = uniqueCandidates[0].name;

    // 5文字の名前は末尾がOCRノイズの可能性が高い
    // 例: 「石口俊一知」→「石口俊一」(知はOCRノイズ)
    // 短い候補（2文字 = 姓のみ）が存在する場合、5文字版の末尾はノイズと判断
    if (bestName.length === 5) {
      const shorter = uniqueCandidates.find(c => c.name.length <= 3 && bestName.startsWith(c.name));
      if (shorter) {
        // 姓の部分が一致 → 末尾1文字をOCRノイズとして除去
        bestName = bestName.substring(0, 4);
      }
    }

    info.plaintiffLawyer = bestName;
  }

  // --- 裁判所FAX番号（辞書引き） ---
  if (info.courtName) {
    // 「民事第N部」や係情報を除去して辞書キーとマッチ
    const courtBase = info.courtName
      .replace(/民事第[０-９\d]+部.*$/, '')
      .replace(/第[０-９\d]+[民刑]事部$/, '');
    info.courtFax = COURT_FAX_MAP[courtBase] || '';
  }

  // --- FAX番号の抽出（裁判所FAX と 原告代理人FAX を分離）---

  // 自事務所のFAX番号は除外（config.jsonで設定）
  const ownFaxPatterns = CONFIG.faxNumbers;
  const courtFaxValues = Object.values(COURT_FAX_MAP);

  // ヘルパー: FAX番号を正規化
  function normalizeFax(raw) {
    return raw
      .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
      .replace(/[－ー]/g, '-');
  }

  // 特別パターン1: 「裁判所...(FAX NNN-NNNN-NNNN)」のような明示的ラベルを先に検出
  // OCRでは「裁判所...御中\n(FAX 082-228-2306)」のように間に改行や他のテキストが入ることがある
  const explicitCourtFaxMatch = normalizedText.match(
    /裁\s*判\s*所[\s\S]{0,60}?[（(]\s*(?:FAX|ＦＡＸ|[Ff]ax)\s*([0-9０-９\-－ー]+)\s*[）)]/
  );
  if (explicitCourtFaxMatch) {
    info.courtFaxFromPdf = normalizeFax(explicitCourtFaxMatch[1]);
  }

  // 特別パターン2: 「と、石口 (FAX NNN-NNNN-NNNN)」のような原告代理人FAXの明示的ラベル
  // 「裁判所 (FAX ...) と、NAME (FAX ...)」パターンで2番目のFAXを取得
  const allExplicitFaxes = [];
  const explicitFaxRegex = /([\u4e00-\u9fff]{1,10})\s*[（(]\s*(?:FAX|ＦＡＸ|[Ff]ax)\s*([0-9０-９\-－ー]+)\s*[）)]/g;
  let efm;
  while ((efm = explicitFaxRegex.exec(normalizedText)) !== null) {
    allExplicitFaxes.push({ label: efm[1], fax: normalizeFax(efm[2]) });
  }
  // 裁判所以外の明示的FAXを原告代理人FAXとして設定
  for (const ef of allExplicitFaxes) {
    if (ef.label.includes('裁判') || ef.label.includes('裁判所')) continue;
    if (info.courtFaxFromPdf && ef.fax === info.courtFaxFromPdf) continue;
    const isOwn = ownFaxPatterns.some(p => ef.fax.includes(p));
    if (!isOwn && !info.plaintiffLawyerFax) {
      info.plaintiffLawyerFax = ef.fax;
    }
  }

  // 通常のFAX番号抽出（明示的パターンで取得できなかった場合のフォールバック）
  const faxRegex = /(?:FAX|ＦＡＸ|[Ff]ax)[：:\s]*([0-9０-９\-－ー]+)/g;
  const allFaxEntries = [];
  let faxMatch;
  while ((faxMatch = faxRegex.exec(normalizedText)) !== null) {
    const faxNum = normalizeFax(faxMatch[1]);
    allFaxEntries.push({ fax: faxNum, index: faxMatch.index });
  }

  for (const entry of allFaxEntries) {
    const isOwnFax = ownFaxPatterns.some(p => entry.fax.includes(p));
    if (isOwnFax) continue;

    // 既に明示的に裁判所FAXとして検出されたものはスキップ
    if (info.courtFaxFromPdf && entry.fax.includes(info.courtFaxFromPdf)) continue;

    const isKnownCourtFax = courtFaxValues.some(cf => entry.fax.includes(cf));

    // FAX送信書の構造解析:
    // 「原告訴訟代理人弁護士」の直後のTEL/FAXは原告代理人のもの
    // 「被告訴訟代理人弁護士」の直後のTEL/FAXは被告代理人（＝自分）のもの
    // 裁判所のFAXは辞書から引くのが最も確実
    const textBefore = normalizedText.substring(
      Math.max(0, entry.index - 200), entry.index
    );
    const textAfter = normalizedText.substring(
      entry.index, Math.min(normalizedText.length, entry.index + 100)
    );

    // この FAX の直前に「原告」「訴訟代理人」「弁護士」があるか
    const isNearPlaintiffLawyer = /原告\s*(?:ら)?\s*訴\s*訟\s*代\s*理\s*人/.test(textBefore) ||
      (/弁護\s*士/.test(textBefore) && !textBefore.includes('被告'));

    // この FAX の直前に「被告訴訟代理人」があるか（＝自分のFAX）
    const isNearDefendantLawyer = /被告\s*(?:ら)?\s*訴\s*訟\s*代\s*理\s*人/.test(textBefore);

    if (isNearDefendantLawyer) {
      // 被告代理人（自分）のFAX → スキップ
      continue;
    }

    if (isKnownCourtFax) {
      // 辞書に登録されている裁判所FAX
      if (!info.courtFaxFromPdf) {
        info.courtFaxFromPdf = entry.fax;
      }
    } else if (isNearPlaintiffLawyer) {
      // 原告代理人の近くにあるFAX → 原告代理人FAX
      if (!info.plaintiffLawyerFax) {
        info.plaintiffLawyerFax = entry.fax;
      }
    } else {
      // それ以外: 原告代理人FAX候補（裁判所FAXでなければ）
      if (!info.plaintiffLawyerFax) {
        info.plaintiffLawyerFax = entry.fax;
      }
    }
  }

  // 裁判所FAX: PDFから取れたら優先、なければ辞書
  if (info.courtFaxFromPdf) {
    info.courtFax = info.courtFaxFromPdf;
  }

  return info;
}

// ===== 全角数字変換 =====
function toFullWidthNumber(str) {
  return str.replace(/[0-9]/g, c => String.fromCharCode(c.charCodeAt(0) + 0xFEE0));
}

// ===== 今日の日付を令和で取得 =====
function getTodayReiwa() {
  const now = new Date();
  const year = now.getFullYear() - 2018;
  const month = now.getMonth() + 1;
  const day = now.getDate();
  return { year, month, day };
}

// ===== XML安全テキスト置換 =====
// docx XML内でテキストが複数のw:r/w:tに分割されていても安全に置換する。
// 段落(<w:p>)単位で処理。oldTextが跨るw:tのみを修正し、他のw:tは保持する。
function safeReplaceInXml(xml, oldText, newText) {
  const paraRegex = /(<w:p[\s>][\s\S]*?<\/w:p>)/g;
  return xml.replace(paraRegex, (paraXml) => {
    // この段落内の全w:tを収集
    const wtRegex = /<w:t([^>]*)>([^<]*)<\/w:t>/g;
    const segments = [];
    let m;
    while ((m = wtRegex.exec(paraXml)) !== null) {
      segments.push({
        fullMatch: m[0],
        attrs: m[1],
        text: m[2],
        index: m.index,
      });
    }

    if (segments.length === 0) return paraXml;

    const joinedText = segments.map(s => s.text).join('');

    // このoldTextがこの段落に含まれるか
    if (!joinedText.includes(oldText)) return paraXml;

    // oldTextが段落内のどの位置にあるかを特定
    const matchStart = joinedText.indexOf(oldText);
    const matchEnd = matchStart + oldText.length;

    // 各セグメントの累積開始/終了位置を計算
    let cumulative = 0;
    for (const seg of segments) {
      seg.startPos = cumulative;
      seg.endPos = cumulative + seg.text.length;
      cumulative += seg.text.length;
    }

    // oldTextが跨るセグメントを特定
    const affectedSegs = segments.filter(
      seg => seg.endPos > matchStart && seg.startPos < matchEnd
    );

    if (affectedSegs.length === 0) return paraXml;

    // 単一セグメント内で完結する場合：そのw:tだけを修正
    if (affectedSegs.length === 1) {
      const seg = affectedSegs[0];
      const localStart = matchStart - seg.startPos;
      const localEnd = matchEnd - seg.startPos;
      const newSegText = seg.text.substring(0, localStart) + newText + seg.text.substring(localEnd);
      const hasPreserve = seg.attrs.includes('xml:space="preserve"');
      const newAttrs = hasPreserve ? seg.attrs : ' xml:space="preserve"';
      const newWt = `<w:t${newAttrs}>${newSegText}</w:t>`;

      // このセグメントだけを置換（他はそのまま）
      return paraXml.substring(0, seg.index) + newWt +
        paraXml.substring(seg.index + seg.fullMatch.length);
    }

    // 複数セグメントに跨る場合：
    // 最初のセグメントに置換結果を入れ、中間・最後のセグメントからoldText部分を除去
    let segIdx = 0;
    const result = paraXml.replace(/<w:t([^>]*)>([^<]*)<\/w:t>/g, (match, attrs, text) => {
      const seg = segments[segIdx];
      segIdx++;

      if (!affectedSegs.includes(seg)) {
        // 影響を受けないセグメント → そのまま保持
        return match;
      }

      const isFirst = seg === affectedSegs[0];
      const isLast = seg === affectedSegs[affectedSegs.length - 1];
      const hasPreserve = attrs.includes('xml:space="preserve"');
      const newAttrs = hasPreserve ? attrs : ' xml:space="preserve"';

      if (isFirst && isLast) {
        // 1セグメントのみ（上で処理済みだが念のため）
        const localStart = matchStart - seg.startPos;
        const localEnd = matchEnd - seg.startPos;
        return `<w:t${newAttrs}>${text.substring(0, localStart)}${newText}${text.substring(localEnd)}</w:t>`;
      } else if (isFirst) {
        // 最初のセグメント：matchStartより前の部分を残し、newTextを追加
        const localStart = matchStart - seg.startPos;
        return `<w:t${newAttrs}>${text.substring(0, localStart)}${newText}</w:t>`;
      } else if (isLast) {
        // 最後のセグメント：matchEndより後の部分を残す
        const localEnd = matchEnd - seg.startPos;
        const remaining = text.substring(localEnd);
        if (remaining.length > 0) {
          return `<w:t${newAttrs}>${remaining}</w:t>`;
        } else {
          return `<w:t${attrs}></w:t>`;
        }
      } else {
        // 中間のセグメント：oldTextに完全に含まれるので空にする
        return `<w:t${attrs}></w:t>`;
      }
    });

    return result;
  });
}

// ===== テンプレートへの差し込み処理（共通ロジック） =====
function applyInfoToTemplate(docXml, info, documentTitle) {
  const today = getTodayReiwa();

  // --- 裁判所名 ---
  if (info.courtName) {
    const ORIG_COURT = '神戸地方裁判所尼崎支部第２民事部'; // 14文字
    const courtDiff = ORIG_COURT.length - info.courtName.length;
    const courtPad = courtDiff > 0 ? '　'.repeat(courtDiff) : '';
    docXml = safeReplaceInXml(docXml, ORIG_COURT, info.courtName + courtPad);
  }

  // --- 裁判所FAX番号 ---
  if (info.courtFax) {
    docXml = safeReplaceInXml(docXml, '06-6438-1710', info.courtFax);
    const fullWidthFax = toFullWidthNumber(info.courtFax).replace(/-/g, '－');
    docXml = safeReplaceInXml(docXml, '０６―６４３８－１７１０', fullWidthFax);
  }

  // --- 原告代理人弁護士名 ---
  if (info.plaintiffLawyer) {
    const ORIG_LAWYER = '四方久寛'; // 4文字
    const lawyerDiff = ORIG_LAWYER.length - info.plaintiffLawyer.length;
    const lawyerPad = lawyerDiff > 0 ? '　'.repeat(lawyerDiff) : '';
    docXml = safeReplaceInXml(docXml, ORIG_LAWYER, info.plaintiffLawyer + lawyerPad);
  }

  // --- 原告代理人FAX番号 ---
  if (info.plaintiffLawyerFax) {
    docXml = safeReplaceInXml(docXml, '06-4708-3638', info.plaintiffLawyerFax);
  }

  // --- 日付（送付書） ---
  docXml = safeReplaceInXml(docXml, '令和6年11月7日', `令和${today.year}年${today.month}月${today.day}日`);

  // --- 日付（受領証明書） ---
  docXml = safeReplaceInXml(docXml, '令和6年9月', `令和${today.year}年${today.month}月`);

  // --- 事件番号 ---
  if (info.caseNumber) {
    const fullWidthCaseNumber = toFullWidthNumber(info.caseNumber);
    docXml = safeReplaceInXml(docXml, '令和３年（ワ）第８００号', fullWidthCaseNumber);
  }

  // --- 事件名 ---
  if (info.caseName) {
    docXml = safeReplaceInXml(docXml, '損害賠償請求事件', info.caseName);
  }

  // --- 原告名 ---
  if (info.plaintiffName) {
    docXml = safeReplaceInXml(docXml, '木村治紀', info.plaintiffName);
  }

  // --- 被告名 ---
  if (info.defendantName) {
    docXml = safeReplaceInXml(docXml, '独立行政法人国立病院機構', info.defendantName);
  }

  // --- 送付書類名 ---
  docXml = safeReplaceInXml(docXml, '被告第９準備書面', documentTitle);

  return docXml;
}

// ===== Word生成の共通出力処理 =====
async function writeDocxOutput(zip, documentTitle, outputDir) {
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
  const outputFileName = `文書送付書_${documentTitle}_${timestamp}.docx`;
  const outputPath = path.join(outputDir, outputFileName);

  const outputBuffer = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(outputPath, outputBuffer);

  return { outputPath, outputFileName, outputBuffer };
}

// ===== メイン処理 =====
async function generateDocument(pdfPath, documentTitle) {
  // 1. PDF読み取り
  console.log('📄 PDFを読み取り中...');
  const pdfText = await extractTextFromPDF(pdfPath);
  console.log('--- PDF抽出テキスト (先頭500文字) ---');
  console.log(pdfText.substring(0, 500));
  console.log('---');

  // 2. 情報抽出
  console.log('\n🔍 情報を抽出中...');
  const info = extractInfoFromText(pdfText);
  console.log('抽出結果:', JSON.stringify(info, null, 2));

  // 3. 送付書類名
  if (!documentTitle) {
    let baseName = path.basename(pdfPath, path.extname(pdfPath));
    baseName = baseName.replace(/^【[^】]+】\s*/, '');
    baseName = baseName.replace(/^[\u4e00-\u9fff]+事案[\s　]+/, '');
    documentTitle = baseName;
  }

  // 4. テンプレート読み込み＆差し込み
  console.log('\n📝 テンプレートに差し込み中...');
  const templateData = fs.readFileSync(TEMPLATE_PATH);
  const zip = await JSZip.loadAsync(templateData);
  let docXml = await zip.file('word/document.xml').async('string');

  docXml = applyInfoToTemplate(docXml, info, documentTitle);
  zip.file('word/document.xml', docXml);

  // 5. 出力
  const { outputPath } = await writeDocxOutput(zip, documentTitle, OUTPUT_DIR);

  console.log(`\n✅ 文書送付書を生成しました！`);
  console.log(`📁 出力先: ${outputPath}`);
  console.log('\n--- 差し込み内容 ---');
  console.log(`裁判所名: ${info.courtName || '(未検出)'}`);
  console.log(`裁判所FAX: ${info.courtFax || '(未検出 - 手動入力が必要)'}`);
  console.log(`事件番号: ${info.caseNumber || '(未検出)'}${info.caseNumberGuessed ? ' ⚠️ OCR推測値' : ''}`);
  console.log(`事件名: ${info.caseName || '(未検出)'}`);
  console.log(`原告: ${info.plaintiffName || '(未検出)'}`);
  console.log(`被告: ${info.defendantName || '(未検出)'}`);
  console.log(`原告代理人: ${info.plaintiffLawyer || '(未検出)'}`);
  console.log(`原告代理人FAX: ${info.plaintiffLawyerFax || '(未検出 - 手動入力が必要)'}`);
  console.log(`送付書類名: ${documentTitle}`);
  console.log(`送付日: 令和${getTodayReiwa().year}年${getTodayReiwa().month}月${getTodayReiwa().day}日`);

  if (info.caseNumberGuessed) {
    console.log('\n⚠️  注意: 事件番号はOCRの認識精度が低かったため推測値です。');
    console.log('   生成されたWordファイルを開いて事件番号をご確認ください。');
  }

  return outputPath;
}

// ===== Web UI用: ユーザー編集済み情報からWord生成 =====
async function generateDocumentFromInfo(info, documentTitle, outputDir) {
  outputDir = outputDir || OUTPUT_DIR;

  const templateData = fs.readFileSync(TEMPLATE_PATH);
  const zip = await JSZip.loadAsync(templateData);
  let docXml = await zip.file('word/document.xml').async('string');

  docXml = applyInfoToTemplate(docXml, info, documentTitle);
  zip.file('word/document.xml', docXml);

  return await writeDocxOutput(zip, documentTitle, outputDir);
}

// ===== PDF送付書類名の自動取得 =====
function getDocumentTitleFromFilename(pdfPath) {
  let baseName = path.basename(pdfPath, path.extname(pdfPath));
  baseName = baseName.replace(/^【[^】]+】\s*/, '');
  baseName = baseName.replace(/^[\u4e00-\u9fff]+事案[\s　]+/, '');
  return baseName;
}

// ===== エントリーポイント =====
async function main() {
  console.log(`[起動] ${new Date().toISOString()}`);
  console.log(`[引数] ${JSON.stringify(process.argv)}`);
  console.log(`[CWD]  ${process.cwd()}`);
  console.log(`[__dirname] ${__dirname}`);

  const args = process.argv.slice(2);

  if (args.length === 0) {
    console.log('==========================================');
    console.log('  文書送付書 自動生成システム');
    console.log('==========================================');
    console.log('');
    console.log('使い方:');
    console.log('  node system/generate.js <PDFファイルパス> [送付書類名]');
    console.log('');
    console.log('例:');
    console.log('  node system/generate.js 準備書面.pdf "被告第10準備書面"');
    console.log('  node system/generate.js 準備書面.pdf');
    console.log('');
    console.log('※ 送付書類名を省略した場合、PDFファイル名を使用します。');
    process.exit(0);
  }

  // --- PDFパスの解決 ---
  let pdfPath;
  let documentTitle = null;

  // VBSラッパーからの呼び出し: --from-file フラグで一時ファイルからパスを読む
  const PDF_PATH_FILE = path.join(__dirname, '_pdf_path.txt');
  if (args[0] === '--from-file' && fs.existsSync(PDF_PATH_FILE)) {
    let rawPath = fs.readFileSync(PDF_PATH_FILE, 'utf-8').trim();
    // BOM除去
    if (rawPath.charCodeAt(0) === 0xFEFF) rawPath = rawPath.slice(1);
    pdfPath = path.resolve(rawPath);
    console.log(`[VBS経由] パスファイルから読み込み: ${pdfPath}`);
    // 一時ファイル削除
    try { fs.unlinkSync(PDF_PATH_FILE); } catch (e) { /* ignore */ }
  } else {
    // 通常の引数渡し
    // ドラッグ＆ドロップ時、ファイル名にスペースがあるとbatが引数を分割することがある。
    pdfPath = path.resolve(args[0]);

    if (!fs.existsSync(pdfPath)) {
      // 引数を全部つなげてパスとして試す
      const joined = args.join(' ');
      const joinedPath = path.resolve(joined);
      console.log(`[引数分割検出] args[0]="${args[0]}" が見つからないため、結合を試行: "${joined}"`);
      if (fs.existsSync(joinedPath)) {
        pdfPath = joinedPath;
      } else {
        // .pdf で終わるところまでをパスとして結合
        let pathPart = '';
        for (let i = 0; i < args.length; i++) {
          pathPart += (i > 0 ? ' ' : '') + args[i];
          if (pathPart.toLowerCase().endsWith('.pdf')) {
            const testPath = path.resolve(pathPart);
            if (fs.existsSync(testPath)) {
              pdfPath = testPath;
              documentTitle = args.slice(i + 1).join(' ') || null;
              break;
            }
          }
        }
      }
    } else {
      documentTitle = args[1] || null;
    }
  }

  console.log(`[PDFパス] ${pdfPath}`);
  console.log(`[テンプレート] ${TEMPLATE_PATH}`);
  console.log(`[テンプレート存在] ${fs.existsSync(TEMPLATE_PATH)}`);

  if (!fs.existsSync(pdfPath)) {
    console.error(`❌ エラー: ファイルが見つかりません: ${pdfPath}`);
    console.error(`   受け取った引数: ${JSON.stringify(args)}`);
    process.exit(1);
  }

  try {
    await generateDocument(pdfPath, documentTitle);
  } catch (err) {
    console.error(`❌ エラーが発生しました:`, err.message);
    console.error(err.stack);
    process.exit(1);
  }
}

// ===== エントリーポイント / モジュールエクスポート =====
if (require.main === module) {
  main();
} else {
  module.exports = {
    extractTextFromPDF,
    extractInfoFromText,
    generateDocument,
    generateDocumentFromInfo,
    getDocumentTitleFromFilename,
    getTodayReiwa,
    toFullWidthNumber,
    COURT_FAX_MAP,
    TEMPLATE_PATH,
    OUTPUT_DIR,
  };
}
