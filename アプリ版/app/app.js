/**
 * 文書送付書・受領書 自動でつくる君 - 統合ブラウザ版
 *
 * 3ファイル（generate-browser.js, receipt-browser.js, app.js）を統合。
 * pdf.js + Tesseract.js + JSZip + pdf-lib で全てブラウザ内で処理。
 * サーバー不要・インストール不要。
 */
(function() {
  'use strict';

  // =============================================
  // Section 1: 共通設定
  // =============================================

  const CMAP_URL = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/';

  function getConfig() {
    try {
      return JSON.parse(localStorage.getItem('tsukurukun_config') || '{}');
    } catch (e) { return {}; }
  }

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

  function toFullWidthNumber(str) {
    return str.replace(/[0-9]/g, c => String.fromCharCode(c.charCodeAt(0) + 0xFEE0));
  }

  function getTodayReiwa() {
    const now = new Date();
    const year = now.getFullYear() - 2018;
    const month = now.getMonth() + 1;
    const day = now.getDate();
    return { year, month, day };
  }

  // --- IndexedDBキャッシュ（フォント・テンプレート用）---
  async function _idbGet(storeName, key) {
    return new Promise(resolve => {
      try {
        const req = indexedDB.open('tsukurukun_cache', 1);
        req.onupgradeneeded = e => {
          const db = e.target.result;
          ['fonts', 'templates'].forEach(s => {
            if (!db.objectStoreNames.contains(s)) db.createObjectStore(s);
          });
        };
        req.onsuccess = () => {
          try {
            const g = req.result.transaction(storeName, 'readonly').objectStore(storeName).get(key);
            g.onsuccess = () => resolve(g.result || null);
            g.onerror = () => resolve(null);
          } catch (e) { resolve(null); }
        };
        req.onerror = () => resolve(null);
      } catch (e) { resolve(null); }
    });
  }

  async function _idbPut(storeName, key, value) {
    try {
      const req = indexedDB.open('tsukurukun_cache', 1);
      req.onupgradeneeded = e => {
        const db = e.target.result;
        ['fonts', 'templates'].forEach(s => {
          if (!db.objectStoreNames.contains(s)) db.createObjectStore(s);
        });
      };
      req.onsuccess = () => {
        try { req.result.transaction(storeName, 'readwrite').objectStore(storeName).put(value, key); }
        catch (e) { /* ignore */ }
      };
    } catch (e) { /* ignore */ }
  }

  // --- 日本語フォント読み込み（IndexedDBキャッシュ + CDN + ファイル選択フォールバック）---
  let _cachedFontBytes = null;
  async function loadJapaneseFont() {
    if (_cachedFontBytes) return _cachedFontBytes;

    // 1. IndexedDBキャッシュ
    const cached = await _idbGet('fonts', 'NotoSerifJP');
    if (cached) {
      console.log('[フォント] IndexedDBキャッシュから読み込み');
      _cachedFontBytes = cached;
      return _cachedFontBytes;
    }

    // 2. fetch（ローカル → CDN）
    const urls = [
      'fonts/NotoSerifJP.ttf',
      'https://cdn.jsdelivr.net/gh/Rachmaninovpiano/bunsho-sofusho-juryosho-auto@master/docs/fonts/NotoSerifJP.ttf',
    ];
    for (const url of urls) {
      try {
        console.log('[フォント] 読み込み試行:', url);
        const resp = await fetch(url);
        if (resp.ok) {
          _cachedFontBytes = await resp.arrayBuffer();
          console.log('[フォント] 読み込み成功:', url, _cachedFontBytes.byteLength, 'bytes');
          await _idbPut('fonts', 'NotoSerifJP', _cachedFontBytes);
          return _cachedFontBytes;
        }
      } catch (e) {
        console.warn('[フォント] 読み込み失敗:', url, e.message);
        continue;
      }
    }

    // 3. ファイル選択ダイアログ（file://プロトコル用）
    return new Promise((resolve, reject) => {
      alert('フォントの自動読み込みに失敗しました。\n\nfonts/NotoSerifJP.ttf を手動で選択してください。\n（一度選択すると次回以降はキャッシュされます）');
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.ttf,.otf';
      input.onchange = async () => {
        if (!input.files[0]) { reject(new Error('フォントが選択されませんでした')); return; }
        _cachedFontBytes = await input.files[0].arrayBuffer();
        await _idbPut('fonts', 'NotoSerifJP', _cachedFontBytes);
        console.log('[フォント] ファイル選択で読み込み成功、キャッシュ保存');
        resolve(_cachedFontBytes);
      };
      input.click();
    });
  }

  // --- テンプレート読み込み（CDNフォールバック + キャッシュ）---
  async function loadTemplate(localPath, cdnPath, cacheKey) {
    // 1. IndexedDBキャッシュ
    const cached = await _idbGet('templates', cacheKey);
    if (cached) {
      console.log('[テンプレート] キャッシュから読み込み:', cacheKey);
      return cached;
    }
    // 2. fetch（ローカル → CDN）
    const urls = [localPath];
    if (cdnPath) urls.push(cdnPath);
    for (const url of urls) {
      try {
        const resp = await fetch(url);
        if (resp.ok) {
          const data = await resp.arrayBuffer();
          await _idbPut('templates', cacheKey, data);
          console.log('[テンプレート] 読み込み成功:', url);
          return data;
        }
      } catch (e) { continue; }
    }
    throw new Error('テンプレートの読み込みに失敗しました: ' + cacheKey);
  }

  // =============================================
  // Section 2: PDFテキスト抽出（2パス最適化）
  // =============================================

  // 指定ページ群からテキスト抽出（座標ベース改行）
  async function extractPagesText(pdfDoc, pageNums) {
    let text = '';
    for (const i of pageNums) {
      const page = await pdfDoc.getPage(i);
      const content = await page.getTextContent();
      let lastY = null;
      let lastEndX = null;
      for (const item of content.items) {
        if (!item.str && item.str !== '') continue;
        const tx = item.transform;
        if (tx) {
          const y = Math.round(tx[5]);
          const x = tx[4];
          const itemWidth = item.width || 0;
          if (lastY !== null && Math.abs(y - lastY) > 3) {
            text += '\n';
            lastEndX = null;
          } else if (lastEndX !== null && x > lastEndX + 5) {
            text += ' ';
          }
          text += item.str;
          lastY = y;
          lastEndX = x + itemWidth;
        } else {
          text += item.str;
        }
        if (item.hasEOL) {
          text += '\n';
          lastY = null;
          lastEndX = null;
        }
      }
      text += '\n\n';
    }
    return text;
  }

  // 2パス方式: まず[1,2,末尾]→プローブ→不足時のみ残りページ
  async function extractTextFromPDFBrowser(pdfArrayBuffer, onProgress) {
    onProgress && onProgress('PDFからテキストを抽出中...');
    const pdfDoc = await pdfjsLib.getDocument({
      data: pdfArrayBuffer,
      cMapUrl: CMAP_URL,
      cMapPacked: true,
    }).promise;
    const totalPages = pdfDoc.numPages;

    // Pass 1: ページ [1, 2, 末尾]（重複除去）
    const pass1Set = new Set([1, Math.min(2, totalPages), totalPages]);
    const pass1Pages = [...pass1Set].sort((a, b) => a - b);
    console.log('[抽出] Pass1: ページ', pass1Pages.join(','), '/', totalPages);
    const pass1Text = await extractPagesText(pdfDoc, pass1Pages);

    // プローブ: courtName + caseNumber があれば十分
    const probe = extractInfoFromText(pass1Text);
    if (probe.courtName && probe.caseNumber) {
      console.log('[抽出] Pass1完了:', pass1Text.length, '文字');
      return pass1Text;
    }

    // Pass 2: 残りページ
    const remaining = [];
    for (let i = 1; i <= totalPages; i++) {
      if (!pass1Set.has(i)) remaining.push(i);
    }
    if (remaining.length === 0) {
      console.log('[抽出] 全ページ読み取り済み:', pass1Text.length, '文字');
      return pass1Text;
    }

    console.log('[抽出] Pass2: ページ', remaining.join(','));
    onProgress && onProgress('追加ページを読み取り中...');
    const pass2Text = await extractPagesText(pdfDoc, remaining);
    const fullText = pass1Text + pass2Text;
    console.log('[抽出] Pass1+2完了:', fullText.length, '文字');
    return fullText;
  }

  // OCR: ページ[1, 末尾]のみ（OCRは重いため最小限に）
  async function extractTextWithOCRBrowser(pdfArrayBuffer, onProgress) {
    onProgress && onProgress('画像PDFを検出。OCRで文字認識中...');
    const pdfDoc = await pdfjsLib.getDocument({
      data: pdfArrayBuffer,
      cMapUrl: CMAP_URL,
      cMapPacked: true,
    }).promise;
    const totalPages = pdfDoc.numPages;
    const ocrPages = totalPages === 1 ? [1] : [1, totalPages];
    let allText = '';

    for (let idx = 0; idx < ocrPages.length; idx++) {
      const pageNum = ocrPages[idx];
      onProgress && onProgress(`ページ ${pageNum}/${totalPages} をOCR中...`);
      const page = await pdfDoc.getPage(pageNum);
      const viewport = page.getViewport({ scale: 400 / 72 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport }).promise;

      const worker = await Tesseract.createWorker('jpn', 1, {
        logger: m => {
          if (m.status === 'recognizing text' && onProgress) {
            onProgress(`ページ${pageNum} OCR処理中... ${Math.round((m.progress || 0) * 100)}%`);
          }
        }
      });
      const { data } = await worker.recognize(canvas);
      await worker.terminate();
      if (data && data.text) {
        allText += data.text + '\n';
      }
    }
    return allText;
  }

  // テキスト抽出 → OCRフォールバック
  async function extractTextBrowser(pdfArrayBuffer, onProgress) {
    try {
      const text = await extractTextFromPDFBrowser(pdfArrayBuffer, onProgress);
      const trimmed = text.replace(/[\s\n\r]/g, '');
      console.log('[抽出] テキスト結果:', trimmed.length, '文字 (空白除去後)');
      if (trimmed.length < 10) {
        console.log('[抽出] テキスト埋め込みなし → OCRに切り替え');
        return await extractTextWithOCRBrowser(pdfArrayBuffer, onProgress);
      }
      return text;
    } catch (err) {
      console.error('[抽出] エラー:', err);
      console.log('[抽出] テキスト抽出失敗 → OCRに切り替え');
      onProgress && onProgress('テキスト抽出に失敗。OCRで文字認識中...');
      return await extractTextWithOCRBrowser(pdfArrayBuffer, onProgress);
    }
  }

  // =============================================
  // Section 2b: Word (.docx) テキスト抽出
  // =============================================

  async function extractTextFromDocx(arrayBuffer, onProgress) {
    onProgress && onProgress('Wordファイルからテキストを抽出中...');
    const zip = await JSZip.loadAsync(arrayBuffer);
    const docXmlFile = zip.file('word/document.xml');
    if (!docXmlFile) {
      throw new Error('Word文書の解析に失敗しました（document.xmlが見つかりません）');
    }
    const xmlStr = await docXmlFile.async('string');
    let text = '';
    var endParaTag = new RegExp('<\/w:p>', 'g');
    var wtTagGlobal = new RegExp('<w:t[^>]*>([^<]*)<\/w:t>', 'g');
    var wtTagSingle = new RegExp('<w:t[^>]*>([^<]*)<\/w:t>');
    var paragraphs = xmlStr.split(endParaTag);
    for (var pi = 0; pi < paragraphs.length; pi++) {
      var para = paragraphs[pi];
      var textMatches = para.match(wtTagGlobal);
      if (textMatches) {
        var line = '';
        for (var ti = 0; ti < textMatches.length; ti++) {
          var inner = textMatches[ti].match(wtTagSingle);
          if (inner) line += inner[1];
        }
        text += line + '\n';
      }
    }
    console.log('[抽出] Word文書テキスト:', text.length, '文字');
    return text;
  }

  // =============================================
  // Section 2c: テキスト正規化（OCR誤読修正）
  // =============================================

  function normalizeExtractedText(text) {
    let t = text;
    // 全角数字→半角
    t = t.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
    // 全角英字→半角
    t = t.replace(/[Ａ-Ｚａ-ｚ]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
    // 特殊スペース→通常スペース
    t = t.replace(/[  -​　﻿]/g, ' ');
    // OCR誤読: よくある文字化けパターン修正
    t = t.replace(/裁判\s*所/g, '裁判所');
    t = t.replace(/地方\s*裁判/g, '地方裁判');
    t = t.replace(/高等\s*裁判/g, '高等裁判');
    t = t.replace(/家庭\s*裁判/g, '家庭裁判');
    t = t.replace(/簡易\s*裁判/g, '簡易裁判');
    t = t.replace(/弁護\s*士/g, '弁護士');
    t = t.replace(/原\s*告/g, '原告');
    t = t.replace(/被\s*告/g, '被告');
    t = t.replace(/事\s*件/g, '事件');
    t = t.replace(/損\s*害\s*賠\s*償/g, '損害賠償');
    t = t.replace(/請\s*求/g, '請求');
    t = t.replace(/訴\s*訟\s*代\s*理\s*人/g, '訴訟代理人');
    t = t.replace(/令\s*和/g, '令和');
    t = t.replace(/平\s*成/g, '平成');
    // 連続スペースを1つに
    t = t.replace(/ {2,}/g, ' ');
    return t;
  }

  // =============================================
  // Section 3: 文書送付書 情報抽出
  // =============================================

  function extractInfoFromText(text) {
    const config = getConfig();
    const info = {};
    const normalizedText = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const cleanText = normalizeExtractedText(normalizedText);

    // 都市名リスト
    const cityNames = '東京|大阪|名古屋|広島|福岡|仙台|札幌|高松|京都|神戸|横浜|さいたま|千葉|山口|岡山|福山|松山|高知|那覇|長崎|熊本|鹿児島|大分|宮崎|佐賀|秋田|青森|盛岡|山形|福島|水戸|宇都宮|前橋|甲府|長野|新潟|富山|金沢|福井|津|大津|奈良|和歌山|鳥取|松江|徳島|旭川|釧路|函館';

    // --- 裁判所名 ---
    const courtPattern = new RegExp(
      `((?:${cityNames})\\s*(?:地方|高等|家庭|簡易)\\s*裁判\\s*所(?:\\s*[\\u4e00-\\u9fff]+\\s*支部)?(?:\\s*民事\\s*第\\s*[０-９\\d]+\\s*部)?)`,
      'g'
    );
    let courtMatch;
    const courtCandidates = [];
    while ((courtMatch = courtPattern.exec(cleanText)) !== null) {
      const cleaned = courtMatch[1].replace(/\s+/g, '');
      courtCandidates.push(cleaned);
    }
    if (courtCandidates.length > 0) {
      info.courtName = courtCandidates.reduce((a, b) => a.length >= b.length ? a : b);
    }

    // --- 事件番号 ---
    const caseSymbols = 'ワヲネレモハノニナラ行わをねれもはのになら';
    const caseNumberPatterns = [
      new RegExp(`([令平]和\\d+年[（(][${caseSymbols}][）)]\\s*第?\\s*\\d+号)`),
      new RegExp(`([令平]\\s*和\\s*\\d+\\s*年\\s*[（(]\\s*[${caseSymbols}]\\s*[）)]\\s*第?\\s*\\d+\\s*号)`),
      new RegExp(`([令平]\\s*和\\s*\\d+\\s*年\\s*\\(\\s*[${caseSymbols}]\\s*\\)\\s*第?\\s*\\d+\\s*号)`),
      new RegExp(`(令\\s*和\\s*(\\d+)\\s*年\\s*[（(]\\s*([${caseSymbols}])\\s*[）)]\\s*第\\s*(\\d+)\\s*号)`),
    ];
    for (const pattern of caseNumberPatterns) {
      const match = cleanText.match(pattern);
      if (match) {
        let cn = match[1].replace(/\s+/g, '');
        cn = cn.replace(/（/g, '(').replace(/）/g, ')');
        info.caseNumber = cn;
        break;
      }
    }

    // パターン2: 「事件の表示」セクションから
    if (!info.caseNumber) {
      const displaySectionMatch = cleanText.match(
        /事\s*件\s*の\s*表\s*示[】\]\s]*([^\n]{1,80})/
      );
      if (displaySectionMatch) {
        const sectionText = displaySectionMatch[1];
        const fullMatch = sectionText.match(
          new RegExp(`令?\\s*和?\\s*(\\d+)\\s*年?\\s*[（(]\\s*([${caseSymbols}])\\s*[）)]\\s*第\\s*(\\d+)\\s*号`)
        );
        if (fullMatch) {
          info.caseNumber = `令和${fullMatch[1]}年(${fullMatch[2]})第${fullMatch[3]}号`;
        } else {
          const numMatch = sectionText.match(/第\s*(\d+)\s*号/);
          if (numMatch) {
            const caseNum = numMatch[1];
            const symbolMatch = sectionText.match(new RegExp(`[（(]\\s*([${caseSymbols}])\\s*[）)]`));
            const symbol = symbolMatch ? symbolMatch[1] : 'ワ';
            const guessed = !symbolMatch;
            const yearMatches = [];
            const yearRegex = /令\s*和\s*(\d+)\s*年/g;
            let ym;
            while ((ym = yearRegex.exec(cleanText)) !== null) {
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
      /損\s*害\s*賠\s*償[\s\S]{0,80}?請\s*求\s*事\s*件/,
      /(?:号\s*)([\u4e00-\u9fff][\u4e00-\u9fff\s]*(?:請\s*求|確\s*認|等?)\s*事\s*件)/,
      /(損\s*害\s*賠\s*償\s*請\s*求\s*事\s*件|貸\s*金\s*返\s*還\s*請\s*求\s*事\s*件|建\s*物\s*明\s*渡\s*請\s*求\s*事\s*件|不\s*当\s*利\s*得\s*返\s*還\s*請\s*求\s*事\s*件)/,
      /([\u4e00-\u9fff][\u4e00-\u9fff\s]*(?:請\s*求|確\s*認)\s*事\s*件)/,
    ];
    for (const pattern of caseNamePatterns) {
      const match = cleanText.match(pattern);
      if (match) {
        let caseName = match[0];
        if (match[1] && !match[0].startsWith('損害')) {
          caseName = match[1];
        }
        caseName = caseName.replace(/^号\s*/, '');
        caseName = caseName.replace(/[\s\n\r]+/g, '');
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
    const partySection = cleanText.match(/当\s*事\s*者[\s\S]{0,200}/);
    if (partySection) {
      const partySectionText = partySection[0];
      const plaintiffInParty = partySectionText.match(
        /原\s*[告&]\s*[_\s]*([^\n原被]{1,40})/
      );
      if (plaintiffInParty) {
        let name = plaintiffInParty[1].trim();
        name = name.replace(/\s*(外\s*\d+\s*名)\s*$/, (_, suffix) => ' ' + suffix.replace(/\s+/g, ''));
        const parts = name.split(/ (外\d+名)$/);
        if (parts.length > 1) {
          info.plaintiffName = parts[0].replace(/\s+/g, '') + ' ' + parts[1];
        } else {
          info.plaintiffName = name.replace(/\s+/g, '');
        }
      }
      const defendantInParty = partySectionText.match(
        /被\s*告\s*[_\s]*([^\n原被]{1,40})/
      );
      if (defendantInParty) {
        let name = defendantInParty[1].trim();
        name = name.replace(/\s*(外\s*\d+\s*名)\s*$/, (_, suffix) => ' ' + suffix.replace(/\s+/g, ''));
        const parts = name.split(/ (外\d+名)$/);
        if (parts.length > 1) {
          info.defendantName = parts[0].replace(/\s+/g, '') + ' ' + parts[1];
        } else {
          info.defendantName = name.replace(/\s+/g, '');
        }
      }
    }

    // フォールバック: 当事者セクションで見つからなかった場合
    if (!info.plaintiffName) {
      const plaintiffPatterns = [
        /[【\[［]\s*原\s*告\s*[】\]］]\s*\n?\s*([\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff][^\n【\[［]{1,30})/,
        /原\s*告\s+(?!.*(?:訴\s*訟|代\s*理))([\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff][^\n（(被代訴】\]］]{1,20})/,
      ];
      for (const pattern of plaintiffPatterns) {
        const allMatches = [];
        const globalPattern = new RegExp(pattern.source, 'g');
        let match;
        while ((match = globalPattern.exec(cleanText)) !== null) {
          allMatches.push(match);
        }
        for (const match of allMatches) {
          let name = match[1].trim();
          const cleanedName = name.replace(/\s+/g, '');
          if (cleanedName.length <= 1) continue;
          name = name.replace(/\s*(外\s*\d+\s*名)\s*$/, (_, suffix) => ' ' + suffix.replace(/\s+/g, ''));
          const parts = name.split(/ (外\d+名)$/);
          if (parts.length > 1) {
            info.plaintiffName = parts[0].replace(/\s+/g, '') + ' ' + parts[1];
          } else {
            info.plaintiffName = cleanedName;
          }
          break;
        }
        if (info.plaintiffName) break;
      }
    }

    if (!info.defendantName) {
      const defendantPatterns = [
        /[【\[［]\s*(?:被|a)\s*告\s*[】\]］]\s*\n?\s*([\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff][^\n【\[［]{1,30})/,
        /被\s*告\s+(?!.*(?:訴\s*訟|代\s*理))([\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff][^\n（(原代訴】\]］]{1,30})/,
      ];
      for (const pattern of defendantPatterns) {
        const allMatches = [];
        const globalPattern = new RegExp(pattern.source, 'g');
        let match;
        while ((match = globalPattern.exec(cleanText)) !== null) {
          allMatches.push(match);
        }
        for (const match of allMatches) {
          let name = match[1].trim();
          const cleanedName = name.replace(/\s+/g, '');
          if (cleanedName.length <= 1) continue;
          name = name.replace(/\s*(外\s*\d+\s*名)\s*$/, (_, suffix) => ' ' + suffix.replace(/\s+/g, ''));
          const parts = name.split(/ (外\d+名)$/);
          if (parts.length > 1) {
            info.defendantName = parts[0].replace(/\s+/g, '') + ' ' + parts[1];
          } else {
            info.defendantName = cleanedName;
          }
          break;
        }
        if (info.defendantName) break;
      }
    }

    // --- 原告代理人弁護士名 ---
    const lawyerCandidates = [];

    function cleanLawyerName(rawName) {
      let name = rawName.trim();
      const cjkChars = name.match(/[\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]+/g);
      if (cjkChars) {
        name = cjkChars.join('');
      } else {
        return null;
      }
      name = name.replace(/宛て$/, '');
      name = name.replace(/[宛殿様御中方和]+$/, '');
      if (name.length < 2 || name.length > 6) return null;
      if (/^[法会事件番号裁判]/.test(name)) return null;
      return name;
    }

    const ownLawyerNames = config.lawyerNames || [];

    const formalPattern = /原告\s*(?:ら)?\s*(?:訴\s*訟)?\s*代理\s*人\s*弁護\s*士\s*([^\n]{2,20})/g;
    let lm;
    while ((lm = formalPattern.exec(cleanText)) !== null) {
      const name = cleanLawyerName(lm[1]);
      if (name) lawyerCandidates.push({ name, priority: 1 });
    }

    const senderPattern = /人\s*弁護\s*士\s*([^\n]{2,20})/g;
    while ((lm = senderPattern.exec(cleanText)) !== null) {
      const contextBefore = cleanText.substring(Math.max(0, lm.index - 30), lm.index);
      if (contextBefore.includes('被告')) continue;
      const name = cleanLawyerName(lm[1]);
      if (name) lawyerCandidates.push({ name, priority: 2 });
    }

    const atePattern = /弁護\s*士\s*([^\n]{2,15})\s*宛/g;
    while ((lm = atePattern.exec(cleanText)) !== null) {
      const name = cleanLawyerName(lm[1]);
      if (name) lawyerCandidates.push({ name, priority: 3 });
    }

    const generalPattern = /弁護\s*士\s*([^\n]{2,15})/g;
    while ((lm = generalPattern.exec(cleanText)) !== null) {
      const contextBefore = cleanText.substring(Math.max(0, lm.index - 50), lm.index);
      if (contextBefore.includes('被告')) continue;
      const name = cleanLawyerName(lm[1]);
      if (name && !ownLawyerNames.some(own => name.includes(own))) {
        lawyerCandidates.push({ name, priority: 4 });
      }
    }

    if (lawyerCandidates.length > 0) {
      const uniqueNames = [...new Set(lawyerCandidates.map(c => c.name))];
      const uniqueCandidates = uniqueNames.map(name => {
        const best = lawyerCandidates.filter(c => c.name === name)
          .sort((a, b) => a.priority - b.priority)[0];
        return best;
      });
      uniqueCandidates.sort((a, b) => {
        if (a.priority !== b.priority) return a.priority - b.priority;
        const aIdeal = (a.name.length >= 3 && a.name.length <= 4) ? 0 : 1;
        const bIdeal = (b.name.length >= 3 && b.name.length <= 4) ? 0 : 1;
        if (aIdeal !== bIdeal) return aIdeal - bIdeal;
        return b.name.length - a.name.length;
      });
      let bestName = uniqueCandidates[0].name;
      if (bestName.length === 5) {
        const shorter = uniqueCandidates.find(c => c.name.length <= 3 && bestName.startsWith(c.name));
        if (shorter) {
          bestName = bestName.substring(0, 4);
        }
      }
      info.plaintiffLawyer = bestName;
    }

    // --- 裁判所FAX番号（辞書引き）---
    if (info.courtName) {
      const courtBase = info.courtName
        .replace(/民事第[０-９\d]+部.*$/, '')
        .replace(/第[０-９\d]+[民刑]事部$/, '');
      info.courtFax = COURT_FAX_MAP[courtBase] || '';
    }

    // --- FAX番号の抽出 ---
    const ownFaxPatterns = config.faxNumbers || [];
    const courtFaxValues = Object.values(COURT_FAX_MAP);

    function normalizeFax(raw) {
      return raw
        .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
        .replace(/[－ー・]/g, '-');
    }

    const explicitCourtFaxMatch = cleanText.match(
      /裁\s*判\s*所[\s\S]{0,60}?[（(]\s*(?:FAX|ＦＡＸ|[Ff]ax)\s*([0-9０-９\-－ー・]+)\s*[）)]/
    );
    if (explicitCourtFaxMatch) {
      info.courtFaxFromPdf = normalizeFax(explicitCourtFaxMatch[1]);
    }

    const allExplicitFaxes = [];
    const explicitFaxRegex = /([\u4e00-\u9fff]{1,10})\s*[（(]\s*(?:FAX|ＦＡＸ|[Ff]ax)\s*([0-9０-９\-－ー・]+)\s*[）)]/g;
    let efm;
    while ((efm = explicitFaxRegex.exec(cleanText)) !== null) {
      allExplicitFaxes.push({ label: efm[1], fax: normalizeFax(efm[2]) });
    }
    for (const ef of allExplicitFaxes) {
      if (ef.label.includes('裁判') || ef.label.includes('裁判所')) continue;
      if (info.courtFaxFromPdf && ef.fax === info.courtFaxFromPdf) continue;
      const isOwn = ownFaxPatterns.some(p => ef.fax.includes(p));
      if (!isOwn && !info.plaintiffLawyerFax) {
        info.plaintiffLawyerFax = ef.fax;
      }
    }

    const faxRegex = /(?:FAX|ＦＡＸ|[Ff]ax)[：:\s]*([0-9０-９\-－ー・]+)/g;
    const allFaxEntries = [];
    let faxMatch;
    while ((faxMatch = faxRegex.exec(cleanText)) !== null) {
      const faxNum = normalizeFax(faxMatch[1]);
      allFaxEntries.push({ fax: faxNum, index: faxMatch.index });
    }

    for (const entry of allFaxEntries) {
      const isOwnFax = ownFaxPatterns.some(p => entry.fax.includes(p));
      if (isOwnFax) continue;
      if (info.courtFaxFromPdf && entry.fax.includes(info.courtFaxFromPdf)) continue;
      const isKnownCourtFax = courtFaxValues.some(cf => entry.fax.includes(cf));
      const textBefore = cleanText.substring(
        Math.max(0, entry.index - 200), entry.index
      );
      const isNearPlaintiffLawyer = /原告\s*(?:ら)?\s*訴\s*訟\s*代\s*理\s*人/.test(textBefore) ||
        (/弁護\s*士/.test(textBefore) && !textBefore.includes('被告'));
      const isNearDefendantLawyer = /被告\s*(?:ら)?\s*訴\s*訟\s*代\s*理\s*人/.test(textBefore);
      if (isNearDefendantLawyer) continue;
      if (isKnownCourtFax) {
        if (!info.courtFaxFromPdf) {
          info.courtFaxFromPdf = entry.fax;
        }
      } else if (isNearPlaintiffLawyer) {
        if (!info.plaintiffLawyerFax) {
          info.plaintiffLawyerFax = entry.fax;
        }
      } else {
        if (!info.plaintiffLawyerFax) {
          info.plaintiffLawyerFax = entry.fax;
        }
      }
    }

    if (info.courtFaxFromPdf) {
      info.courtFax = info.courtFaxFromPdf;
    }

    return info;
  }

  // =============================================
  // Section 4: 文書送付書 Word生成
  // =============================================

  function safeReplaceInXml(xml, oldText, newText) {
    const paraRegex = /(<w:p[\s>][\s\S]*?<\/w:p>)/g;
    return xml.replace(paraRegex, (paraXml) => {
      const wtRegex = /<w:t([^>]*)>([^<]*)<\/w:t>/g;
      const segments = [];
      let m;
      while ((m = wtRegex.exec(paraXml)) !== null) {
        segments.push({ fullMatch: m[0], attrs: m[1], text: m[2], index: m.index });
      }
      if (segments.length === 0) return paraXml;
      const joinedText = segments.map(s => s.text).join('');
      if (!joinedText.includes(oldText)) return paraXml;

      const matchStart = joinedText.indexOf(oldText);
      const matchEnd = matchStart + oldText.length;
      let cumulative = 0;
      for (const seg of segments) {
        seg.startPos = cumulative;
        seg.endPos = cumulative + seg.text.length;
        cumulative += seg.text.length;
      }
      const affectedSegs = segments.filter(
        seg => seg.endPos > matchStart && seg.startPos < matchEnd
      );
      if (affectedSegs.length === 0) return paraXml;

      if (affectedSegs.length === 1) {
        const seg = affectedSegs[0];
        const localStart = matchStart - seg.startPos;
        const localEnd = matchEnd - seg.startPos;
        const newSegText = seg.text.substring(0, localStart) + newText + seg.text.substring(localEnd);
        const hasPreserve = seg.attrs.includes('xml:space="preserve"');
        const newAttrs = hasPreserve ? seg.attrs : ' xml:space="preserve"';
        const newWt = `<w:t${newAttrs}>${newSegText}</w:t>`;
        return paraXml.substring(0, seg.index) + newWt +
          paraXml.substring(seg.index + seg.fullMatch.length);
      }

      let segIdx = 0;
      const result = paraXml.replace(/<w:t([^>]*)>([^<]*)<\/w:t>/g, (match, attrs, text) => {
        const seg = segments[segIdx];
        segIdx++;
        if (!affectedSegs.includes(seg)) return match;
        const isFirst = seg === affectedSegs[0];
        const isLast = seg === affectedSegs[affectedSegs.length - 1];
        const hasPreserve = attrs.includes('xml:space="preserve"');
        const newAttrs = hasPreserve ? attrs : ' xml:space="preserve"';
        if (isFirst && isLast) {
          const localStart = matchStart - seg.startPos;
          const localEnd = matchEnd - seg.startPos;
          return `<w:t${newAttrs}>${text.substring(0, localStart)}${newText}${text.substring(localEnd)}</w:t>`;
        } else if (isFirst) {
          const localStart = matchStart - seg.startPos;
          return `<w:t${newAttrs}>${text.substring(0, localStart)}${newText}</w:t>`;
        } else if (isLast) {
          const localEnd = matchEnd - seg.startPos;
          const remaining = text.substring(localEnd);
          if (remaining.length > 0) {
            return `<w:t${newAttrs}>${remaining}</w:t>`;
          } else {
            return `<w:t${attrs}></w:t>`;
          }
        } else {
          return `<w:t${attrs}></w:t>`;
        }
      });
      return result;
    });
  }

  function applyInfoToTemplate(docXml, info, documentTitle) {
    const today = getTodayReiwa();
    if (info.courtName) {
      const ORIG_COURT = '神戸地方裁判所尼崎支部第２民事部';
      const courtDiff = ORIG_COURT.length - info.courtName.length;
      const courtPad = courtDiff > 0 ? '\u3000'.repeat(courtDiff) : '';
      docXml = safeReplaceInXml(docXml, ORIG_COURT, info.courtName + courtPad);
    }
    if (info.courtFax) {
      docXml = safeReplaceInXml(docXml, '06-6438-1710', info.courtFax);
      const fullWidthFax = toFullWidthNumber(info.courtFax).replace(/-/g, '\uFF0D');
      docXml = safeReplaceInXml(docXml, '\uFF10\uFF16\u2015\uFF16\uFF14\uFF13\uFF18\uFF0D\uFF11\uFF17\uFF11\uFF10', fullWidthFax);
    }
    if (info.plaintiffLawyer) {
      const ORIG_LAWYER = '四方久寛';
      const lawyerDiff = ORIG_LAWYER.length - info.plaintiffLawyer.length;
      const lawyerPad = lawyerDiff > 0 ? '\u3000'.repeat(lawyerDiff) : '';
      docXml = safeReplaceInXml(docXml, ORIG_LAWYER, info.plaintiffLawyer + lawyerPad);
    }
    if (info.plaintiffLawyerFax) {
      docXml = safeReplaceInXml(docXml, '06-4708-3638', info.plaintiffLawyerFax);
    }
    docXml = safeReplaceInXml(docXml, '令和6年11月7日', `令和${today.year}年${today.month}月${today.day}日`);
    docXml = safeReplaceInXml(docXml, '令和6年9月', `令和${today.year}年${today.month}月`);
    if (info.caseNumber) {
      const fullWidthCaseNumber = toFullWidthNumber(info.caseNumber);
      docXml = safeReplaceInXml(docXml, '令和３年（ワ）第８００号', fullWidthCaseNumber);
    }
    if (info.caseName) {
      docXml = safeReplaceInXml(docXml, '損害賠償請求事件', info.caseName);
    }
    if (info.plaintiffName) {
      docXml = safeReplaceInXml(docXml, '木村治紀', info.plaintiffName);
    }
    if (info.defendantName) {
      docXml = safeReplaceInXml(docXml, '独立行政法人国立病院機構', info.defendantName);
    }
    docXml = safeReplaceInXml(docXml, '被告第９準備書面', documentTitle);
    return docXml;
  }

  function getDocumentTitleFromFilename(fileName) {
    let baseName = fileName.replace(/\.pdf$/i, '');
    baseName = baseName.replace(/^【[^】]+】\s*/, '');
    baseName = baseName.replace(/^[\u4e00-\u9fff]+事案[\s\u3000]+/, '');
    return baseName;
  }

  async function uploadAndExtractBrowser(file, onProgress) {
    const fileName = file.name.toLowerCase();
    const isDocx = fileName.endsWith('.docx') || fileName.endsWith('.doc');
    console.log('[つくる君] 解析開始:', file.name, file.size, 'bytes', isDocx ? '(Word)' : '(PDF)');

    const arrayBuffer = await file.arrayBuffer();
    let extractedText;

    if (isDocx) {
      onProgress && onProgress('Wordファイルを読み込み中...');
      extractedText = await extractTextFromDocx(arrayBuffer, onProgress);
    } else {
      onProgress && onProgress('PDFを読み込み中...');
      extractedText = await extractTextBrowser(arrayBuffer, onProgress);
    }

    console.log('[つくる君] 抽出テキスト:', extractedText.length, '文字');
    onProgress && onProgress('情報を抽出中...');
    const info = extractInfoFromText(extractedText);
    console.log('[つくる君] 抽出結果:', JSON.stringify(info, null, 2));
    const documentTitle = getDocumentTitleFromFilename(file.name);
    return { info, documentTitle, originalName: file.name };
  }

  async function generateDocumentBrowser(info, documentTitle, onProgress) {
    onProgress && onProgress('テンプレートを読み込み中...');
    const templateData = await loadTemplate(
      'template/文書送付書.doc.docx',
      'https://cdn.jsdelivr.net/gh/Rachmaninovpiano/bunsho-sofusho-juryosho-auto@master/docs/template/%E6%96%87%E6%9B%B8%E9%80%81%E4%BB%98%E6%9B%B8.doc.docx',
      'sofusho_template'
    );
    onProgress && onProgress('テンプレートにデータを差し込み中...');
    const zip = await JSZip.loadAsync(templateData);
    let docXml = await zip.file('word/document.xml').async('string');
    docXml = applyInfoToTemplate(docXml, info, documentTitle);
    zip.file('word/document.xml', docXml);
    onProgress && onProgress('Wordファイルを生成中...');
    const outputBlob = await zip.generateAsync({ type: 'blob' });
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    const outputFileName = `文書送付書_${documentTitle}_${timestamp}.docx`;
    return { blob: outputBlob, fileName: outputFileName };
  }

  // =============================================
  // Section 5: 受領書 OCR・PDF生成
  // =============================================

  async function runOcrBrowser(pdfArrayBuffer, pageNum, onProgress) {
    onProgress && onProgress(`ページ${pageNum}を描画中...`);
    const pdfDoc = await pdfjsLib.getDocument({
      data: pdfArrayBuffer,
      cMapUrl: CMAP_URL,
      cMapPacked: true,
    }).promise;
    const page = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: 400 / 72 });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport }).promise;
    const imgWidth = canvas.width;
    const imgHeight = canvas.height;
    onProgress && onProgress(`ページ${pageNum}をOCR中...`);
    const worker = await Tesseract.createWorker('jpn', 1, {
      logger: m => {
        if (m.status === 'recognizing text' && onProgress) {
          onProgress(`OCR処理中... ${Math.round((m.progress || 0) * 100)}%`);
        }
      }
    });
    const { data } = await worker.recognize(canvas);
    await worker.terminate();
    const words = [];
    if (data && data.words) {
      for (const w of data.words) {
        const text = w.text.trim();
        if (!text) continue;
        const bbox = w.bbox;
        words.push({ x1: bbox.x0, y1: bbox.y0, x2: bbox.x1, y2: bbox.y1, text });
      }
    }
    console.log(`  [OCR] ページ${pageNum}: ${words.length}語検出`);
    return { words, imgWidth, imgHeight };
  }

  function px2pdf(px, py, imgW, imgH, pgW, pgH) {
    return {
      x: px * pgW / imgW,
      y: pgH - (py * pgH / imgH),
    };
  }

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
      if (score > bestScore) {
        bestScore = score;
        bestPageNum = p;
        bestOcr = ocr;
      }
      if (score >= 50) {
        return { pageNum: p, ocr };
      }
    }
    return { pageNum: bestPageNum, ocr: bestOcr };
  }

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

  async function generateReceiptBrowser(file, options, onProgress) {
    onProgress && onProgress('PDFを読み込み中...');
    const today = new Date();
    const reiwaYear = today.getFullYear() - 2018;
    const defaultDate = `令和${reiwaYear}年${today.getMonth() + 1}月${today.getDate()}日`;
    const config = getConfig();
    const receiptDate = (options && options.receiptDate) || defaultDate;
    const signerTitle = (options && options.signerTitle) || '被告訴訟代理人';
    const signerName  = (options && options.signerName)  || config.signerName || '山田太郎';

    const pdfArrayBuffer = await file.arrayBuffer();
    const pdfDoc = await PDFLib.PDFDocument.load(pdfArrayBuffer);
    pdfDoc.registerFontkit(fontkit);
    const totalPages = pdfDoc.getPageCount();

    onProgress && onProgress('受領書ページを探しています...');
    const { pageNum: receiptPageNum, ocr } = await findReceiptPage(pdfArrayBuffer, totalPages, onProgress);
    const receiptPageIndex = receiptPageNum - 1;
    const words = ocr.words;
    const imgWidth = ocr.imgWidth;
    const imgHeight = ocr.imgHeight;
    const page = pdfDoc.getPage(receiptPageIndex);
    const { width: pgW, height: pgH } = page.getSize();

    onProgress && onProgress('フォントを読み込み中...');
    const fontBytes = await loadJapaneseFont();
    const font = await pdfDoc.embedFont(fontBytes, { subset: false });

    const allChars = `行先生${receiptDate}${signerTitle}　${signerName}㊞`;
    try { font.encodeText(allChars); } catch (e) { /* ignore */ }

    onProgress && onProgress('書き込み位置を検出中...');
    const pos = detectPositions(words, imgWidth, imgHeight, pgW, pgH);
    const fs_ = 10.5;
    const { rgb } = PDFLib;

    // 「行」→ 二重打消し線 + 「先生」
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

    // 受領日記入
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

    // 署名記入
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

      // 印鑑画像
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
    return { blob, fileName: outFileName };
  }

  // =============================================
  // Section 7: 証拠番号つくる君 PDF生成 + 証拠説明書
  // =============================================

  /**
   * 証拠番号ラベル構築
   */
  function buildEvidenceLabel(party, num, subNum, useDai) {
    const fullNum = toFullWidthNumber(String(num));
    let label = party;
    if (useDai) label += '第';
    label += fullNum + '号証';
    if (subNum) {
      label += 'の' + toFullWidthNumber(String(subNum));
    }
    return label;
  }

  /**
   * mints形式ファイル名生成
   * 例: 甲001 売買契約書.pdf, 甲001-1~5 陳述書.pdf
   */
  function buildMintsFileName(party, num, subNum, title, subRange) {
    const numStr = String(num).padStart(3, '0');
    let name = party + numStr;
    if (subNum) {
      name += '-' + subNum;
    } else if (subRange) {
      name += '-' + subRange;
    }
    if (title) name += ' ' + title;
    return name + '.pdf';
  }

  /** XML特殊文字エスケープ */
  function escXml(str) {
    return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /**
   * スタンプ色取得
   */
  function getStampColor(colorName) {
    const { rgb } = PDFLib;
    switch (colorName) {
      case 'red': return rgb(0.86, 0.15, 0.15);
      case 'blue': return rgb(0.1, 0.2, 0.7);
      case 'black': return rgb(0, 0, 0);
      default: return rgb(0.86, 0.15, 0.15);
    }
  }

  /**
   * PDFの指定ページに証拠番号をスタンプ（太字）。
   * 書誌情報（標目）はPDFには打ち込まない（証拠説明書のみ）。
   */
  async function generateEvidenceBrowser(file, opts) {
    const {
      evidenceLabel, evidenceTitle, allPages, onProgress,
      stampSize = 20, stampColor = 'red',
      stampBg = true, stampBorder = false,
      addPageNum = false,
      customX = 0.85, customY = 0.03
    } = opts;

    onProgress && onProgress('PDFを読み込み中...');
    const pdfArrayBuffer = await file.arrayBuffer();
    const pdfDoc = await PDFLib.PDFDocument.load(pdfArrayBuffer);
    pdfDoc.registerFontkit(fontkit);

    onProgress && onProgress('フォントを読み込み中...');
    const fontBytes = await loadJapaneseFont();
    const font = await pdfDoc.embedFont(fontBytes, { subset: false });

    const labelFontSize = parseInt(stampSize, 10) || 20;
    const color = getStampColor(stampColor);
    const { rgb } = PDFLib;

    const pageCount = pdfDoc.getPageCount();
    const pagesToStamp = allPages
      ? Array.from({ length: pageCount }, (_, i) => i)
      : [0];

    // 太字化: テキストを微小オフセットで複数回描画
    function drawBoldText(page, text, x, y, size, font, color) {
      var offsets = [
        [0, 0], [0.4, 0], [-0.4, 0], [0, 0.4], [0, -0.4],
        [0.2, 0.2], [-0.2, 0.2], [0.2, -0.2], [-0.2, -0.2]
      ];
      for (var k = 0; k < offsets.length; k++) {
        page.drawText(text, {
          x: x + offsets[k][0], y: y + offsets[k][1],
          size: size, font: font, color: color,
        });
      }
    }

    onProgress && onProgress('証拠番号を書き込み中...');
    for (const pageIndex of pagesToStamp) {
      const page = pdfDoc.getPage(pageIndex);
      const { width: pgW, height: pgH } = page.getSize();

      // ラベルテキストのサイズ計算（証拠番号のみ、書誌情報は入れない）
      const labelWidth = font.widthOfTextAtSize(evidenceLabel, labelFontSize);
      const boxPadH = 8;
      const boxPadV = 6;
      const boxWidth = labelWidth + boxPadH * 2;
      const boxHeight = labelFontSize + boxPadV * 2;

      // 位置計算
      let boxX = customX * pgW - boxWidth / 2;
      let boxY = pgH - (customY * pgH) - boxHeight;
      boxX = Math.max(2, Math.min(pgW - boxWidth - 2, boxX));
      boxY = Math.max(2, Math.min(pgH - boxHeight - 2, boxY));

      // 背景描画
      if (stampBg) {
        page.drawRectangle({
          x: boxX, y: boxY,
          width: boxWidth, height: boxHeight,
          color: rgb(1, 1, 1),
          opacity: 0.92,
          borderColor: stampBorder ? color : undefined,
          borderWidth: stampBorder ? 1 : 0,
        });
      } else if (stampBorder) {
        page.drawRectangle({
          x: boxX, y: boxY,
          width: boxWidth, height: boxHeight,
          borderColor: color,
          borderWidth: 1,
        });
      }

      // 証拠番号ラベル（太字描画）
      const labelX = boxX + (boxWidth - labelWidth) / 2;
      const labelY = boxY + boxPadV;
      drawBoldText(page, evidenceLabel, labelX, labelY, labelFontSize, font, color);
    }

    // ページ番号付与
    if (addPageNum) {
      const pageNumSize = 10;
      for (let i = 0; i < pageCount; i++) {
        const page = pdfDoc.getPage(i);
        const { width: pgW } = page.getSize();
        const pageNumText = '- ' + (i + 1) + ' -';
        const numWidth = font.widthOfTextAtSize(pageNumText, pageNumSize);
        page.drawText(pageNumText, {
          x: (pgW - numWidth) / 2,
          y: 24,
          size: pageNumSize, font,
          color: rgb(0.3, 0.3, 0.3),
        });
      }
    }

    const savedBytes = await pdfDoc.save();
    const blob = new Blob([savedBytes], { type: 'application/pdf' });
    // ファイル名: 甲○号証.pdf or 甲○号証（△△）.pdf
    let fileName = evidenceLabel;
    if (evidenceTitle && evidenceTitle.trim()) {
      fileName += '\uff08' + evidenceTitle.trim() + '\uff09';
    }
    fileName += '.pdf';
    return { blob, fileName, pageCount };
  }

  /**
   * 複数PDFを結合（枝番結合用）
   */
  async function mergePdfs(files, onProgress) {
    onProgress && onProgress('PDFを結合中...');
    const mergedPdf = await PDFLib.PDFDocument.create();
    for (let i = 0; i < files.length; i++) {
      onProgress && onProgress('結合中 (' + (i + 1) + '/' + files.length + ')...');
      const ab = await files[i].arrayBuffer();
      const srcPdf = await PDFLib.PDFDocument.load(ab);
      const pages = await mergedPdf.copyPages(srcPdf, srcPdf.getPageIndices());
      pages.forEach(function(p) { mergedPdf.addPage(p); });
    }
    const mergedBytes = await mergedPdf.save();
    return new File(
      [mergedBytes],
      '結合_' + files[0].name,
      { type: 'application/pdf' }
    );
  }


  async function generateEvidenceSheetDocx(entries, options) {
    const opts = options || {};
    const party = opts.party || '甲';
    const today = getTodayReiwa();
    const dateStr = `令和${today.year}年${today.month}月${today.day}日`;
    const titleLabel = `証拠説明書（${party}号証）`;

    // テーブルヘッダー行
    const headerRow = [
      '<w:tr>',
      '<w:tc><w:tcPr><w:tcW w:w="1100" w:type="dxa"/></w:tcPr>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:b/><w:sz w:val="22"/></w:rPr>',
      '<w:t>号証</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:b/><w:sz w:val="22"/></w:rPr>',
      '<w:t xml:space="preserve">標　　　目</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr>',
      '<w:t>（原本・写しの別）</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:b/><w:sz w:val="22"/></w:rPr>',
      '<w:t>作　成</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:b/><w:sz w:val="22"/></w:rPr>',
      '<w:t>年月日</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:b/><w:sz w:val="22"/></w:rPr>',
      '<w:t>作成者</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="3200" w:type="dxa"/></w:tcPr>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:b/><w:sz w:val="22"/></w:rPr>',
      '<w:t>立証趣旨</w:t></w:r></w:p></w:tc>',
      '</w:tr>',
    ].join('');

    // データ行
    const dataRows = entries.map(e => [
      '<w:tr><w:trPr><w:trHeight w:val="500"/></w:trPr>',
      '<w:tc><w:tcPr><w:tcW w:w="1100" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr>',
      '<w:t>' + escXml(e.label) + '</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>',
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr>',
      '<w:t>' + escXml(e.title || '') + '</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="800" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr>',
      '<w:t>' + escXml(e.originalOrCopy || '') + '</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr>',
      '<w:t>' + escXml(e.createdDate || '') + '</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr>',
      '<w:t>' + escXml(e.author || '') + '</w:t></w:r></w:p></w:tc>',
      '<w:tc><w:tcPr><w:tcW w:w="3200" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>',
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr>',
      '<w:t>' + escXml(e.purpose || '') + '</w:t></w:r></w:p></w:tc>',
      '</w:tr>',
    ].join('')).join('\n');

    // document.xml
    const documentXml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
      ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">',
      '<w:body>',
      // タイトル: 証拠説明書（甲号証）
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>',
      '<w:t>' + escXml(titleLabel) + '</w:t></w:r></w:p>',
      // 空行
      '<w:p/>',
      // 日付（右寄せ）
      '<w:p><w:pPr><w:jc w:val="right"/></w:pPr>',
      '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="24"/></w:rPr>',
      '<w:t>' + escXml(dateStr) + '</w:t></w:r></w:p>',
      // 空行
      '<w:p/>',
      // テーブル
      '<w:tbl>',
      '<w:tblPr>',
      '<w:tblW w:w="9500" w:type="dxa"/>',
      '<w:tblBorders>',
      '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>',
      '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>',
      '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>',
      '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>',
      '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>',
      '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>',
      '</w:tblBorders>',
      '<w:tblLayout w:type="fixed"/>',
      '</w:tblPr>',
      '<w:tblGrid>',
      '<w:gridCol w:w="1100"/><w:gridCol w:w="2000"/><w:gridCol w:w="800"/>',
      '<w:gridCol w:w="1200"/><w:gridCol w:w="1200"/><w:gridCol w:w="3200"/>',
      '</w:tblGrid>',
      headerRow,
      dataRows,
      '</w:tbl>',
      '<w:p/>',
      // ページ設定（A4縦）
      '<w:sectPr>',
      '<w:pgSz w:w="11906" w:h="16838"/>',
      '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>',
      '</w:sectPr>',
      '</w:body>',
      '</w:document>',
    ].join('\n');

    // [Content_Types].xml
    const contentTypes = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
      '</Types>',
    ].join('\n');

    // _rels/.rels
    const relsXml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
      '</Relationships>',
    ].join('\n');

    // word/_rels/document.xml.rels
    const docRels = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '</Relationships>',
    ].join('\n');

    // JSZipでdocxパッケージ
    const zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.file('_rels/.rels', relsXml);
    zip.file('word/document.xml', documentXml);
    zip.file('word/_rels/document.xml.rels', docRels);

    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });

    return blob;
  }

  // =============================================
  // Section 6: UIコントローラー
  // =============================================

  let currentState = 'upload';
  let currentMode = 'sofusho';
  let receiptUploadFiles = [];
  let evidenceUploadFiles = [];

  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  const states = {
    upload: $('#state-upload'),
    processing: $('#state-processing'),
    confirm: $('#state-confirm'),
    'receipt-confirm': $('#state-receipt-confirm'),
    'evidence-confirm': $('#state-evidence-confirm'),
    complete: $('#state-complete'),
  };

  const dropZone = $('#dropZone');
  const fileInput = $('#fileInput');
  const processingTitle = $('#processingTitle');
  const processingMessage = $('#processingMessage');
  const sourceFileName = $('#sourceFileName');
  const caseNumberWarning = $('#caseNumberWarning');
  const outputFileName = $('#outputFileName');
  const btnBack = $('#btnBack');
  const btnGenerate = $('#btnGenerate');
  const btnDownload = $('#btnDownload');
  const btnNewFile = $('#btnNewFile');
  const errorToast = $('#errorToast');
  const errorMessage = $('#errorMessage');
  const errorClose = $('#errorClose');
  const extractStatus = $('#extractStatus');
  const confettiContainer = $('#confetti');
  const dragOverlay = $('#dragOverlay');
  const singleDownloadArea = $('#singleDownloadArea');
  const multiDownloadArea = $('#multiDownloadArea');
  const downloadList = $('#downloadList');
  const receiptFileListWrap = $('#receiptFileListWrap');
  const receiptFileList = $('#receiptFileList');
  const receiptFileCountBadge = $('#receiptFileCountBadge');
  const procStep1 = $('#procStep1');
  const procStep2 = $('#procStep2');
  const procStep3 = $('#procStep3');

  const fields = {
    courtName: $('#courtName'),
    courtFax: $('#courtFax'),
    caseNumber: $('#caseNumber'),
    caseName: $('#caseName'),
    plaintiffName: $('#plaintiffName'),
    defendantName: $('#defendantName'),
    plaintiffLawyer: $('#plaintiffLawyer'),
    plaintiffLawyerFax: $('#plaintiffLawyerFax'),
    documentTitle: $('#documentTitle'),
  };

  const modeSofushoBtn = $('#modeSofusho');
  const modeReceiptBtn = $('#modeReceipt');
  const modeEvidenceBtn = $('#modeEvidence');
  const appTitle = $('#appTitle');
  const logoIcon = $('#logoIcon');
  const completeTitle = $('#completeTitle');
  const downloadLabel = $('#downloadLabel');
  const uploadHeading = $('.upload-heading');
  const uploadDesc = $('.upload-desc');
  const settingsBtn = $('#settingsBtn');
  const settingsModal = $('#settingsModal');
  const settingsClose = $('#settingsClose');
  const settingsSave = $('#settingsSave');
  const receiptSourceFileName = $('#receiptSourceFileName');
  const receiptSignerTitle = $('#receiptSignerTitle');
  const receiptSignerName = $('#receiptSignerName');
  const receiptDateInput = $('#receiptDate');
  const btnReceiptBack = $('#btnReceiptBack');
  const btnReceiptGenerate = $('#btnReceiptGenerate');

  // --- ステート切り替え ---
  function setState(newState) {
    states[currentState].classList.remove('active');
    const stateOrder = ['upload', 'processing', 'confirm', 'complete'];
    const mapState = (newState === 'receipt-confirm' || newState === 'evidence-confirm') ? 'confirm' : newState;
    const newIndex = stateOrder.indexOf(mapState);
    const stepElements = $$('.step');
    const connectorFills = $$('.step-connector-fill');
    stepElements.forEach((step, i) => {
      step.classList.remove('active', 'completed');
      if (i < newIndex) step.classList.add('completed');
      if (i === newIndex) step.classList.add('active');
    });
    connectorFills.forEach((fill, i) => {
      fill.style.width = i < newIndex ? '100%' : '0%';
    });
    currentState = newState;
    states[currentState].classList.add('active');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // --- エラー表示 ---
  let errorTimer = null;
  function showError(msg) {
    errorMessage.textContent = msg;
    errorToast.hidden = false;
    if (errorTimer) clearTimeout(errorTimer);
    errorTimer = setTimeout(() => { errorToast.hidden = true; }, 15000);
  }
  errorClose.addEventListener('click', () => {
    errorToast.hidden = true;
    if (errorTimer) clearTimeout(errorTimer);
  });

  // --- 処理ステップアニメーション ---
  let processingTimers = [];
  function resetProcessingSteps() {
    processingTimers.forEach(t => clearTimeout(t));
    processingTimers = [];
    [procStep1, procStep2, procStep3].forEach(step => {
      if (step) step.classList.remove('active', 'done');
    });
  }
  function startProcessingSteps(mode) {
    resetProcessingSteps();
    if (mode === 'upload') {
      if (procStep1) procStep1.classList.add('active');
    } else if (mode === 'generate') {
      [procStep1, procStep2, procStep3].forEach(step => {
        if (step) step.classList.add('done');
      });
    }
  }

  // --- 進捗更新 ---
  function updateProgress(msg) {
    if (processingMessage) processingMessage.textContent = msg;
    if (msg.includes('OCR') || msg.includes('文字認識')) {
      if (procStep1) { procStep1.classList.remove('active'); procStep1.classList.add('done'); }
      if (procStep2) procStep2.classList.add('active');
    }
    if (msg.includes('抽出') || msg.includes('検出')) {
      if (procStep1) { procStep1.classList.remove('active'); procStep1.classList.add('done'); }
      if (procStep2) { procStep2.classList.remove('active'); procStep2.classList.add('done'); }
      if (procStep3) procStep3.classList.add('active');
    }
  }

  // --- 紙吹雪 ---
  function launchConfetti() {
    if (!confettiContainer) return;
    confettiContainer.innerHTML = '';
    const colors = ['#4f46e5', '#06b6d4', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899'];
    for (let i = 0; i < 50; i++) {
      const piece = document.createElement('div');
      piece.style.cssText = `
        position: absolute;
        width: ${Math.random() * 8 + 4}px;
        height: ${Math.random() * 8 + 4}px;
        background: ${colors[Math.floor(Math.random() * colors.length)]};
        border-radius: ${Math.random() > 0.5 ? '50%' : '2px'};
        left: ${Math.random() * 100}%;
        top: -10px;
        opacity: 1;
        transform: rotate(${Math.random() * 360}deg);
      `;
      const duration = Math.random() * 1500 + 1000;
      const delay = Math.random() * 500;
      piece.animate([
        { transform: 'translateY(0) rotate(0deg) scale(1)', opacity: 1 },
        {
          transform: `translateY(${Math.random() * 200 + 100}px) translateX(${(Math.random() - 0.5) * 150}px) rotate(${Math.random() * 720}deg) scale(0)`,
          opacity: 0,
        }
      ], { duration, delay, easing: 'cubic-bezier(0.25, 0.46, 0.45, 0.94)', fill: 'forwards' });
      confettiContainer.appendChild(piece);
    }
    setTimeout(() => { if (confettiContainer) confettiContainer.innerHTML = ''; }, 3000);
  }

  // --- ファイルアップロード ---
  function handleFiles(fileList) {
    if (!fileList || fileList.length === 0) return;
    const validExts = ['.pdf', '.docx', '.doc'];
    const allFiles = Array.from(fileList).filter(f => {
      const name = f.name.toLowerCase();
      return validExts.some(ext => name.endsWith(ext));
    });
    // 受領書・証拠モードはPDFのみ
    if (currentMode === 'receipt' || currentMode === 'evidence') {
      const pdfs = allFiles.filter(f => f.name.toLowerCase().endsWith('.pdf'));
      if (pdfs.length === 0) {
        showError('PDFファイルを選択してください。');
        return;
      }
      if (currentMode === 'receipt') {
        prepareReceiptFiles(pdfs);
      } else {
        prepareEvidenceFiles(pdfs);
      }
      return;
    }
    // 文書送付書モード: PDF + Word
    if (allFiles.length === 0) {
      showError('PDF または Word (.docx) ファイルを選択してください。');
      return;
    }
    const oversized = allFiles.filter(f => f.size > 50 * 1024 * 1024);
    if (oversized.length > 0) {
      showError(`ファイルサイズが大きすぎます（上限: 50MB）: ${oversized.map(f=>f.name).join(', ')}`);
      return;
    }
    uploadFiles(allFiles);
  }

  // --- ドラッグ＆ドロップ ---
  document.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
  }, true);

  let dragCounter = 0;
  document.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragCounter++;
    if (currentState === 'upload' && dragCounter === 1) {
      dropZone.classList.add('dragover');
      if (dragOverlay) dragOverlay.classList.add('visible');
    }
  }, true);

  document.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dragCounter--;
    if (dragCounter <= 0) {
      dragCounter = 0;
      dropZone.classList.remove('dragover');
      if (dragOverlay) dragOverlay.classList.remove('visible');
    }
  }, true);

  document.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter = 0;
    dropZone.classList.remove('dragover');
    if (dragOverlay) dragOverlay.classList.remove('visible');
    if (currentState === 'upload' && e.dataTransfer && e.dataTransfer.files.length > 0) {
      handleFiles(e.dataTransfer.files);
    }
  }, true);

  dropZone.addEventListener('click', (e) => {
    if (e.target.closest('.upload-btn-label')) return;
    if (e.target.closest('label')) return;
    fileInput.click();
  });
  fileInput.addEventListener('change', () => {
    handleFiles(fileInput.files);
    fileInput.value = '';
  });

  // --- 文書送付書モード: 複数PDFアップロード → 情報統合 ---
  async function uploadFiles(pdfs) {
    setState('processing');
    processingTitle.textContent = 'ファイルを解析中...';
    processingMessage.textContent = 'ファイルを読み込んでいます...';
    startProcessingSteps('upload');

    try {
      const total = pdfs.length;
      const allResults = [];

      for (let i = 0; i < total; i++) {
        if (total > 1) {
          processingTitle.textContent = `ファイルを解析中... (${i + 1}/${total})`;
        }
        const result = await uploadAndExtractBrowser(pdfs[i], updateProgress);
        allResults.push(result);
      }

      // 情報をマージ（最初に見つかった非空値を採用）
      let mergedInfo, documentTitle, displayName;

      if (allResults.length === 1) {
        mergedInfo = allResults[0].info;
        documentTitle = allResults[0].documentTitle;
        displayName = allResults[0].originalName;
      } else {
        mergedInfo = {};
        const infoFields = ['courtName', 'courtFax', 'courtFaxFromPdf', 'caseNumber',
                            'caseName', 'plaintiffName', 'defendantName',
                            'plaintiffLawyer', 'plaintiffLawyerFax'];
        for (const field of infoFields) {
          for (const r of allResults) {
            if (r.info[field]) {
              mergedInfo[field] = r.info[field];
              break;
            }
          }
        }
        // caseNumberGuessed: 確信のある事件番号があればfalse
        const hasConfident = allResults.some(r => r.info.caseNumber && !r.info.caseNumberGuessed);
        mergedInfo.caseNumberGuessed = mergedInfo.caseNumber ? !hasConfident : false;
        documentTitle = allResults[0].documentTitle;
        displayName = allResults.map(r => r.originalName).join(' + ');
      }

      // ステップ完了
      resetProcessingSteps();
      [procStep1, procStep2, procStep3].forEach(step => {
        if (step) step.classList.add('done');
      });

      // ファイル数バッジ
      const badge = $('#sofushoFileCountBadge');
      if (badge) {
        if (total > 1) {
          badge.textContent = `${total}件統合`;
          badge.style.display = '';
        } else {
          badge.style.display = 'none';
        }
      }

      populateForm(mergedInfo, documentTitle, displayName);
      await new Promise(resolve => setTimeout(resolve, 500));
      setState('confirm');

    } catch (err) {
      console.error('[つくる君] PDF解析エラー:', err);
      resetProcessingSteps();
      showError(err.message || 'PDF解析中にエラーが発生しました。');
      setState('upload');
    }
  }

  // --- フォームにデータ流し込み ---
  function populateForm(info, docTitle, originalName) {
    fields.courtName.value = info.courtName || '';
    fields.courtFax.value = info.courtFax || '';
    fields.caseNumber.value = info.caseNumber || '';
    fields.caseName.value = info.caseName || '';
    fields.plaintiffName.value = info.plaintiffName || '';
    fields.defendantName.value = info.defendantName || '';
    fields.plaintiffLawyer.value = info.plaintiffLawyer || '';
    fields.plaintiffLawyerFax.value = info.plaintiffLawyerFax || '';
    fields.documentTitle.value = docTitle || '';
    sourceFileName.textContent = originalName;
    caseNumberWarning.hidden = !info.caseNumberGuessed;

    if (extractStatus) {
      const emptyCount = Object.entries(fields).filter(([key, input]) => !input.value).length;
      if (emptyCount === 0) {
        extractStatus.textContent = '全項目抽出完了';
        extractStatus.className = 'status-badge status-success';
      } else if (emptyCount <= 2) {
        extractStatus.textContent = `${emptyCount}件の未検出項目あり`;
        extractStatus.className = 'status-badge status-warning';
      } else {
        extractStatus.textContent = `${emptyCount}件の未検出項目あり`;
        extractStatus.className = 'status-badge status-error-badge';
      }
    }

    Object.entries(fields).forEach(([key, input]) => {
      if (!input.value) {
        input.classList.add('field-empty');
      } else {
        input.classList.remove('field-empty');
      }
    });
    Object.values(fields).forEach(input => {
      input.addEventListener('focus', () => {
        input.classList.remove('field-empty');
      });
    });
    // プレビュー初期更新
    updateSofushoPreview();
  }

  // --- 文書送付書プレビュー更新 ---
  function updateSofushoPreview() {
    var today = getTodayReiwa();
    var dateStr = '令和' + today.year + '年' + today.month + '月' + today.day + '日';
    var el = function(id) { return document.getElementById(id); };
    var pDate = el('sofushoPreviewDate');
    if (pDate) pDate.textContent = dateStr;
    var pCourt = el('sofushoPreviewCourt');
    if (pCourt) pCourt.textContent = fields.courtName.value.trim();
    var pCourtFax = el('sofushoPreviewCourtFax');
    if (pCourtFax) pCourtFax.textContent = fields.courtFax.value.trim();
    var pCaseNum = el('sofushoPreviewCaseNumber');
    if (pCaseNum) pCaseNum.textContent = fields.caseNumber.value.trim();
    var pCaseName = el('sofushoPreviewCaseName');
    if (pCaseName) pCaseName.textContent = fields.caseName.value.trim();
    var pPlaintiff = el('sofushoPreviewPlaintiff');
    if (pPlaintiff) pPlaintiff.textContent = fields.plaintiffName.value.trim();
    var pDefendant = el('sofushoPreviewDefendant');
    if (pDefendant) pDefendant.textContent = fields.defendantName.value.trim();
    var pDocTitle = el('sofushoPreviewDocTitle');
    if (pDocTitle) pDocTitle.textContent = fields.documentTitle.value.trim();
    var pLawyer = el('sofushoPreviewLawyer');
    if (pLawyer) pLawyer.textContent = fields.plaintiffLawyer.value.trim() ? '原告訴訟代理人弁護士　' + fields.plaintiffLawyer.value.trim() : '';
    var pLawyerFax = el('sofushoPreviewLawyerFax');
    if (pLawyerFax) pLawyerFax.textContent = fields.plaintiffLawyerFax.value.trim() ? 'FAX: ' + fields.plaintiffLawyerFax.value.trim() : '';
  }

  // 文書送付書の全入力フィールドにリアルタイムプレビュー更新をバインド
  Object.values(fields).forEach(function(input) {
    input.addEventListener('input', updateSofushoPreview);
  });

  // --- 戻るボタン ---
  btnBack.addEventListener('click', () => { setState('upload'); });

  // --- 生成ボタン（文書送付書）---
  btnGenerate.addEventListener('click', async () => {
    const requiredFields = ['courtName', 'courtFax', 'caseNumber', 'caseName',
                           'plaintiffName', 'defendantName', 'plaintiffLawyer',
                           'plaintiffLawyerFax', 'documentTitle'];
    const emptyFields = requiredFields.filter(key => !fields[key].value.trim());
    if (emptyFields.length > 0) {
      const fieldNames = {
        courtName: '裁判所名', courtFax: '裁判所FAX', caseNumber: '事件番号',
        caseName: '事件名', plaintiffName: '原告', defendantName: '被告',
        plaintiffLawyer: '原告代理人弁護士', plaintiffLawyerFax: '原告代理人FAX',
        documentTitle: '送付書類名',
      };
      const names = emptyFields.map(k => fieldNames[k]).join('、');
      if (!confirm(`以下の項目が未入力です:\n${names}\n\n空欄のまま生成しますか？`)) {
        return;
      }
    }
    setState('processing');
    processingTitle.textContent = '文書送付書を生成中...';
    processingMessage.textContent = 'テンプレートにデータを差し込んでいます。';
    startProcessingSteps('generate');

    const info = {
      courtName: fields.courtName.value.trim(),
      courtFax: fields.courtFax.value.trim(),
      caseNumber: fields.caseNumber.value.trim(),
      caseName: fields.caseName.value.trim(),
      plaintiffName: fields.plaintiffName.value.trim(),
      defendantName: fields.defendantName.value.trim(),
      plaintiffLawyer: fields.plaintiffLawyer.value.trim(),
      plaintiffLawyerFax: fields.plaintiffLawyerFax.value.trim(),
    };
    const documentTitle = fields.documentTitle.value.trim();

    try {
      const result = await generateDocumentBrowser(info, documentTitle, updateProgress);
      const blobUrl = URL.createObjectURL(result.blob);
      completeTitle.textContent = '文書送付書の生成が完了しました';
      downloadLabel.textContent = 'Wordファイルをダウンロード';
      outputFileName.textContent = result.fileName;
      btnDownload.href = blobUrl;
      btnDownload.download = result.fileName;
      singleDownloadArea.style.display = '';
      multiDownloadArea.style.display = 'none';
      setState('complete');
      setTimeout(launchConfetti, 300);
    } catch (err) {
      resetProcessingSteps();
      showError(err.message || '生成中にエラーが発生しました。');
      setState('confirm');
    }
  });

  // --- 新規ファイルボタン ---
  btnNewFile.addEventListener('click', () => {
    receiptUploadFiles = [];
    evidenceUploadFiles = [];
    completeTitle.textContent = '文書送付書の生成が完了しました';
    downloadLabel.textContent = 'Wordファイルをダウンロード';
    setState('upload');
  });

  // --- モード切替 ---
  function switchMode(mode) {
    currentMode = mode;
    modeSofushoBtn.classList.toggle('active', mode === 'sofusho');
    modeReceiptBtn.classList.toggle('active', mode === 'receipt');
    if (modeEvidenceBtn) modeEvidenceBtn.classList.toggle('active', mode === 'evidence');
    logoIcon.classList.remove('receipt-mode', 'evidence-mode');
    if (mode === 'receipt') {
      appTitle.textContent = '受領書自動でつくる君';
      logoIcon.classList.add('receipt-mode');
      uploadHeading.textContent = '相手方の文書送付書PDFをドロップ';
      uploadDesc.textContent = '受領日・署名・押印を自動で書き込みます';
    } else if (mode === 'evidence') {
      appTitle.textContent = '証拠番号つくる君';
      logoIcon.classList.add('evidence-mode');
      uploadHeading.textContent = '証拠PDFをドロップ';
      uploadDesc.textContent = '証拠番号と標目を右上に赤字でスタンプします';
    } else {
      appTitle.textContent = '文書送付書自動でつくる君';
      uploadHeading.textContent = 'PDF・Wordファイルをドロップ';
      uploadDesc.textContent = 'ファイルをここにドラッグ＆ドロップしてください';
    }
    receiptUploadFiles = [];
    evidenceUploadFiles = [];
    setState('upload');
  }
  modeSofushoBtn.addEventListener('click', () => switchMode('sofusho'));
  modeReceiptBtn.addEventListener('click', () => switchMode('receipt'));
  if (modeEvidenceBtn) modeEvidenceBtn.addEventListener('click', () => switchMode('evidence'));

  // --- 受領書モード: 複数ファイルを確認画面に表示 ---
  // --- 受領書プレビュー用変数 ---
  var receiptPreviewCanvas = $('#receiptPreviewCanvas');
  var receiptPreviewBox = $('#receiptPreviewBox');
  var receiptPreviewPlaceholder = $('#receiptPreviewPlaceholder');
  var receiptOverlaySensei = $('#receiptOverlaySensei');
  var receiptOverlayDate = $('#receiptOverlayDate');
  var receiptOverlaySign = $('#receiptOverlaySign');
  var receiptOverlaySeal = $('#receiptOverlaySeal');
  var receiptPreviewOcrData = null; // { words, imgWidth, imgHeight, pageNum, pgW, pgH }

  async function renderReceiptPreview(file) {
    if (!receiptPreviewCanvas || !receiptPreviewBox) return;
    try {
      // PDFを読み込み受領書ページを検索
      if (receiptPreviewPlaceholder) {
        receiptPreviewPlaceholder.style.display = '';
        receiptPreviewPlaceholder.innerHTML = '<div style="padding:20px;color:#9ca3af;font-size:0.85em;">受領書ページを検出中...</div>';
      }
      receiptPreviewCanvas.style.display = 'none';
      hideReceiptOverlays();

      var ab = await file.arrayBuffer();
      var pdfDocProxy = await pdfjsLib.getDocument({
        data: ab, cMapUrl: CMAP_URL, cMapPacked: true
      }).promise;

      var totalPages = pdfDocProxy.numPages;

      // 受領書ページ検出（OCR）
      var pdfDoc = await PDFLib.PDFDocument.load(ab);
      var result = await findReceiptPage(ab, totalPages, function(msg) {
        if (receiptPreviewPlaceholder) {
          receiptPreviewPlaceholder.innerHTML = '<div style="padding:20px;color:#9ca3af;font-size:0.85em;">' + msg + '</div>';
        }
      });
      var pageNum = result.pageNum;
      var ocr = result.ocr;
      var pgPage = pdfDoc.getPage(pageNum - 1);
      var pgSize = pgPage.getSize();

      receiptPreviewOcrData = {
        words: ocr.words,
        imgWidth: ocr.imgWidth,
        imgHeight: ocr.imgHeight,
        pageNum: pageNum,
        pgW: pgSize.width,
        pgH: pgSize.height,
      };

      // 該当ページをcanvasに描画
      var page = await pdfDocProxy.getPage(pageNum);
      var vp = page.getViewport({ scale: 1 });
      var boxW = receiptPreviewBox.clientWidth || 360;
      var scale = boxW / vp.width;
      var scaledVp = page.getViewport({ scale: scale });

      receiptPreviewCanvas.width = scaledVp.width;
      receiptPreviewCanvas.height = scaledVp.height;
      receiptPreviewCanvas.style.display = 'block';
      receiptPreviewCanvas.style.width = '100%';
      if (receiptPreviewPlaceholder) receiptPreviewPlaceholder.style.display = 'none';
      receiptPreviewBox.style.height = scaledVp.height + 'px';

      var ctx = receiptPreviewCanvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport: scaledVp }).promise;

      // 書き込み位置をオーバーレイ表示
      updateReceiptPreviewOverlays();

    } catch (e) {
      console.warn('[受領書Preview]', e);
      if (receiptPreviewPlaceholder) {
        receiptPreviewPlaceholder.style.display = '';
        receiptPreviewPlaceholder.innerHTML = '<div style="padding:20px;color:#f59e0b;font-size:0.85em;">プレビューを表示できませんでした<br><span style="font-size:0.85em;color:#9ca3af;">（生成は可能です）</span></div>';
      }
      receiptPreviewCanvas.style.display = 'none';
    }
  }

  function hideReceiptOverlays() {
    [receiptOverlaySensei, receiptOverlayDate, receiptOverlaySign, receiptOverlaySeal].forEach(function(el) {
      if (el) el.style.display = 'none';
    });
  }

  function updateReceiptPreviewOverlays() {
    if (!receiptPreviewOcrData || !receiptPreviewBox) return;
    var d = receiptPreviewOcrData;
    var canvasH = receiptPreviewCanvas.height;
    var canvasW = receiptPreviewCanvas.width;
    var boxW = receiptPreviewBox.clientWidth || 360;
    var boxH = receiptPreviewBox.clientHeight || canvasH;
    var scaleX = boxW / d.imgWidth;
    var scaleY = boxH / d.imgHeight;

    var pos = detectPositions(d.words, d.imgWidth, d.imgHeight, d.pgW, d.pgH);

    // 先生位置（行の右隣）
    if (pos.gyou && receiptOverlaySensei) {
      var gyouWord = d.words.find(function(w) { return w.text === '行' || w.text.includes('行'); });
      if (gyouWord) {
        receiptOverlaySensei.style.display = '';
        receiptOverlaySensei.style.left = ((gyouWord.x2 + 2) * scaleX) + 'px';
        receiptOverlaySensei.style.top = (gyouWord.y1 * scaleY) + 'px';
        receiptOverlaySensei.style.fontSize = Math.max(7, Math.round((gyouWord.y2 - gyouWord.y1) * scaleY * 0.7)) + 'px';
      }
    }

    // 日付位置
    var dateWord = d.words.find(function(w) {
      return w.text.includes('年') && w.text.includes('月');
    });
    if (!dateWord) {
      dateWord = d.words.find(function(w) { return w.text.includes('日付') || w.text.includes('年月日'); });
    }
    if (dateWord && receiptOverlayDate) {
      var today = new Date();
      var reiwaYear = today.getFullYear() - 2018;
      var dateText = receiptDateInput.value.trim() || ('令和' + reiwaYear + '年' + (today.getMonth()+1) + '月' + today.getDate() + '日');
      receiptOverlayDate.textContent = dateText;
      receiptOverlayDate.style.display = '';
      receiptOverlayDate.style.left = (dateWord.x1 * scaleX) + 'px';
      receiptOverlayDate.style.top = (dateWord.y1 * scaleY) + 'px';
      receiptOverlayDate.style.fontSize = Math.max(7, Math.round((dateWord.y2 - dateWord.y1) * scaleY * 0.7)) + 'px';
      receiptOverlayDate.style.background = 'rgba(255,255,255,0.9)';
      receiptOverlayDate.style.padding = '0 2px';
    }

    // 署名位置（代理人行）
    var agentWord = d.words.find(function(w) {
      return w.text.includes('代理人') || w.text.includes('弁護士');
    });
    if (agentWord && receiptOverlaySign) {
      var signerName = receiptSignerName.value.trim() || '';
      var signerTitle = receiptSignerTitle.value || '被告訴訟代理人';
      receiptOverlaySign.textContent = '　' + signerName;
      receiptOverlaySign.style.display = '';
      receiptOverlaySign.style.left = ((agentWord.x2 + 4) * scaleX) + 'px';
      receiptOverlaySign.style.top = (agentWord.y1 * scaleY) + 'px';
      receiptOverlaySign.style.fontSize = Math.max(7, Math.round((agentWord.y2 - agentWord.y1) * scaleY * 0.7)) + 'px';
      receiptOverlaySign.style.background = 'rgba(255,255,255,0.9)';
      receiptOverlaySign.style.padding = '0 2px';

      // 印鑑
      if (receiptOverlaySeal) {
        receiptOverlaySeal.style.display = '';
        var nameW = (signerName.length + 1) * Math.max(7, Math.round((agentWord.y2 - agentWord.y1) * scaleY * 0.7)) * 0.6;
        receiptOverlaySeal.style.left = ((agentWord.x2 + 4) * scaleX + nameW + 4) + 'px';
        receiptOverlaySeal.style.top = (agentWord.y1 * scaleY) + 'px';
        receiptOverlaySeal.style.fontSize = Math.max(7, Math.round((agentWord.y2 - agentWord.y1) * scaleY * 0.7)) + 'px';
      }
    }
  }

  // 受領書フォーム変更時にプレビューオーバーレイ更新
  [receiptSignerTitle, receiptSignerName, receiptDateInput].forEach(function(el) {
    if (el) {
      el.addEventListener('input', updateReceiptPreviewOverlays);
      el.addEventListener('change', updateReceiptPreviewOverlays);
    }
  });

  function prepareReceiptFiles(pdfs) {
    receiptUploadFiles = pdfs;
    if (pdfs.length === 1) {
      receiptSourceFileName.textContent = pdfs[0].name;
      receiptFileCountBadge.textContent = '受領書モード';
      receiptFileListWrap.style.display = 'none';
    } else {
      receiptSourceFileName.textContent = `${pdfs.length}件のPDFを処理します`;
      receiptFileCountBadge.textContent = `${pdfs.length}件`;
      receiptFileList.innerHTML = pdfs.map(f => `<li>${f.name}</li>`).join('');
      receiptFileListWrap.style.display = 'block';
    }
    setState('receipt-confirm');
    // 最初のPDFをプレビュー
    renderReceiptPreview(pdfs[0]);
  }

  // --- 受領書：戻るボタン ---
  btnReceiptBack.addEventListener('click', () => {
    receiptUploadFiles = [];
    setState('upload');
  });

  // --- 受領書：生成ボタン（複数ファイル対応）---
  btnReceiptGenerate.addEventListener('click', async () => {
    if (!receiptUploadFiles || receiptUploadFiles.length === 0) {
      showError('ファイルが選択されていません。');
      setState('upload');
      return;
    }
    const files = receiptUploadFiles;
    const total = files.length;
    const signerTitleVal = receiptSignerTitle.value;
    const signerNameVal = receiptSignerName.value;
    const receiptDateVal = receiptDateInput.value.trim();
    setState('processing');
    startProcessingSteps('upload');
    const results = [];

    for (let i = 0; i < total; i++) {
      const file = files[i];
      processingTitle.textContent = total > 1
        ? `受領書を生成中... (${i + 1}/${total})`
        : '受領書を生成中...';
      processingMessage.textContent = `${file.name} - OCRで位置検出＆書き込み中`;
      try {
        const result = await generateReceiptBrowser(file, {
          signerTitle: signerTitleVal,
          signerName: signerNameVal,
          receiptDate: receiptDateVal || undefined,
        }, updateProgress);
        const blobUrl = URL.createObjectURL(result.blob);
        results.push({ fileName: result.fileName, downloadUrl: blobUrl, error: null });
      } catch (err) {
        results.push({ fileName: file.name, downloadUrl: null, error: err.message || '生成失敗' });
      }
    }

    resetProcessingSteps();
    [procStep1, procStep2, procStep3].forEach(step => { if (step) step.classList.add('done'); });
    const succeeded = results.filter(r => !r.error);
    const failed = results.filter(r => r.error);
    completeTitle.textContent = total === 1
      ? '受領書の生成が完了しました'
      : `受領書の生成が完了しました（${succeeded.length}/${total}件）`;
    outputFileName.textContent = '';

    if (total === 1 && succeeded.length === 1) {
      singleDownloadArea.style.display = '';
      multiDownloadArea.style.display = 'none';
      downloadLabel.textContent = 'PDFファイルをダウンロード';
      btnDownload.href = succeeded[0].downloadUrl;
      btnDownload.download = succeeded[0].fileName;
    } else {
      singleDownloadArea.style.display = 'none';
      multiDownloadArea.style.display = '';
      downloadList.innerHTML = results.map(r => {
        if (r.error) {
          return `<li style="display:flex;align-items:center;gap:8px;padding:8px 12px;background:#fef2f2;border-radius:8px;color:#dc2626;">
            <span style="flex:1;font-size:0.85em;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${r.fileName}</span>
            <span style="font-size:0.8em;color:#dc2626;">失敗: ${r.error}</span>
          </li>`;
        }
        return `<li style="display:flex;align-items:center;gap:8px;padding:8px 12px;background:#f0fdf4;border-radius:8px;">
          <span style="flex:1;font-size:0.85em;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text-primary);">${r.fileName}</span>
          <a href="${r.downloadUrl}" download="${r.fileName}" class="btn btn-primary" style="padding:6px 14px;font-size:0.82em;white-space:nowrap;">
            ダウンロード
          </a>
        </li>`;
      }).join('');
    }

    if (failed.length > 0) {
      showError(`${failed.length}件の生成に失敗しました: ${failed.map(r=>r.fileName).join(', ')}`);
    }
    setState('complete');
    if (succeeded.length > 0) setTimeout(launchConfetti, 300);
  });

  // --- 証拠番号モード: DOM参照 ---
  const evidencePartySelect = $('#evidenceParty');
  const evidenceCustomParty = $('#evidenceCustomParty');
  const evidenceNumberInput = $('#evidenceNumber');
  const evidenceSubNumInput = $('#evidenceSubNum');
  const evidenceUseDai = $('#evidenceUseDai');
  const evidenceTitleInput = $('#evidenceTitle');
  const evidenceOriginalCopy = $('#evidenceOriginalCopy');
  const evidenceCreatedDate = $('#evidenceCreatedDate');
  const evidenceAuthor = $('#evidenceAuthor');
  const evidencePurpose = $('#evidencePurpose');
  const evidenceAllPages = $('#evidenceAllPages');
  const evidenceAddPageNum = $('#evidenceAddPageNum');
  const evidenceMintsName = $('#evidenceMintsName');
  const evidenceOutputSheet = $('#evidenceOutputSheet');
  const evidenceStampSize = $('#evidenceStampSize');
  const evidenceStampColor = $('#evidenceStampColor');
  const evidenceStampBg = $('#evidenceStampBg');
  const evidenceStampBorder = $('#evidenceStampBorder');
  const evidenceStampPreview = $('#evidenceStampPreview');
  const evidencePageNumPreview = $('#evidencePageNumPreview');
  const evidencePreviewBox = $('#evidencePreviewBox');
  const evidencePreviewCanvas = $('#evidencePreviewCanvas');
  const evidencePreviewPlaceholder = $('#evidencePreviewPlaceholder');
  const evidenceSourceFileName = $('#evidenceSourceFileName');
  const btnEvidenceBack = $('#btnEvidenceBack');
  const btnEvidenceGenerate = $('#btnEvidenceGenerate');
  const btnEvidenceMerge = $('#btnEvidenceMerge');
  const btnEvidenceAddFiles = $('#btnEvidenceAddFiles');
  const evidenceFileCountBadge = $('#evidenceFileCountBadge');
  const evidenceFileListWrap = $('#evidenceFileListWrap');
  const evidenceFileList = $('#evidenceFileList');

  let evidenceMergeMode = false;
  // スタンプ位置: PDF座標系での比率 (0~1), 左上原点
  // デフォルト: 右上 (x=0.85 y=0.03)
  let stampPosRatioX = 0.85;
  let stampPosRatioY = 0.03;
  // PDFページサイズ(pt)
  let previewPdfW = 595;
  let previewPdfH = 842;

  function getEvidenceParty() {
    if (evidenceCustomParty && evidenceCustomParty.value.trim()) {
      return evidenceCustomParty.value.trim();
    }
    return evidencePartySelect ? evidencePartySelect.value : '\u7532';
  }

  function getEvidencePreviewLabel() {
    const party = getEvidenceParty();
    const num = evidenceNumberInput ? evidenceNumberInput.value.trim() : '';
    const subNum = evidenceSubNumInput ? evidenceSubNumInput.value.trim() : '';
    const useDai = evidenceUseDai ? evidenceUseDai.checked : false;
    if (!num) return party + (useDai ? '\u7b2c' : '') + '\u25cb\u53f7\u8a3c';
    return buildEvidenceLabel(party, parseInt(num, 10), subNum ? parseInt(subNum, 10) : null, useDai);
  }

  // --- PDFプレビューレンダリング ---
  async function renderEvidencePreviewPdf(file) {
    if (!evidencePreviewCanvas || !evidencePreviewBox) return;
    try {
      var ab = await file.arrayBuffer();
      var pdf = await pdfjsLib.getDocument({
        data: ab, cMapUrl: CMAP_URL, cMapPacked: true
      }).promise;
      var page = await pdf.getPage(1);
      var vp = page.getViewport({ scale: 1 });
      previewPdfW = vp.width;
      previewPdfH = vp.height;

      // プレビューボックスの幅に合わせてスケール
      var boxW = evidencePreviewBox.clientWidth || 360;
      var scale = boxW / vp.width;
      var scaledVp = page.getViewport({ scale: scale });

      evidencePreviewCanvas.width = scaledVp.width;
      evidencePreviewCanvas.height = scaledVp.height;
      evidencePreviewCanvas.style.display = 'block';
      evidencePreviewCanvas.style.width = '100%';
      if (evidencePreviewPlaceholder) evidencePreviewPlaceholder.style.display = 'none';

      // ボックス高さをcanvas比率に合わせる
      evidencePreviewBox.style.height = scaledVp.height + 'px';

      var ctx = evidencePreviewCanvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport: scaledVp }).promise;

      // スタンプオーバーレイを表示
      if (evidenceStampPreview) {
        evidenceStampPreview.style.display = '';
        positionStampFromRatio();
      }
    } catch (e) {
      console.warn('[Preview] PDF\u30ec\u30f3\u30c0\u30ea\u30f3\u30b0\u5931\u6557:', e);
      // フォールバック: プレースホルダーのまま
      showPreviewPlaceholder();
    }
  }

  function showPreviewPlaceholder() {
    if (evidencePreviewCanvas) evidencePreviewCanvas.style.display = 'none';
    if (evidencePreviewPlaceholder) evidencePreviewPlaceholder.style.display = '';
    if (evidenceStampPreview) evidenceStampPreview.style.display = '';
    evidencePreviewBox.style.height = '320px';
    positionStampFromRatio();
  }

  function positionStampFromRatio() {
    if (!evidenceStampPreview || !evidencePreviewBox) return;
    var boxW = evidencePreviewBox.clientWidth || 360;
    var boxH = evidencePreviewBox.clientHeight || 320;
    var stampW = evidenceStampPreview.offsetWidth;
    var stampH = evidenceStampPreview.offsetHeight;
    var left = stampPosRatioX * boxW - stampW / 2;
    var top = stampPosRatioY * boxH;
    // Clamp
    left = Math.max(0, Math.min(boxW - stampW, left));
    top = Math.max(0, Math.min(boxH - stampH, top));
    evidenceStampPreview.style.left = left + 'px';
    evidenceStampPreview.style.top = top + 'px';
    evidenceStampPreview.style.right = '';
  }

  // --- スタンプドラッグ ---
  (function setupStampDrag() {
    if (!evidenceStampPreview || !evidencePreviewBox) return;
    var dragging = false;
    var offsetX = 0, offsetY = 0;

    function onStart(ex, ey) {
      dragging = true;
      var rect = evidenceStampPreview.getBoundingClientRect();
      offsetX = ex - rect.left;
      offsetY = ey - rect.top;
      evidenceStampPreview.style.cursor = 'grabbing';
      evidenceStampPreview.style.transition = 'none';
    }
    function onMove(ex, ey) {
      if (!dragging) return;
      var boxRect = evidencePreviewBox.getBoundingClientRect();
      var boxW = evidencePreviewBox.clientWidth;
      var boxH = evidencePreviewBox.clientHeight;
      var stampW = evidenceStampPreview.offsetWidth;
      var stampH = evidenceStampPreview.offsetHeight;

      var newLeft = ex - boxRect.left - offsetX;
      var newTop = ey - boxRect.top - offsetY;
      newLeft = Math.max(0, Math.min(boxW - stampW, newLeft));
      newTop = Math.max(0, Math.min(boxH - stampH, newTop));

      evidenceStampPreview.style.left = newLeft + 'px';
      evidenceStampPreview.style.top = newTop + 'px';
      evidenceStampPreview.style.right = '';

      // 比率を更新 (中心基準)
      stampPosRatioX = (newLeft + stampW / 2) / boxW;
      stampPosRatioY = newTop / boxH;
    }
    function onEnd() {
      if (!dragging) return;
      dragging = false;
      evidenceStampPreview.style.cursor = 'grab';
      evidenceStampPreview.style.transition = '';
    }

    // マウス
    evidenceStampPreview.addEventListener('mousedown', function(e) {
      e.preventDefault();
      onStart(e.clientX, e.clientY);
    });
    document.addEventListener('mousemove', function(e) {
      if (dragging) { e.preventDefault(); onMove(e.clientX, e.clientY); }
    });
    document.addEventListener('mouseup', onEnd);

    // タッチ
    evidenceStampPreview.addEventListener('touchstart', function(e) {
      e.preventDefault();
      var t = e.touches[0];
      onStart(t.clientX, t.clientY);
    }, { passive: false });
    document.addEventListener('touchmove', function(e) {
      if (dragging) { var t = e.touches[0]; onMove(t.clientX, t.clientY); }
    }, { passive: false });
    document.addEventListener('touchend', onEnd);
  })();

  function updateEvidencePreview() {
    var label = getEvidencePreviewLabel();
    var title = evidenceTitleInput ? evidenceTitleInput.value.trim() : '';

    // スタンプテキスト更新（証拠番号のみ、書誌情報はPDFに打ち込まない）
    if (evidenceStampPreview) {
      var text = label;
      evidenceStampPreview.textContent = text;

      // 色
      var colorMap = { red: '#dc2626', black: '#111', blue: '#1d4ed8' };
      var colorVal = evidenceStampColor ? evidenceStampColor.value : 'red';
      evidenceStampPreview.style.color = colorMap[colorVal] || '#dc2626';

      // サイズ
      var sizeVal = evidenceStampSize ? parseInt(evidenceStampSize.value) : 20;
      evidenceStampPreview.style.fontSize = Math.round(sizeVal * 0.55) + 'px';

      // 背景・枠線
      var hasBg = evidenceStampBg ? evidenceStampBg.checked : true;
      var hasBorder = evidenceStampBorder ? evidenceStampBorder.checked : false;
      evidenceStampPreview.style.background = hasBg ? 'rgba(255,255,255,0.95)' : 'transparent';
      if (hasBorder) {
        evidenceStampPreview.style.border = '1px solid ' + (colorMap[colorVal] || '#dc2626');
      } else {
        evidenceStampPreview.style.border = '1px solid transparent';
      }

      // 位置を再適用（サイズ変更で幅が変わるため）
      setTimeout(positionStampFromRatio, 10);
    }

    // ページ番号プレビュー
    if (evidencePageNumPreview) {
      evidencePageNumPreview.style.display = (evidenceAddPageNum && evidenceAddPageNum.checked) ? '' : 'none';
    }

    updateEvidenceFileListLabels();
  }

  function updateEvidenceFileListLabels() {
    if (!evidenceFileList || !evidenceUploadFiles || evidenceUploadFiles.length === 0) return;
    var party = getEvidenceParty();
    var startNum = parseInt(evidenceNumberInput ? evidenceNumberInput.value : '1', 10) || 1;
    var useDai = evidenceUseDai ? evidenceUseDai.checked : false;
    var subNum = evidenceSubNumInput ? evidenceSubNumInput.value.trim() : '';

    var items = evidenceFileList.querySelectorAll('li');
    items.forEach(function(li, i) {
      var labelEl = li.querySelector('.file-label');
      if (!labelEl) return;
      if (evidenceMergeMode) {
        labelEl.textContent = buildEvidenceLabel(party, startNum, i + 1, useDai);
      } else {
        var currentSub = (evidenceUploadFiles.length === 1 && subNum) ? parseInt(subNum, 10) : null;
        labelEl.textContent = buildEvidenceLabel(party, startNum + i, currentSub, useDai);
      }
    });
  }

  // プレビュー更新対象コントロール
  var evidencePreviewControls = [
    evidencePartySelect, evidenceCustomParty, evidenceNumberInput,
    evidenceSubNumInput, evidenceTitleInput,
    evidenceStampSize, evidenceStampColor
  ];
  evidencePreviewControls.forEach(function(el) {
    if (el) {
      el.addEventListener('input', updateEvidencePreview);
      el.addEventListener('change', updateEvidencePreview);
    }
  });
  [evidenceUseDai, evidenceStampBg, evidenceStampBorder, evidenceAddPageNum].forEach(function(el) {
    if (el) el.addEventListener('change', updateEvidencePreview);
  });

  function renderEvidenceFileList() {
    if (!evidenceFileList || !evidenceUploadFiles) return;
    var dragIcon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 6h.01M8 12h.01M8 18h.01M12 6h.01M12 12h.01M12 18h.01"/></svg>';
    var removeIcon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';

    evidenceFileList.innerHTML = evidenceUploadFiles.map(function(f, i) {
      return '<li draggable="true" data-index="' + i + '">' +
        '<span class="drag-handle">' + dragIcon + '</span>' +
        '<span class="file-name">' + f.name + '</span>' +
        '<span class="file-label"></span>' +
        '<button class="file-remove" data-index="' + i + '" title="\u524a\u9664">' + removeIcon + '</button>' +
        '</li>';
    }).join('');

    updateEvidenceFileListLabels();

    // ドラッグ並べ替え
    var dragSrcIndex = null;
    evidenceFileList.querySelectorAll('li').forEach(function(li) {
      li.addEventListener('dragstart', function(e) {
        dragSrcIndex = parseInt(li.dataset.index);
        li.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
      });
      li.addEventListener('dragover', function(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
      });
      li.addEventListener('drop', function(e) {
        e.preventDefault();
        var targetIndex = parseInt(li.dataset.index);
        if (dragSrcIndex !== null && dragSrcIndex !== targetIndex) {
          var moved = evidenceUploadFiles.splice(dragSrcIndex, 1)[0];
          evidenceUploadFiles.splice(targetIndex, 0, moved);
          renderEvidenceFileList();
        }
      });
      li.addEventListener('dragend', function() {
        li.classList.remove('dragging');
        dragSrcIndex = null;
      });
    });

    // 削除ボタン
    evidenceFileList.querySelectorAll('.file-remove').forEach(function(btn) {
      btn.addEventListener('click', function() {
        var idx = parseInt(btn.dataset.index);
        evidenceUploadFiles.splice(idx, 1);
        if (evidenceUploadFiles.length === 0) {
          setState('upload');
          return;
        }
        renderEvidenceFileList();
        updateEvidenceSourceInfo();
        // 先頭ファイルのプレビュー更新
        renderEvidencePreviewPdf(evidenceUploadFiles[0]);
      });
    });
  }

  function updateEvidenceSourceInfo() {
    if (evidenceSourceFileName) {
      evidenceSourceFileName.textContent = evidenceUploadFiles.length === 1
        ? evidenceUploadFiles[0].name
        : evidenceUploadFiles.length + '\u4ef6\u306ePDF\u3092\u51e6\u7406\u3057\u307e\u3059';
    }
    if (evidenceFileCountBadge) {
      evidenceFileCountBadge.textContent = evidenceUploadFiles.length > 1
        ? evidenceUploadFiles.length + '\u4ef6'
        : '\u8a3c\u62e0\u756a\u53f7\u30e2\u30fc\u30c9';
    }
    if (evidenceFileListWrap) {
      evidenceFileListWrap.style.display = evidenceUploadFiles.length > 1 ? 'block' : 'none';
    }
  }

  function prepareEvidenceFiles(pdfs) {
    evidenceUploadFiles = Array.from(pdfs);
    evidenceMergeMode = false;
    // デフォルト位置リセット
    stampPosRatioX = 0.85;
    stampPosRatioY = 0.03;
    updateEvidenceSourceInfo();
    renderEvidenceFileList();
    updateEvidencePreview();
    // 1ページ目をプレビュー
    renderEvidencePreviewPdf(evidenceUploadFiles[0]);
    setState('evidence-confirm');
  }

  // 結合ボタン
  if (btnEvidenceMerge) {
    btnEvidenceMerge.addEventListener('click', function() {
      evidenceMergeMode = !evidenceMergeMode;
      btnEvidenceMerge.classList.toggle('active', evidenceMergeMode);
      btnEvidenceMerge.style.background = evidenceMergeMode ? '#fef2f2' : '';
      btnEvidenceMerge.style.color = evidenceMergeMode ? '#dc2626' : '';
      updateEvidenceFileListLabels();
    });
  }

  // ファイル追加ボタン
  if (btnEvidenceAddFiles) {
    btnEvidenceAddFiles.addEventListener('click', function() {
      var input = document.createElement('input');
      input.type = 'file';
      input.accept = '.pdf';
      input.multiple = true;
      input.addEventListener('change', function() {
        if (input.files && input.files.length > 0) {
          for (var i = 0; i < input.files.length; i++) {
            evidenceUploadFiles.push(input.files[i]);
          }
          updateEvidenceSourceInfo();
          renderEvidenceFileList();
        }
      });
      input.click();
    });
  }

  if (btnEvidenceBack) {
    btnEvidenceBack.addEventListener('click', function() {
      evidenceUploadFiles = [];
      evidenceMergeMode = false;
      showPreviewPlaceholder();
      setState('upload');
    });
  }

  if (btnEvidenceGenerate) {
    btnEvidenceGenerate.addEventListener('click', async function() {
      if (!evidenceUploadFiles || evidenceUploadFiles.length === 0) {
        showError('\u30d5\u30a1\u30a4\u30eb\u304c\u9078\u629e\u3055\u308c\u3066\u3044\u307e\u305b\u3093\u3002');
        setState('upload');
        return;
      }
      var party = getEvidenceParty();
      var numStr = evidenceNumberInput ? evidenceNumberInput.value.trim() : '';
      if (!numStr) {
        showError('\u8a3c\u62e0\u756a\u53f7\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002');
        return;
      }
      var startNum = parseInt(numStr, 10);
      if (isNaN(startNum) || startNum < 1) {
        showError('\u8a3c\u62e0\u756a\u53f7\u306f1\u4ee5\u4e0a\u306e\u6570\u5b57\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002');
        return;
      }
      var subNumStr = evidenceSubNumInput ? evidenceSubNumInput.value.trim() : '';
      var subNum = subNumStr ? parseInt(subNumStr, 10) : null;
      var useDai = evidenceUseDai ? evidenceUseDai.checked : false;
      var titleStr = evidenceTitleInput ? evidenceTitleInput.value.trim() : '';
      var allPages = evidenceAllPages ? evidenceAllPages.checked : false;
      var addPageNum = evidenceAddPageNum ? evidenceAddPageNum.checked : false;
      var useMintsName = evidenceMintsName ? evidenceMintsName.checked : false;
      var outputSheet = evidenceOutputSheet ? evidenceOutputSheet.checked : false;
      var originalCopy = evidenceOriginalCopy ? evidenceOriginalCopy.value : '';
      var createdDate = evidenceCreatedDate ? evidenceCreatedDate.value.trim() : '';
      var author = evidenceAuthor ? evidenceAuthor.value.trim() : '';
      var purpose = evidencePurpose ? evidencePurpose.value.trim() : '';

      // スタンプ位置: ドラッグ結果のratioからPDF座標(pt)に変換
      var stampOpts = {
        stampSize: evidenceStampSize ? evidenceStampSize.value : '20',
        stampColor: evidenceStampColor ? evidenceStampColor.value : 'red',
        stampBg: evidenceStampBg ? evidenceStampBg.checked : true,
        stampBorder: evidenceStampBorder ? evidenceStampBorder.checked : false,
        addPageNum: addPageNum,
        // カスタム位置 (PDF座標系)
        customX: stampPosRatioX,
        customY: stampPosRatioY,
      };

      var files = evidenceUploadFiles;
      var total = files.length;

      setState('processing');
      startProcessingSteps('generate');
      var results = [];
      var sheetEntries = [];

      // 結合モード
      if (evidenceMergeMode && total > 1) {
        try {
          processingTitle.textContent = 'PDF\u3092\u7d50\u5408\u4e2d...';
          var mergedFile = await mergePdfs(files, updateProgress);
          var subRange = '1~' + total;
          var evidenceLabel = buildEvidenceLabel(party, startNum, null, useDai);
          processingTitle.textContent = '\u8a3c\u62e0\u756a\u53f7\u3092\u66f8\u304d\u8fbc\u307f\u4e2d...';
          processingMessage.textContent = evidenceLabel;

          var result = await generateEvidenceBrowser(mergedFile, Object.assign({
            evidenceLabel: evidenceLabel, evidenceTitle: titleStr, allPages: allPages, onProgress: updateProgress,
          }, stampOpts));

          var outName = useMintsName
            ? buildMintsFileName(party, startNum, null, titleStr, subRange)
            : result.fileName;
          var blobUrl = URL.createObjectURL(result.blob);
          results.push({ fileName: outName, downloadUrl: blobUrl, error: null });
          sheetEntries.push({
            label: evidenceLabel + '\u306e\uff11\uff5e' + toFullWidthNumber(String(total)),
            title: titleStr, originalOrCopy: originalCopy,
            createdDate: createdDate, author: author, purpose: purpose,
          });
        } catch (err) {
          results.push({ fileName: '\u7d50\u5408\u30d5\u30a1\u30a4\u30eb', downloadUrl: null, error: err.message || '\u7d50\u5408\u5931\u6557' });
        }
      } else {
        for (var i = 0; i < total; i++) {
          var currentNum = startNum + i;
          var currentSub = (total === 1 && subNum) ? subNum : null;
          var evidenceLabel = buildEvidenceLabel(party, currentNum, currentSub, useDai);
          processingTitle.textContent = total > 1
            ? '\u8a3c\u62e0\u756a\u53f7\u3092\u66f8\u304d\u8fbc\u307f\u4e2d... (' + (i + 1) + '/' + total + ')'
            : '\u8a3c\u62e0\u756a\u53f7\u3092\u66f8\u304d\u8fbc\u307f\u4e2d...';
          processingMessage.textContent = files[i].name + ' \u2192 ' + evidenceLabel;

          try {
            var result = await generateEvidenceBrowser(files[i], Object.assign({
              evidenceLabel: evidenceLabel, evidenceTitle: titleStr, allPages: allPages, onProgress: updateProgress,
            }, stampOpts));

            var outName = useMintsName
              ? buildMintsFileName(party, currentNum, currentSub, titleStr, null)
              : result.fileName;
            var blobUrl = URL.createObjectURL(result.blob);
            results.push({ fileName: outName, downloadUrl: blobUrl, error: null });
            sheetEntries.push({
              label: evidenceLabel, title: titleStr, originalOrCopy: originalCopy,
              createdDate: createdDate, author: author, purpose: purpose,
            });
          } catch (err) {
            results.push({ fileName: files[i].name, downloadUrl: null, error: err.message || '\u751f\u6210\u5931\u6557' });
          }
        }
      }

      // 証拠説明書
      if (outputSheet && sheetEntries.length > 0) {
        try {
          var sheetBlob = await generateEvidenceSheetDocx(sheetEntries, { party: party });
          var sheetUrl = URL.createObjectURL(sheetBlob);
          results.push({
            fileName: '\u8a3c\u62e0\u8aac\u660e\u66f8_' + party + '\u53f7\u8a3c.docx',
            downloadUrl: sheetUrl, error: null, isSheet: true,
          });
        } catch (sheetErr) {
          console.error('\u8a3c\u62e0\u8aac\u660e\u66f8\u306e\u751f\u6210\u306b\u5931\u6557:', sheetErr);
        }
      }

      resetProcessingSteps();
      [procStep1, procStep2, procStep3].forEach(function(step) { if (step) step.classList.add('done'); });
      var succeeded = results.filter(function(r) { return !r.error; });
      var failed = results.filter(function(r) { return r.error; });
      var pdfCount = succeeded.filter(function(r) { return !r.isSheet; }).length;
      completeTitle.textContent = total === 1
        ? '\u8a3c\u62e0\u756a\u53f7\u306e\u66f8\u304d\u8fbc\u307f\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f'
        : '\u8a3c\u62e0\u756a\u53f7\u306e\u66f8\u304d\u8fbc\u307f\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\uff08' + pdfCount + '/' + total + '\u4ef6\uff09';
      outputFileName.textContent = '';

      if (succeeded.length === 1 && !outputSheet) {
        singleDownloadArea.style.display = '';
        multiDownloadArea.style.display = 'none';
        downloadLabel.textContent = 'PDF\u30d5\u30a1\u30a4\u30eb\u3092\u30c0\u30a6\u30f3\u30ed\u30fc\u30c9';
        btnDownload.href = succeeded[0].downloadUrl;
        btnDownload.download = succeeded[0].fileName;
      } else {
        singleDownloadArea.style.display = 'none';
        multiDownloadArea.style.display = '';
        downloadList.innerHTML = results.map(function(r) {
          if (r.error) {
            return '<li style="display:flex;align-items:center;gap:8px;padding:8px 12px;background:#fef2f2;border-radius:8px;color:#dc2626;">' +
              '<span style="flex:1;font-size:0.85em;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">' + r.fileName + '</span>' +
              '<span style="font-size:0.8em;color:#dc2626;">\u5931\u6557: ' + r.error + '</span></li>';
          }
          var bgColor = r.isSheet ? '#eff6ff' : '#f0fdf4';
          return '<li style="display:flex;align-items:center;gap:8px;padding:8px 12px;background:' + bgColor + ';border-radius:8px;">' +
            '<span style="flex:1;font-size:0.85em;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text-primary);">' + r.fileName + '</span>' +
            '<a href="' + r.downloadUrl + '" download="' + r.fileName + '" class="btn btn-primary" style="padding:6px 14px;font-size:0.82em;white-space:nowrap;">\u30c0\u30a6\u30f3\u30ed\u30fc\u30c9</a></li>';
        }).join('');
      }

      if (failed.length > 0) {
        showError(failed.length + '\u4ef6\u306e\u751f\u6210\u306b\u5931\u6557\u3057\u307e\u3057\u305f: ' + failed.map(function(r) { return r.fileName; }).join(', '));
      }
      setState('complete');
      if (pdfCount > 0) setTimeout(launchConfetti, 300);
    });
  }


  // --- 設定モーダル ---
  if (settingsBtn && settingsModal) {
    settingsBtn.addEventListener('click', () => {
      const config = JSON.parse(localStorage.getItem('tsukurukun_config') || '{}');
      const settingsOfficeName = $('#settingsOfficeName');
      const settingsSignerName = $('#settingsSignerName');
      const settingsLawyerNames = $('#settingsLawyerNames');
      const settingsFaxNumbers = $('#settingsFaxNumbers');
      const sealPreview = $('#sealPreview');
      if (settingsOfficeName) settingsOfficeName.value = config.officeName || '';
      if (settingsSignerName) settingsSignerName.value = config.signerName || '';
      if (settingsLawyerNames) settingsLawyerNames.value = (config.lawyerNames || []).join(', ');
      if (settingsFaxNumbers) settingsFaxNumbers.value = (config.faxNumbers || []).join(', ');
      const sealBase64 = localStorage.getItem('tsukurukun_seal');
      if (sealPreview) {
        if (sealBase64) {
          sealPreview.innerHTML = `<img src="${sealBase64}" style="max-width:60px;max-height:60px;">`;
        } else {
          sealPreview.innerHTML = '<span style="color:var(--text-secondary);font-size:0.85em;">未設定（㊞で代替）</span>';
        }
      }
      settingsModal.classList.add('visible');
    });

    settingsClose.addEventListener('click', () => {
      settingsModal.classList.remove('visible');
    });

    const sealInput = $('#sealInput');
    if (sealInput) {
      sealInput.addEventListener('change', () => {
        const file = sealInput.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
          localStorage.setItem('tsukurukun_seal', e.target.result);
          const sealPreview = $('#sealPreview');
          if (sealPreview) {
            sealPreview.innerHTML = `<img src="${e.target.result}" style="max-width:60px;max-height:60px;">`;
          }
        };
        reader.readAsDataURL(file);
      });
    }

    const sealRemove = $('#sealRemove');
    if (sealRemove) {
      sealRemove.addEventListener('click', () => {
        localStorage.removeItem('tsukurukun_seal');
        const sealPreview = $('#sealPreview');
        if (sealPreview) {
          sealPreview.innerHTML = '<span style="color:var(--text-secondary);font-size:0.85em;">未設定（㊞で代替）</span>';
        }
        if (sealInput) sealInput.value = '';
      });
    }

    settingsSave.addEventListener('click', () => {
      const settingsOfficeName = $('#settingsOfficeName');
      const settingsSignerName = $('#settingsSignerName');
      const settingsLawyerNames = $('#settingsLawyerNames');
      const settingsFaxNumbers = $('#settingsFaxNumbers');
      const config = {
        officeName: settingsOfficeName ? settingsOfficeName.value.trim() : '',
        signerName: settingsSignerName ? settingsSignerName.value.trim() : '',
        lawyerNames: settingsLawyerNames
          ? settingsLawyerNames.value.split(/[,、]/).map(s => s.trim()).filter(Boolean)
          : [],
        faxNumbers: settingsFaxNumbers
          ? settingsFaxNumbers.value.split(/[,、]/).map(s => s.trim()).filter(Boolean)
          : [],
      };
      localStorage.setItem('tsukurukun_config', JSON.stringify(config));
      const subtitle = $('#officeSubtitle');
      if (subtitle) subtitle.textContent = config.officeName;
      if (receiptSignerName && config.signerName) {
        receiptSignerName.value = config.signerName;
      }
      settingsModal.classList.remove('visible');
    });

    settingsModal.addEventListener('click', (e) => {
      if (e.target === settingsModal) {
        settingsModal.classList.remove('visible');
      }
    });
  }

  // --- 起動時に設定を読み込み ---
  (function loadConfig() {
    try {
      const config = JSON.parse(localStorage.getItem('tsukurukun_config') || '{}');
      const subtitle = $('#officeSubtitle');
      if (subtitle && config.officeName) subtitle.textContent = config.officeName;
      if (config.signerName && receiptSignerName && !receiptSignerName.value) {
        receiptSignerName.value = config.signerName;
      }
    } catch (e) { /* ignore */ }
  })();

  // --- ページ離脱前の確認 ---
  window.addEventListener('beforeunload', (e) => {
    if (currentState === 'confirm' || currentState === 'receipt-confirm' || currentState === 'evidence-confirm') {
      e.preventDefault();
      e.returnValue = '';
    }
  });

})();
