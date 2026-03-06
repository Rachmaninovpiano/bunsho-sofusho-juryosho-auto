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

  // --- 日本語フォント読み込み（キャッシュ + CDNフォールバック） ---
  let _cachedFontBytes = null;
  async function loadJapaneseFont() {
    if (_cachedFontBytes) return _cachedFontBytes;
    const urls = [
      'fonts/NotoSerifJP.ttf',
      'https://cdn.jsdelivr.net/gh/Rachmaninovpiano/bunsho-sofusho-juryosho-auto@master/docs/fonts/NotoSerifJP.ttf',
    ];
    for (const url of urls) {
      try {
        const resp = await fetch(url);
        if (resp.ok) {
          _cachedFontBytes = await resp.arrayBuffer();
          return _cachedFontBytes;
        }
      } catch (e) { continue; }
    }
    throw new Error('日本語フォントの読み込みに失敗しました。GitHub Pagesでお試しください。');
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
      const viewport = page.getViewport({ scale: 300 / 72 });
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
  // Section 3: 文書送付書 情報抽出
  // =============================================

  function extractInfoFromText(text) {
    const config = getConfig();
    const info = {};
    const normalizedText = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

    // 都市名リスト
    const cityNames = '東京|大阪|名古屋|広島|福岡|仙台|札幌|高松|京都|神戸|横浜|さいたま|千葉|山口|岡山|福山|松山|高知|那覇|長崎|熊本|鹿児島|大分|宮崎|佐賀|秋田|青森|盛岡|山形|福島|水戸|宇都宮|前橋|甲府|長野|新潟|富山|金沢|福井|津|大津|奈良|和歌山|鳥取|松江|徳島|旭川|釧路|函館';

    // --- 裁判所名 ---
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
      const match = normalizedText.match(pattern);
      if (match) {
        let cn = match[1].replace(/\s+/g, '');
        cn = cn.replace(/（/g, '(').replace(/）/g, ')');
        info.caseNumber = cn;
        break;
      }
    }

    // パターン2: 「事件の表示」セクションから
    if (!info.caseNumber) {
      const displaySectionMatch = normalizedText.match(
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
      /損\s*害\s*賠\s*償[\s\S]{0,80}?請\s*求\s*事\s*件/,
      /(?:号\s*)([\u4e00-\u9fff][\u4e00-\u9fff\s]*(?:請\s*求|確\s*認|等?)\s*事\s*件)/,
      /(損\s*害\s*賠\s*償\s*請\s*求\s*事\s*件|貸\s*金\s*返\s*還\s*請\s*求\s*事\s*件|建\s*物\s*明\s*渡\s*請\s*求\s*事\s*件|不\s*当\s*利\s*得\s*返\s*還\s*請\s*求\s*事\s*件)/,
      /([\u4e00-\u9fff][\u4e00-\u9fff\s]*(?:請\s*求|確\s*認)\s*事\s*件)/,
    ];
    for (const pattern of caseNamePatterns) {
      const match = normalizedText.match(pattern);
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
    const partySection = normalizedText.match(/当\s*事\s*者[\s\S]{0,200}/);
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
        while ((match = globalPattern.exec(normalizedText)) !== null) {
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
        while ((match = globalPattern.exec(normalizedText)) !== null) {
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
    while ((lm = formalPattern.exec(normalizedText)) !== null) {
      const name = cleanLawyerName(lm[1]);
      if (name) lawyerCandidates.push({ name, priority: 1 });
    }

    const senderPattern = /人\s*弁護\s*士\s*([^\n]{2,20})/g;
    while ((lm = senderPattern.exec(normalizedText)) !== null) {
      const contextBefore = normalizedText.substring(Math.max(0, lm.index - 30), lm.index);
      if (contextBefore.includes('被告')) continue;
      const name = cleanLawyerName(lm[1]);
      if (name) lawyerCandidates.push({ name, priority: 2 });
    }

    const atePattern = /弁護\s*士\s*([^\n]{2,15})\s*宛/g;
    while ((lm = atePattern.exec(normalizedText)) !== null) {
      const name = cleanLawyerName(lm[1]);
      if (name) lawyerCandidates.push({ name, priority: 3 });
    }

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

    const explicitCourtFaxMatch = normalizedText.match(
      /裁\s*判\s*所[\s\S]{0,60}?[（(]\s*(?:FAX|ＦＡＸ|[Ff]ax)\s*([0-9０-９\-－ー・]+)\s*[）)]/
    );
    if (explicitCourtFaxMatch) {
      info.courtFaxFromPdf = normalizeFax(explicitCourtFaxMatch[1]);
    }

    const allExplicitFaxes = [];
    const explicitFaxRegex = /([\u4e00-\u9fff]{1,10})\s*[（(]\s*(?:FAX|ＦＡＸ|[Ff]ax)\s*([0-9０-９\-－ー・]+)\s*[）)]/g;
    let efm;
    while ((efm = explicitFaxRegex.exec(normalizedText)) !== null) {
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
    while ((faxMatch = faxRegex.exec(normalizedText)) !== null) {
      const faxNum = normalizeFax(faxMatch[1]);
      allFaxEntries.push({ fax: faxNum, index: faxMatch.index });
    }

    for (const entry of allFaxEntries) {
      const isOwnFax = ownFaxPatterns.some(p => entry.fax.includes(p));
      if (isOwnFax) continue;
      if (info.courtFaxFromPdf && entry.fax.includes(info.courtFaxFromPdf)) continue;
      const isKnownCourtFax = courtFaxValues.some(cf => entry.fax.includes(cf));
      const textBefore = normalizedText.substring(
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
    console.log('[つくる君] PDF解析開始:', file.name, file.size, 'bytes');
    onProgress && onProgress('PDFを読み込み中...');
    const pdfArrayBuffer = await file.arrayBuffer();
    const pdfText = await extractTextBrowser(pdfArrayBuffer, onProgress);
    console.log('[つくる君] 抽出テキスト:', pdfText.length, '文字');
    onProgress && onProgress('情報を抽出中...');
    const info = extractInfoFromText(pdfText);
    console.log('[つくる君] 抽出結果:', JSON.stringify(info, null, 2));
    const documentTitle = getDocumentTitleFromFilename(file.name);
    return { info, documentTitle, originalName: file.name };
  }

  async function generateDocumentBrowser(info, documentTitle, onProgress) {
    onProgress && onProgress('テンプレートを読み込み中...');
    const templateResp = await fetch('template/文書送付書.doc.docx');
    if (!templateResp.ok) {
      throw new Error('テンプレートファイルの読み込みに失敗しました。');
    }
    const templateData = await templateResp.arrayBuffer();
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
    const viewport = page.getViewport({ scale: 300 / 72 });
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
   * 証拠番号ラベルを組み立てる。
   * @param {string} party - 甲, 乙, 丙, 甲A, 乙B 等
   * @param {number} num - 号証番号
   * @param {number|null} subNum - 枝番（null = なし）
   * @param {boolean} useDai - 「第」を入れるか（甲第1号証 vs 甲1号証）
   * @returns {string}
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
   * Canvasでテキストを描画し、PNG画像のバイト列を返す。
   * フォントファイル不要でfile://プロトコルでも動作する。
   */
  function renderTextToImageBytes(text, fontSize, options) {
    const opts = options || {};
    const color = opts.color || 'red';
    const fontFamily = opts.fontFamily || '"Yu Mincho", "游明朝", "MS Mincho", "ＭＳ 明朝", "Hiragino Mincho ProN", serif';
    const scale = opts.scale || 4;
    const bold = opts.bold !== false;

    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const fontStr = (bold ? 'bold ' : '') + (fontSize * scale) + 'px ' + fontFamily;

    ctx.font = fontStr;
    const metrics = ctx.measureText(text);
    const pad = scale * 2;

    canvas.width = Math.ceil(metrics.width) + pad * 2;
    canvas.height = Math.ceil(fontSize * scale * 1.35) + pad;

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.font = fontStr;
    ctx.fillStyle = color;
    ctx.textBaseline = 'top';
    ctx.fillText(text, pad, pad / 2);

    const dataUrl = canvas.toDataURL('image/png');
    const base64 = dataUrl.split(',')[1];
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

    return { bytes: bytes, width: canvas.width / scale, height: canvas.height / scale };
  }

  /**
   * PDFの指定ページ右上に証拠番号と標目を赤字でスタンプする。
   * Canvas描画 → PNG埋め込み方式（フォントファイル不要）。
   */
  async function generateEvidenceBrowser(file, opts) {
    const { evidenceLabel, evidenceTitle, allPages } = opts;
    const pdfArrayBuffer = await file.arrayBuffer();
    const pdfDoc = await PDFLib.PDFDocument.load(pdfArrayBuffer);

    // 証拠番号をCanvas→PNG画像として生成
    const labelImg = renderTextToImageBytes(evidenceLabel, 28);
    const pdfLabelImage = await pdfDoc.embedPng(labelImg.bytes);

    let pdfTitleImage = null;
    let titleImgDims = null;
    if (evidenceTitle && evidenceTitle.trim()) {
      const titleText = '（' + evidenceTitle.trim() + '）';
      const titleImg = renderTextToImageBytes(titleText, 16, { bold: false });
      pdfTitleImage = await pdfDoc.embedPng(titleImg.bytes);
      titleImgDims = { width: titleImg.width, height: titleImg.height };
    }

    const margin = 36;
    const topMargin = 36;
    const pageCount = pdfDoc.getPageCount();
    const pagesToStamp = allPages
      ? Array.from({ length: pageCount }, (_, i) => i)
      : [0];

    for (const pageIndex of pagesToStamp) {
      const page = pdfDoc.getPage(pageIndex);
      const { width: pgW, height: pgH } = page.getSize();

      const labelX = pgW - margin - labelImg.width;
      const labelY = pgH - topMargin - labelImg.height;

      page.drawImage(pdfLabelImage, {
        x: labelX, y: labelY,
        width: labelImg.width, height: labelImg.height,
      });

      if (pdfTitleImage && titleImgDims) {
        page.drawImage(pdfTitleImage, {
          x: pgW - margin - titleImgDims.width,
          y: labelY - titleImgDims.height - 4,
          width: titleImgDims.width, height: titleImgDims.height,
        });
      }
    }

    const savedBytes = await pdfDoc.save();
    const blob = new Blob([savedBytes], { type: 'application/pdf' });
    const baseName = file.name.replace(/\.pdf$/i, '');
    const fileName = `${baseName}_${evidenceLabel}.pdf`;
    return { blob, fileName };
  }

  /**
   * 証拠説明書をExcel互換HTML形式(.xls)で生成する。
   * @param {Array} entries - [{label, title, originalOrCopy, createdDate, author, purpose}]
   * @param {string} partyRole - "原告" or "被告" 等
   * @returns {Blob}
   */
  function generateEvidenceSheet(entries, partyRole) {
    const rows = entries.map(e => `<tr>
      <td style="text-align:center;white-space:nowrap;">${e.label}</td>
      <td>${e.title || ''}</td>
      <td style="text-align:center;">${e.originalOrCopy || ''}</td>
      <td style="text-align:center;white-space:nowrap;">${e.createdDate || ''}</td>
      <td style="text-align:center;">${e.author || ''}</td>
      <td>${e.purpose || ''}</td>
    </tr>`).join('\n');

    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel">
<head><meta charset="UTF-8">
<style>
  table { border-collapse: collapse; font-family: 'ＭＳ 明朝','MS Mincho',serif; font-size: 11pt; width: 100%; }
  th, td { border: 1px solid #000; padding: 4px 8px; vertical-align: top; }
  th { background: #f0f0f0; font-weight: bold; text-align: center; }
  .title { font-size: 14pt; font-weight: bold; text-align: center; margin-bottom: 8px; }
  .party { text-align: right; margin-bottom: 4px; font-size: 11pt; }
</style>
</head><body>
<p class="title">証拠説明書</p>
<p class="party">${partyRole || ''}</p>
<table>
<tr><th style="width:12%;">号証</th><th style="width:20%;">標目</th><th style="width:10%;">原本・写し<br>の別</th><th style="width:12%;">作成日</th><th style="width:12%;">作成者</th><th style="width:34%;">立証趣旨</th></tr>
${rows}
</table>
</body></html>`;

    return new Blob(['\ufeff' + html], { type: 'application/vnd.ms-excel;charset=utf-8' });
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
    const pdfs = Array.from(fileList).filter(f => f.name.toLowerCase().endsWith('.pdf'));
    if (pdfs.length === 0) {
      showError('PDFファイルを選択してください。');
      return;
    }
    const oversized = pdfs.filter(f => f.size > 50 * 1024 * 1024);
    if (oversized.length > 0) {
      showError(`ファイルサイズが大きすぎます（上限: 50MB）: ${oversized.map(f=>f.name).join(', ')}`);
      return;
    }
    if (currentMode === 'receipt') {
      prepareReceiptFiles(pdfs);
    } else if (currentMode === 'evidence') {
      prepareEvidenceFiles(pdfs);
    } else {
      uploadFiles(pdfs);
    }
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
    processingTitle.textContent = 'PDFを解析中...';
    processingMessage.textContent = 'PDFを読み込んでいます...';
    startProcessingSteps('upload');

    try {
      const total = pdfs.length;
      const allResults = [];

      for (let i = 0; i < total; i++) {
        if (total > 1) {
          processingTitle.textContent = `PDFを解析中... (${i + 1}/${total})`;
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
  }

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
      uploadHeading.textContent = 'FAX送信書のPDFをドロップ';
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
  const evidenceNumberInput = $('#evidenceNumber');
  const evidenceSubNumInput = $('#evidenceSubNum');
  const evidenceUseDai = $('#evidenceUseDai');
  const evidenceTitleInput = $('#evidenceTitle');
  const evidenceOriginalCopy = $('#evidenceOriginalCopy');
  const evidenceCreatedDate = $('#evidenceCreatedDate');
  const evidenceAuthor = $('#evidenceAuthor');
  const evidencePurpose = $('#evidencePurpose');
  const evidenceAllPages = $('#evidenceAllPages');
  const evidenceOutputSheet = $('#evidenceOutputSheet');
  const evidencePreview = $('#evidencePreview');
  const evidenceSourceFileName = $('#evidenceSourceFileName');
  const btnEvidenceBack = $('#btnEvidenceBack');
  const btnEvidenceGenerate = $('#btnEvidenceGenerate');
  const evidenceFileCountBadge = $('#evidenceFileCountBadge');
  const evidenceFileListWrap = $('#evidenceFileListWrap');
  const evidenceFileList = $('#evidenceFileList');

  function getEvidencePreviewLabel() {
    const party = evidencePartySelect ? evidencePartySelect.value : '甲';
    const num = evidenceNumberInput ? evidenceNumberInput.value.trim() : '';
    const subNum = evidenceSubNumInput ? evidenceSubNumInput.value.trim() : '';
    const useDai = evidenceUseDai ? evidenceUseDai.checked : false;
    if (!num) return party + (useDai ? '第' : '') + '○号証';
    return buildEvidenceLabel(party, parseInt(num, 10), subNum ? parseInt(subNum, 10) : null, useDai);
  }

  function updateEvidencePreview() {
    if (evidencePreview) {
      const label = getEvidencePreviewLabel();
      evidencePreview.textContent = label;
      const title = evidenceTitleInput ? evidenceTitleInput.value.trim() : '';
      if (title) {
        evidencePreview.textContent = label + '\n（' + title + '）';
      }
    }
  }

  [evidencePartySelect, evidenceNumberInput, evidenceSubNumInput, evidenceTitleInput].forEach(el => {
    if (el) el.addEventListener('input', updateEvidencePreview);
    if (el) el.addEventListener('change', updateEvidencePreview);
  });
  if (evidenceUseDai) evidenceUseDai.addEventListener('change', updateEvidencePreview);

  function prepareEvidenceFiles(pdfs) {
    evidenceUploadFiles = pdfs;
    if (evidenceSourceFileName) {
      evidenceSourceFileName.textContent = pdfs.length === 1
        ? pdfs[0].name
        : `${pdfs.length}件のPDFを処理します`;
    }
    if (evidenceFileCountBadge) {
      evidenceFileCountBadge.textContent = pdfs.length > 1 ? `${pdfs.length}件` : '証拠番号モード';
    }
    if (evidenceFileListWrap && evidenceFileList) {
      if (pdfs.length > 1) {
        evidenceFileList.innerHTML = pdfs.map(f => `<li>${f.name}</li>`).join('');
        evidenceFileListWrap.style.display = 'block';
      } else {
        evidenceFileListWrap.style.display = 'none';
      }
    }
    updateEvidencePreview();
    setState('evidence-confirm');
  }

  if (btnEvidenceBack) {
    btnEvidenceBack.addEventListener('click', () => {
      evidenceUploadFiles = [];
      setState('upload');
    });
  }

  if (btnEvidenceGenerate) {
    btnEvidenceGenerate.addEventListener('click', async () => {
      if (!evidenceUploadFiles || evidenceUploadFiles.length === 0) {
        showError('ファイルが選択されていません。');
        setState('upload');
        return;
      }
      const party = evidencePartySelect ? evidencePartySelect.value : '甲';
      const numStr = evidenceNumberInput ? evidenceNumberInput.value.trim() : '';
      if (!numStr) {
        showError('証拠番号を入力してください。');
        return;
      }
      const startNum = parseInt(numStr, 10);
      if (isNaN(startNum) || startNum < 1) {
        showError('証拠番号は1以上の数字を入力してください。');
        return;
      }
      const subNumStr = evidenceSubNumInput ? evidenceSubNumInput.value.trim() : '';
      const subNum = subNumStr ? parseInt(subNumStr, 10) : null;
      const useDai = evidenceUseDai ? evidenceUseDai.checked : false;
      const titleStr = evidenceTitleInput ? evidenceTitleInput.value.trim() : '';
      const allPages = evidenceAllPages ? evidenceAllPages.checked : false;
      const outputSheet = evidenceOutputSheet ? evidenceOutputSheet.checked : false;
      const originalCopy = evidenceOriginalCopy ? evidenceOriginalCopy.value : '';
      const createdDate = evidenceCreatedDate ? evidenceCreatedDate.value.trim() : '';
      const author = evidenceAuthor ? evidenceAuthor.value.trim() : '';
      const purpose = evidencePurpose ? evidencePurpose.value.trim() : '';

      const files = evidenceUploadFiles;
      const total = files.length;

      setState('processing');
      startProcessingSteps('generate');
      const results = [];
      const sheetEntries = [];

      for (let i = 0; i < total; i++) {
        const currentNum = startNum + i;
        const currentSub = (total === 1 && subNum) ? subNum : (total > 1 ? null : subNum);
        const evidenceLabel = buildEvidenceLabel(party, currentNum, currentSub, useDai);
        processingTitle.textContent = total > 1
          ? `証拠番号を書き込み中... (${i + 1}/${total})`
          : '証拠番号を書き込み中...';
        processingMessage.textContent = `${files[i].name} → ${evidenceLabel}`;

        try {
          const result = await generateEvidenceBrowser(files[i], {
            evidenceLabel, evidenceTitle: titleStr, allPages,
          });
          const blobUrl = URL.createObjectURL(result.blob);
          results.push({ fileName: result.fileName, downloadUrl: blobUrl, error: null });
          sheetEntries.push({
            label: evidenceLabel,
            title: titleStr,
            originalOrCopy: originalCopy,
            createdDate: createdDate,
            author: author,
            purpose: purpose,
          });
        } catch (err) {
          results.push({ fileName: files[i].name, downloadUrl: null, error: err.message || '生成失敗' });
        }
      }

      // 証拠説明書の生成
      if (outputSheet && sheetEntries.length > 0) {
        const partyLabel = party.startsWith('甲') ? '原告' : party.startsWith('乙') ? '被告' : '';
        const sheetBlob = generateEvidenceSheet(sheetEntries, partyLabel);
        const sheetUrl = URL.createObjectURL(sheetBlob);
        results.push({
          fileName: '証拠説明書.xls',
          downloadUrl: sheetUrl,
          error: null,
          isSheet: true,
        });
      }

      resetProcessingSteps();
      [procStep1, procStep2, procStep3].forEach(step => { if (step) step.classList.add('done'); });
      const succeeded = results.filter(r => !r.error);
      const failed = results.filter(r => r.error);
      const pdfCount = succeeded.filter(r => !r.isSheet).length;
      completeTitle.textContent = total === 1
        ? '証拠番号の書き込みが完了しました'
        : `証拠番号の書き込みが完了しました（${pdfCount}/${total}件）`;
      outputFileName.textContent = '';

      if (succeeded.length === 1 && !outputSheet) {
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
          const bgColor = r.isSheet ? '#eff6ff' : '#f0fdf4';
          const icon = r.isSheet ? '📋' : '📄';
          return `<li style="display:flex;align-items:center;gap:8px;padding:8px 12px;background:${bgColor};border-radius:8px;">
            <span style="font-size:1.1em;">${icon}</span>
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
