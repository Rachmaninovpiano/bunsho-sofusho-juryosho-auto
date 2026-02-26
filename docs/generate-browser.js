/**
 * 文書送付書自動でつくる君 - ブラウザ版文書送付書生成モジュール
 *
 * pdf.js + Tesseract.js + JSZip で全てブラウザ内で処理
 * サーバー不要・インストール不要
 */

// ===== 設定（localStorageから取得）=====
function getGenerateConfig() {
  try {
    return JSON.parse(localStorage.getItem('tsukurukun_config') || '{}');
  } catch (e) { return {}; }
}

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

// ===== pdf.js でPDFからテキスト抽出 =====
async function extractTextFromPDFBrowser(pdfArrayBuffer, onProgress) {
  onProgress && onProgress('PDFからテキストを抽出中...');

  const pdfDoc = await pdfjsLib.getDocument({ data: pdfArrayBuffer }).promise;
  const totalPages = pdfDoc.numPages;
  let text = '';

  for (let i = 1; i <= totalPages; i++) {
    const page = await pdfDoc.getPage(i);
    const content = await page.getTextContent();
    for (const item of content.items) {
      if (item.str) text += item.str;
      if (item.hasEOL) text += '\n';
    }
    text += '\n';
  }

  return text;
}

// ===== Tesseract.js でOCRテキスト抽出 =====
async function extractTextWithOCRBrowser(pdfArrayBuffer, onProgress) {
  onProgress && onProgress('画像PDFを検出。OCRで文字認識中...');

  const pdfDoc = await pdfjsLib.getDocument({ data: pdfArrayBuffer }).promise;
  const totalPages = Math.min(pdfDoc.numPages, 3); // 最大3ページ

  let allText = '';

  for (let i = 1; i <= totalPages; i++) {
    onProgress && onProgress(`ページ ${i}/${totalPages} をOCR中...`);

    const page = await pdfDoc.getPage(i);
    const viewport = page.getViewport({ scale: 300 / 72 });

    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');

    await page.render({ canvasContext: ctx, viewport }).promise;

    // Tesseract.js でOCR
    const worker = await Tesseract.createWorker('jpn', 1, {
      logger: m => {
        if (m.status === 'recognizing text' && onProgress) {
          onProgress(`ページ${i} OCR処理中... ${Math.round((m.progress || 0) * 100)}%`);
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

// ===== PDFからテキスト抽出（テキスト埋め込みなし→OCRフォールバック）=====
async function extractTextBrowser(pdfArrayBuffer, onProgress) {
  const text = await extractTextFromPDFBrowser(pdfArrayBuffer, onProgress);

  // テキストが実質空（空白・改行のみ）ならOCR
  const trimmed = text.replace(/[\s\n\r]/g, '');
  if (trimmed.length < 10) {
    console.log('  テキスト埋め込みなしPDF → OCRに切り替え');
    return await extractTextWithOCRBrowser(pdfArrayBuffer, onProgress);
  }

  return text;
}

// ===== 情報抽出（サーバー版と同一ロジック）=====
function extractInfoFromText(text) {
  const config = getGenerateConfig();
  const info = {};

  const normalizedText = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');

  // 都市名リスト（裁判所パターン用）
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
    /損害\s*賠償[\s\S]{0,50}?請求\s*事件/,
    /(?:号\s*)([\u4e00-\u9fff]+(?:請求|確認|等?)\s*事件)/,
    /(損害賠償請求事件|貸金返還請求事件|建物明渡請求事件|不当利得返還請求事件|(?:[\u4e00-\u9fff]+請求事件))/,
    /([\u4e00-\u9fff]+\s+(?:請求|確認)\s*事件)/,
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
  const partySection = normalizedText.match(
    /当\s*事\s*者[\s\S]{0,200}/
  );

  if (partySection) {
    const partySectionText = partySection[0];

    const plaintiffInParty = partySectionText.match(
      /原\s*[告&]\s*[_\s]*([^\n原被]{1,40})/
    );
    if (plaintiffInParty) {
      let name = plaintiffInParty[1].trim();
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

  // フォールバック: 当事者セクションで見つからなかった場合
  if (!info.plaintiffName) {
    const plaintiffPatterns = [
      /[【\[]\s*原\s*告\s*[】\]]\s*([^\n【\[]{1,30})/,
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

  // パターン1: 原告訴訟代理人弁護士
  const formalPattern = /原告\s*(?:ら)?\s*(?:訴\s*訟)?\s*代理\s*人\s*弁護\s*士\s*([^\n]{2,20})/g;
  let lm;
  while ((lm = formalPattern.exec(normalizedText)) !== null) {
    const name = cleanLawyerName(lm[1]);
    if (name) lawyerCandidates.push({ name, priority: 1 });
  }

  // パターン2: 人 弁護士
  const senderPattern = /人\s*弁護\s*士\s*([^\n]{2,20})/g;
  while ((lm = senderPattern.exec(normalizedText)) !== null) {
    const contextBefore = normalizedText.substring(Math.max(0, lm.index - 30), lm.index);
    if (contextBefore.includes('被告')) continue;
    const name = cleanLawyerName(lm[1]);
    if (name) lawyerCandidates.push({ name, priority: 2 });
  }

  // パターン3: 弁護士 NAME 宛て
  const atePattern = /弁護\s*士\s*([^\n]{2,15})\s*宛/g;
  while ((lm = atePattern.exec(normalizedText)) !== null) {
    const name = cleanLawyerName(lm[1]);
    if (name) lawyerCandidates.push({ name, priority: 3 });
  }

  // パターン4: 弁護士 NAME
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
      .replace(/[－ー]/g, '-');
  }

  // 明示的裁判所FAX
  const explicitCourtFaxMatch = normalizedText.match(
    /裁\s*判\s*所[\s\S]{0,60}?[（(]\s*(?:FAX|ＦＡＸ|[Ff]ax)\s*([0-9０-９\-－ー]+)\s*[）)]/
  );
  if (explicitCourtFaxMatch) {
    info.courtFaxFromPdf = normalizeFax(explicitCourtFaxMatch[1]);
  }

  // 明示的ラベル付きFAX
  const allExplicitFaxes = [];
  const explicitFaxRegex = /([\u4e00-\u9fff]{1,10})\s*[（(]\s*(?:FAX|ＦＡＸ|[Ff]ax)\s*([0-9０-９\-－ー]+)\s*[）)]/g;
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

  // 通常FAX番号抽出
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

  // 裁判所FAX: PDFから取れたら優先、なければ辞書
  if (info.courtFaxFromPdf) {
    info.courtFax = info.courtFaxFromPdf;
  }

  return info;
}

// ===== XML安全テキスト置換（サーバー版と同一）=====
function safeReplaceInXml(xml, oldText, newText) {
  const paraRegex = /(<w:p[\s>][\s\S]*?<\/w:p>)/g;
  return xml.replace(paraRegex, (paraXml) => {
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

      if (!affectedSegs.includes(seg)) {
        return match;
      }

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

// ===== テンプレートへの差し込み処理（サーバー版と同一）=====
function applyInfoToTemplate(docXml, info, documentTitle) {
  const today = getTodayReiwa();

  // --- 裁判所名 ---
  if (info.courtName) {
    const ORIG_COURT = '神戸地方裁判所尼崎支部第２民事部';
    const courtDiff = ORIG_COURT.length - info.courtName.length;
    const courtPad = courtDiff > 0 ? '\u3000'.repeat(courtDiff) : '';
    docXml = safeReplaceInXml(docXml, ORIG_COURT, info.courtName + courtPad);
  }

  // --- 裁判所FAX番号 ---
  if (info.courtFax) {
    docXml = safeReplaceInXml(docXml, '06-6438-1710', info.courtFax);
    const fullWidthFax = toFullWidthNumber(info.courtFax).replace(/-/g, '\uFF0D');
    docXml = safeReplaceInXml(docXml, '\uFF10\uFF16\u2015\uFF16\uFF14\uFF13\uFF18\uFF0D\uFF11\uFF17\uFF11\uFF10', fullWidthFax);
  }

  // --- 原告代理人弁護士名 ---
  if (info.plaintiffLawyer) {
    const ORIG_LAWYER = '四方久寛';
    const lawyerDiff = ORIG_LAWYER.length - info.plaintiffLawyer.length;
    const lawyerPad = lawyerDiff > 0 ? '\u3000'.repeat(lawyerDiff) : '';
    docXml = safeReplaceInXml(docXml, ORIG_LAWYER, info.plaintiffLawyer + lawyerPad);
  }

  // --- 原告代理人FAX番号 ---
  if (info.plaintiffLawyerFax) {
    docXml = safeReplaceInXml(docXml, '06-4708-3638', info.plaintiffLawyerFax);
  }

  // --- 日付（送付書）---
  docXml = safeReplaceInXml(docXml, '令和6年11月7日', `令和${today.year}年${today.month}月${today.day}日`);

  // --- 日付（受領証明書）---
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

// ===== ファイル名から送付書類名を自動取得 =====
function getDocumentTitleFromFilename(fileName) {
  let baseName = fileName.replace(/\.pdf$/i, '');
  baseName = baseName.replace(/^【[^】]+】\s*/, '');
  baseName = baseName.replace(/^[\u4e00-\u9fff]+事案[\s\u3000]+/, '');
  return baseName;
}

// ===== メイン: PDFをアップロードして情報抽出 =====

/**
 * PDFファイルから裁判情報を抽出
 * @param {File} file - PDFファイル（File API）
 * @param {Function} onProgress - 進捗コールバック (msg)
 * @returns {Promise<{ info: Object, documentTitle: string, originalName: string }>}
 */
async function uploadAndExtractBrowser(file, onProgress) {
  console.log('文書送付書: PDF解析開始:', file.name);
  onProgress && onProgress('PDFを読み込み中...');

  const pdfArrayBuffer = await file.arrayBuffer();

  // テキスト抽出（テキスト埋め込み→OCRフォールバック）
  const pdfText = await extractTextBrowser(pdfArrayBuffer, onProgress);

  console.log('--- PDF抽出テキスト (先頭500文字) ---');
  console.log(pdfText.substring(0, 500));

  // 情報抽出
  onProgress && onProgress('情報を抽出中...');
  const info = extractInfoFromText(pdfText);
  console.log('抽出結果:', JSON.stringify(info, null, 2));

  // 送付書類名（ファイル名から）
  const documentTitle = getDocumentTitleFromFilename(file.name);

  return {
    info,
    documentTitle,
    originalName: file.name,
  };
}

// ===== メイン: 情報からWord文書を生成 =====

/**
 * ユーザー編集済み情報からWord文書を生成
 * @param {Object} info - 裁判情報オブジェクト
 * @param {string} documentTitle - 送付書類名
 * @param {Function} onProgress - 進捗コールバック (msg)
 * @returns {Promise<{ blob: Blob, fileName: string }>}
 */
async function generateDocumentBrowser(info, documentTitle, onProgress) {
  console.log('文書送付書: Word生成開始');
  onProgress && onProgress('テンプレートを読み込み中...');

  // Wordテンプレートを取得
  const templateResp = await fetch('template/文書送付書.doc.docx');
  if (!templateResp.ok) {
    throw new Error('テンプレートファイルの読み込みに失敗しました。');
  }
  const templateData = await templateResp.arrayBuffer();

  // JSZipでテンプレートを展開
  onProgress && onProgress('テンプレートにデータを差し込み中...');
  const zip = await JSZip.loadAsync(templateData);
  let docXml = await zip.file('word/document.xml').async('string');

  // データ差し込み
  docXml = applyInfoToTemplate(docXml, info, documentTitle);
  zip.file('word/document.xml', docXml);

  // Word出力
  onProgress && onProgress('Wordファイルを生成中...');
  const outputBlob = await zip.generateAsync({ type: 'blob' });

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
  const outputFileName = `文書送付書_${documentTitle}_${timestamp}.docx`;

  console.log('文書送付書生成完了:', outputFileName);

  return {
    blob: outputBlob,
    fileName: outputFileName,
  };
}

// グローバルに公開
window.uploadAndExtractBrowser = uploadAndExtractBrowser;
window.generateDocumentBrowser = generateDocumentBrowser;
window.extractInfoFromText = extractInfoFromText;
window.COURT_FAX_MAP = COURT_FAX_MAP;
