/**
 * 文書送付書・受領書 自動でつくる君 - ブラウザ版 UI ロジック
 *
 * サーバーAPIへの fetch を全て削除し、
 * receipt-browser.js / generate-browser.js の関数を直接呼び出す。
 */

(function() {
  'use strict';

  // ===== ステート管理 =====
  let currentState = 'upload';
  let currentMode = 'sofusho';
  let receiptUploadFiles = [];

  // ===== DOM参照 =====
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  const states = {
    upload: $('#state-upload'),
    processing: $('#state-processing'),
    confirm: $('#state-confirm'),
    'receipt-confirm': $('#state-receipt-confirm'),
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

  // 処理ステップ
  const procStep1 = $('#procStep1');
  const procStep2 = $('#procStep2');
  const procStep3 = $('#procStep3');

  // フォームフィールド
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

  // モード関連
  const modeSofushoBtn = $('#modeSofusho');
  const modeReceiptBtn = $('#modeReceipt');
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

  // 受領書モード用DOM
  const receiptSourceFileName = $('#receiptSourceFileName');
  const receiptSignerTitle = $('#receiptSignerTitle');
  const receiptSignerName = $('#receiptSignerName');
  const receiptDateInput = $('#receiptDate');
  const btnReceiptBack = $('#btnReceiptBack');
  const btnReceiptGenerate = $('#btnReceiptGenerate');

  // ===== ステート切り替え =====
  function setState(newState) {
    states[currentState].classList.remove('active');

    const stateOrder = ['upload', 'processing', 'confirm', 'complete'];
    const mapState = newState === 'receipt-confirm' ? 'confirm' : newState;
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

  // ===== エラー表示 =====
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

  // ===== 処理ステップのアニメーション =====
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

  // ===== 進捗更新 =====
  function updateProgress(msg) {
    if (processingMessage) processingMessage.textContent = msg;

    // ステップの自動進行
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

  // ===== 紙吹雪エフェクト =====
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

  // ===== ファイルアップロード =====
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
    } else {
      uploadFile(pdfs[0]);
    }
  }

  // ===== ドラッグ＆ドロップ =====
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

  // ===== 文書送付書モード: PDFアップロード → 情報抽出 =====
  async function uploadFile(file) {
    setState('processing');
    processingTitle.textContent = 'PDFを解析中...';
    processingMessage.textContent = 'OCRで文字を認識しています。しばらくお待ちください。';
    startProcessingSteps('upload');

    try {
      // ★ ブラウザ版: サーバーAPI不要、直接関数呼び出し
      const result = await window.uploadAndExtractBrowser(file, updateProgress);

      // 全ステップ完了
      resetProcessingSteps();
      [procStep1, procStep2, procStep3].forEach(step => {
        if (step) step.classList.add('done');
      });

      // フォームにデータを流し込み
      populateForm(result.info, result.documentTitle, result.originalName);

      await new Promise(resolve => setTimeout(resolve, 500));
      setState('confirm');

    } catch (err) {
      resetProcessingSteps();
      showError(err.message || 'PDF解析中にエラーが発生しました。');
      setState('upload');
    }
  }

  // ===== フォームにデータ流し込み =====
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

    // ステータスバッジ更新
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

    // 未検出フィールドのハイライト
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

  // ===== 戻るボタン =====
  btnBack.addEventListener('click', () => {
    setState('upload');
  });

  // ===== 生成ボタン（文書送付書）=====
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
      // ★ ブラウザ版: 直接Word生成
      const result = await window.generateDocumentBrowser(info, documentTitle, updateProgress);

      // Blob URLを生成
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

  // ===== 新規ファイルボタン =====
  btnNewFile.addEventListener('click', () => {
    receiptUploadFiles = [];
    completeTitle.textContent = '文書送付書の生成が完了しました';
    downloadLabel.textContent = 'Wordファイルをダウンロード';
    setState('upload');
  });

  // ===== モード切替 =====
  function switchMode(mode) {
    currentMode = mode;
    modeSofushoBtn.classList.toggle('active', mode === 'sofusho');
    modeReceiptBtn.classList.toggle('active', mode === 'receipt');

    if (mode === 'receipt') {
      appTitle.textContent = '受領書自動でつくる君';
      logoIcon.classList.add('receipt-mode');
      uploadHeading.textContent = '相手方の文書送付書PDFをドロップ';
      uploadDesc.textContent = '受領日・署名・押印を自動で書き込みます';
    } else {
      appTitle.textContent = '文書送付書自動でつくる君';
      logoIcon.classList.remove('receipt-mode');
      uploadHeading.textContent = 'FAX送信書のPDFをドロップ';
      uploadDesc.textContent = 'ファイルをここにドラッグ＆ドロップしてください';
    }

    receiptUploadFiles = [];
    setState('upload');
  }

  modeSofushoBtn.addEventListener('click', () => switchMode('sofusho'));
  modeReceiptBtn.addEventListener('click', () => switchMode('receipt'));

  // ===== 受領書モード: 複数ファイルを確認画面に表示 =====
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

  // ===== 受領書：戻るボタン =====
  btnReceiptBack.addEventListener('click', () => {
    receiptUploadFiles = [];
    setState('upload');
  });

  // ===== 受領書：生成ボタン（複数ファイル対応）=====
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
        // ★ ブラウザ版: 直接関数呼び出し
        const result = await window.generateReceiptBrowser(file, {
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

    // 完了画面へ
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

  // ===== 設定モーダル =====
  if (settingsBtn && settingsModal) {
    settingsBtn.addEventListener('click', () => {
      // 現在の設定を読み込み
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

      // 印鑑プレビュー
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

    // 印鑑画像アップロード
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

    // 印鑑削除ボタン
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

    // 設定保存
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

      // UI更新
      const subtitle = $('#officeSubtitle');
      if (subtitle) subtitle.textContent = config.officeName;
      if (receiptSignerName && config.signerName) {
        receiptSignerName.value = config.signerName;
      }

      settingsModal.classList.remove('visible');
    });

    // モーダル外クリックで閉じる
    settingsModal.addEventListener('click', (e) => {
      if (e.target === settingsModal) {
        settingsModal.classList.remove('visible');
      }
    });
  }

  // ===== 起動時に設定を読み込み =====
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

  // ===== ページ離脱前の確認 =====
  window.addEventListener('beforeunload', (e) => {
    if (currentState === 'confirm' || currentState === 'receipt-confirm') {
      e.preventDefault();
      e.returnValue = '';
    }
  });

})();
