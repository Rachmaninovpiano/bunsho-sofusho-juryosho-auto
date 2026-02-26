/**
 * 文書送付書 自動生成 Web UI - フロントエンドロジック
 */

(function() {
  'use strict';

  // ===== ステート管理 =====
  let currentState = 'upload'; // upload | processing | confirm | receipt-confirm | complete
  let currentMode = 'sofusho'; // sofusho | receipt
  let sessionId = null;
  let processingTimers = [];
  let receiptUploadFiles = []; // 受領書モードで一時保存するファイル群

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

  // ===== ステート切り替え =====
  function setState(newState) {
    // 前のステートを非表示
    states[currentState].classList.remove('active');

    // ステップインジケーター更新
    const stateOrder = ['upload', 'processing', 'confirm', 'complete'];
    // receipt-confirmはconfirmと同じステップ位置(インデックス2)
    const mapState = newState === 'receipt-confirm' ? 'confirm' : newState;
    const newIndex = stateOrder.indexOf(mapState);

    const stepElements = $$('.step');
    const connectorFills = $$('.step-connector-fill');

    stepElements.forEach((step, i) => {
      step.classList.remove('active', 'completed');
      if (i < newIndex) step.classList.add('completed');
      if (i === newIndex) step.classList.add('active');
    });

    // コネクターフィルのアニメーション
    connectorFills.forEach((fill, i) => {
      if (i < newIndex) {
        fill.style.width = '100%';
      } else {
        fill.style.width = '0%';
      }
    });

    // 新しいステートを表示
    currentState = newState;
    states[currentState].classList.add('active');

    // スクロールトップ
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
  function resetProcessingSteps() {
    // 全タイマーをクリア
    processingTimers.forEach(t => clearTimeout(t));
    processingTimers = [];

    // ステップリセット
    [procStep1, procStep2, procStep3].forEach(step => {
      if (step) {
        step.classList.remove('active', 'done');
      }
    });
  }

  function startProcessingSteps(mode) {
    resetProcessingSteps();

    if (mode === 'upload') {
      // PDF解析モード: 段階的にステップを進める
      if (procStep1) procStep1.classList.add('active');

      processingTimers.push(setTimeout(() => {
        if (procStep1) { procStep1.classList.remove('active'); procStep1.classList.add('done'); }
        if (procStep2) procStep2.classList.add('active');
      }, 3000));

      processingTimers.push(setTimeout(() => {
        if (procStep2) { procStep2.classList.remove('active'); procStep2.classList.add('done'); }
        if (procStep3) procStep3.classList.add('active');
      }, 8000));

    } else if (mode === 'generate') {
      // Word生成モード: 全ステップ完了にして非表示
      [procStep1, procStep2, procStep3].forEach(step => {
        if (step) step.classList.add('done');
      });
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
        {
          transform: `translateY(0) rotate(0deg) scale(1)`,
          opacity: 1,
        },
        {
          transform: `translateY(${Math.random() * 200 + 100}px) translateX(${(Math.random() - 0.5) * 150}px) rotate(${Math.random() * 720}deg) scale(0)`,
          opacity: 0,
        }
      ], {
        duration: duration,
        delay: delay,
        easing: 'cubic-bezier(0.25, 0.46, 0.45, 0.94)',
        fill: 'forwards',
      });

      confettiContainer.appendChild(piece);
    }

    // 片付け
    setTimeout(() => {
      if (confettiContainer) confettiContainer.innerHTML = '';
    }, 3000);
  }

  // ===== ファイルアップロード（複数対応）=====
  function handleFiles(fileList) {
    if (!fileList || fileList.length === 0) return;

    // PDFのみフィルタ
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

    // モードに応じて処理を分岐
    if (currentMode === 'receipt') {
      prepareReceiptFiles(pdfs);
    } else {
      // 文書送付書モードは1ファイルのみ
      uploadFile(pdfs[0]);
    }
  }

  // 後方互換
  function handleFile(file) { handleFiles(file ? [file] : []); }

  // ===== ドラッグ＆ドロップ =====
  // ★ 最重要: キャプチャフェーズ(第3引数=true)でページ全体のdragover/dropの
  //   デフォルト動作を完全にブロックする。これにより、どの要素にドロップされても
  //   ブラウザがPDFファイルを開こうとするナビゲーションを防止できる。
  document.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
  }, true);

  // ドラッグカウンター（子要素のdragenter/dragleaveでちらつかないようにする）
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

  // ★ dropハンドラー: キャプチャフェーズで処理（ブラウザのデフォルト動作を確実にブロック）
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter = 0;
    dropZone.classList.remove('dragover');
    if (dragOverlay) dragOverlay.classList.remove('visible');

    console.log('[つくる君] drop detected, state:', currentState,
      'files:', e.dataTransfer ? e.dataTransfer.files.length : 0);

    // アップロード画面のときだけファイルを処理する
    if (currentState === 'upload' && e.dataTransfer && e.dataTransfer.files.length > 0) {
      handleFiles(e.dataTransfer.files);
    }
  }, true);

  // ドロップゾーン全体クリック
  dropZone.addEventListener('click', (e) => {
    // ラベル / ボタンクリック時は除外（label が input を処理する）
    if (e.target.closest('.upload-btn-label')) return;
    if (e.target.closest('label')) return;
    fileInput.click();
  });

  // ファイル選択（複数対応）
  fileInput.addEventListener('change', () => {
    handleFiles(fileInput.files);
    fileInput.value = ''; // リセット
  });

  // ===== API: アップロード =====
  async function uploadFile(file) {
    setState('processing');
    processingTitle.textContent = 'PDFを解析中...';
    processingMessage.textContent = 'OCRで文字を認識しています。しばらくお待ちください。';
    startProcessingSteps('upload');

    const formData = new FormData();
    formData.append('pdf', file);

    try {
      // サーバー接続確認（素早くチェック）
      try {
        await fetch('/', { method: 'GET', signal: AbortSignal.timeout(3000) });
      } catch (_) {
        throw new Error('サーバーに接続できません。\n「文書送付書自動でつくる君.vbs」が起動しているか確認してください。');
      }

      const response = await fetch('/api/upload', {
        method: 'POST',
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'アップロードに失敗しました。');
      }

      // 最終ステップを完了にする
      resetProcessingSteps();
      [procStep1, procStep2, procStep3].forEach(step => {
        if (step) step.classList.add('done');
      });

      // セッションID保存
      sessionId = data.sessionId;

      // フォームにデータを流し込み
      populateForm(data.info, data.documentTitle, data.originalName);

      // 少し間を置いて確認画面へ（最終ステップの完了を見せる）
      await new Promise(resolve => setTimeout(resolve, 500));

      // 確認画面へ
      setState('confirm');

    } catch (err) {
      resetProcessingSteps();
      // ネットワークエラーの場合はわかりやすいメッセージに変換
      let msg = err.message;
      if (msg === 'Failed to fetch' || msg === 'NetworkError when attempting to fetch resource.') {
        msg = 'サーバーに接続できません。\n「文書送付書自動でつくる君.vbs」が起動しているか確認してください。';
      }
      showError(msg);
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

    // OCR推測警告
    if (info.caseNumberGuessed) {
      caseNumberWarning.hidden = false;
    } else {
      caseNumberWarning.hidden = true;
    }

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

    // 未検出フィールドのハイライト（CSSクラス使用）
    Object.entries(fields).forEach(([key, input]) => {
      if (!input.value) {
        input.classList.add('field-empty');
      } else {
        input.classList.remove('field-empty');
      }
    });

    // フォーカス時に空欄ハイライト解除
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

  // ===== 生成ボタン =====
  btnGenerate.addEventListener('click', async () => {
    // バリデーション（最低限のチェック）
    const requiredFields = ['courtName', 'courtFax', 'caseNumber', 'caseName',
                           'plaintiffName', 'defendantName', 'plaintiffLawyer',
                           'plaintiffLawyerFax', 'documentTitle'];
    const emptyFields = requiredFields.filter(key => !fields[key].value.trim());

    if (emptyFields.length > 0) {
      const fieldNames = {
        courtName: '裁判所名',
        courtFax: '裁判所FAX',
        caseNumber: '事件番号',
        caseName: '事件名',
        plaintiffName: '原告',
        defendantName: '被告',
        plaintiffLawyer: '原告代理人弁護士',
        plaintiffLawyerFax: '原告代理人FAX',
        documentTitle: '送付書類名',
      };
      const names = emptyFields.map(k => fieldNames[k]).join('、');

      // 空欄を警告して確認
      if (!confirm(`以下の項目が未入力です:\n${names}\n\n空欄のまま生成しますか？`)) {
        return;
      }
    }

    // 処理中表示
    setState('processing');
    processingTitle.textContent = '文書送付書を生成中...';
    processingMessage.textContent = 'テンプレートにデータを差し込んでいます。';
    startProcessingSteps('generate');

    // フォームからデータ収集
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
      const response = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sessionId,
          info,
          documentTitle,
        }),
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || '生成に失敗しました。');
      }

      // 完了画面へ
      completeTitle.textContent = '文書送付書の生成が完了しました';
      downloadLabel.textContent = 'Wordファイルをダウンロード';
      outputFileName.textContent = data.fileName;
      btnDownload.href = data.downloadUrl;
      setState('complete');

      // 紙吹雪エフェクト
      setTimeout(launchConfetti, 300);

    } catch (err) {
      resetProcessingSteps();
      showError(err.message);
      setState('confirm');
    }
  });

  // ===== 新規ファイルボタン =====
  btnNewFile.addEventListener('click', () => {
    sessionId = null;
    receiptUploadFiles = [];
    // 完了画面のテキストをリセット
    completeTitle.textContent = '文書送付書の生成が完了しました';
    downloadLabel.textContent = 'Wordファイルをダウンロード';
    setState('upload');
  });

  // ===== モード切替 =====
  const modeSofushoBtn = $('#modeSofusho');
  const modeReceiptBtn = $('#modeReceipt');
  const appTitle = $('#appTitle');
  const logoIcon = $('#logoIcon');
  const completeTitle = $('#completeTitle');
  const downloadLabel = $('#downloadLabel');
  const uploadHeading = $('.upload-heading');
  const uploadDesc = $('.upload-desc');

  // 受領書モード用DOM
  const receiptSourceFileName = $('#receiptSourceFileName');
  const receiptSignerTitle = $('#receiptSignerTitle');
  const receiptSignerName = $('#receiptSignerName');
  const receiptDateInput = $('#receiptDate');
  const btnReceiptBack = $('#btnReceiptBack');
  const btnReceiptGenerate = $('#btnReceiptGenerate');

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

    // 現在の画面をリセット
    sessionId = null;
    receiptUploadFiles = [];
    setState('upload');
  }

  modeSofushoBtn.addEventListener('click', () => switchMode('sofusho'));
  modeReceiptBtn.addEventListener('click', () => switchMode('receipt'));

  // ===== handleFileをモード対応に拡張 =====
  const originalHandleFile = handleFile;

  // handleFileを上書きしない代わりに、uploadFileでモード分岐する

  // ===== 受領書モード: 複数ファイルを確認画面に表示 =====
  function prepareReceiptFiles(pdfs) {
    receiptUploadFiles = pdfs;

    // ファイル数表示
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

    const results = []; // { fileName, downloadUrl, error }

    for (let i = 0; i < total; i++) {
      const file = files[i];
      processingTitle.textContent = total > 1
        ? `受領書を生成中... (${i + 1}/${total})`
        : '受領書を生成中...';
      processingMessage.textContent = `${file.name} — OCRで位置検出＆書き込み中`;

      const formData = new FormData();
      formData.append('pdf', file);
      formData.append('signerTitle', signerTitleVal);
      formData.append('signerName', signerNameVal);
      if (receiptDateVal) formData.append('receiptDate', receiptDateVal);

      try {
        const response = await fetch('/api/receipt', { method: 'POST', body: formData });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || '生成失敗');
        results.push({ fileName: data.fileName, downloadUrl: data.downloadUrl, error: null });
      } catch (err) {
        let msg = err.message;
        if (msg === 'Failed to fetch' || msg === 'NetworkError when attempting to fetch resource.') {
          msg = 'サーバー接続エラー';
        }
        results.push({ fileName: file.name, downloadUrl: null, error: msg });
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
      // 単一ファイル: 従来通りのボタン
      singleDownloadArea.style.display = '';
      multiDownloadArea.style.display = 'none';
      downloadLabel.textContent = 'PDFファイルをダウンロード';
      btnDownload.href = succeeded[0].downloadUrl;
      btnDownload.download = succeeded[0].fileName;
    } else {
      // 複数ファイル: リスト表示
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

  // ===== uploadFileを受領書モード対応にラップ =====
  const origUploadFile = uploadFile;
  // handleFileの中でuploadFileが呼ばれるので、uploadFileをオーバーライド
  // uploadFileはlet/constで宣言されているのでそのまま使う

  // ===== ページ離脱前の確認（確認画面の時のみ） =====
  window.addEventListener('beforeunload', (e) => {
    if (currentState === 'confirm' || currentState === 'receipt-confirm') {
      e.preventDefault();
      e.returnValue = '';
    }
  });

})();
