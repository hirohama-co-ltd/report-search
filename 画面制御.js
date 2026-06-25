// ========================================
// 製品PDF検索 — 画面制御
// ========================================

function doGet(e) {
  e = e || {};
  var page = e.parameter && e.parameter.page;
  var fileId = e.parameter && e.parameter.fileId;

  if (page === 'view') return servePdfView_(fileId);
  if (page === 'download') return servePdfDownload_(fileId);

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('検査成績書検索')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPdfAppInitialData() {
  return {
    scopeOptions: [
      { value: PDF_SCOPE_EXTERNAL, label: PDF_SCOPE_LABELS.external },
      { value: PDF_SCOPE_INTERNAL, label: PDF_SCOPE_LABELS.internal }
    ],
    appBaseUrl: getPdfAppBaseUrl_(),
    maxResults: PDF_SEARCH_MAX_RESULTS,
    previewMaxBytes: PDF_PREVIEW_MAX_BYTES,
    chunkSize: PDF_CHUNK_SIZE_BYTES,
    chunkBatchMax: PDF_CHUNK_BATCH_MAX
  };
}

function servePdfView_(fileId) {
  return servePdfBlob_(fileId, false);
}

function servePdfDownload_(fileId) {
  return servePdfBlob_(fileId, true);
}

function servePdfBlob_(fileId, asDownload) {
  if (!isFileUnderPdfRoots_(fileId)) {
    return ContentService.createTextOutput('PDF not found or access denied')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var output = ContentService.create(blob).setMimeType('application/pdf');
    if (asDownload) {
      output.downloadAs(file.getName());
    }
    return output;
  } catch (e) {
    Logger.log('servePdfBlob_: ' + e.message);
    return ContentService.createTextOutput('Error: ' + e.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * プレビュー用メタ情報
 */
function getPdfFileMeta_(fileId) {
  fileId = String(fileId || '').trim();
  if (!fileId) throw new Error('ファイルIDが未指定です');
  if (!isFileUnderPdfRoots_(fileId)) {
    throw new Error('PDFが見つからないか、アクセスできません');
  }

  var file = DriveApp.getFileById(fileId);
  var size = file.getSize();
  if (size > PDF_PREVIEW_MAX_BYTES) {
    throw new Error(
      'ファイルが大きすぎてプレビューできません（'
      + Math.round(size / 1024 / 1024 * 10) / 10 + 'MB）。ダウンロードボタンをご利用ください。'
    );
  }

  return {
    fileId: fileId,
    fileName: file.getName(),
    mimeType: 'application/pdf',
    size: size,
    chunkSize: PDF_CHUNK_SIZE_BYTES
  };
}

/**
 * Drive API Range 取得で PDF を分割返却（google.script.run 用）
 */
function getPdfBytesChunk_(payload) {
  payload = payload || {};
  var fileId = String(payload.fileId || '').trim();
  var offset = Number(payload.offset) || 0;
  if (!fileId) throw new Error('ファイルIDが未指定です');
  if (!isFileUnderPdfRoots_(fileId)) {
    throw new Error('PDFが見つからないか、アクセスできません');
  }

  var file = DriveApp.getFileById(fileId);
  var total = file.getSize();
  if (offset < 0 || offset >= total) {
    throw new Error('読込位置が不正です');
  }

  var chunkSize = PDF_CHUNK_SIZE_BYTES;
  var end = Math.min(offset + chunkSize - 1, total - 1);
  var resp = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId)
      + '?alt=media&supportsAllDrives=true',
    {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        Range: 'bytes=' + offset + '-' + end
      },
      muteHttpExceptions: true
    }
  );

  var code = resp.getResponseCode();
  if (code !== 206 && code !== 200) {
    throw new Error('DriveからPDFを取得できません (HTTP ' + code + ')');
  }

  var content = resp.getContent();
  return {
    offset: offset,
    length: content.length,
    total: total,
    base64: Utilities.base64Encode(content)
  };
}

/**
 * 複数チャンクを1回の google.script.run で返却
 */
function getPdfBytesChunks_(payload) {
  payload = payload || {};
  var fileId = String(payload.fileId || '').trim();
  var offsets = payload.offsets || [];
  if (!fileId) throw new Error('ファイルIDが未指定です');
  if (!offsets.length) throw new Error('読込位置が未指定です');
  if (offsets.length > PDF_CHUNK_BATCH_MAX) {
    throw new Error('一度に取得できるチャンク数を超えています');
  }

  var chunks = [];
  for (var i = 0; i < offsets.length; i++) {
    chunks.push(getPdfBytesChunk_({
      fileId: fileId,
      offset: Number(offsets[i]) || 0
    }));
  }
  return { chunks: chunks };
}

/** クライアント公開API（末尾 _ 付きは google.script.run から呼べない） */
function getPdfFileMeta(fileId) {
  return getPdfFileMeta_(fileId);
}

function getPdfBytesChunk(payload) {
  return getPdfBytesChunk_(payload);
}

function getPdfBytesChunks(payload) {
  return getPdfBytesChunks_(payload);
}

/** @deprecated 大容量PDFは getPdfBytesChunk を使用 */
function getPdfPreviewPayload_(fileId) {
  fileId = String(fileId || '').trim();
  if (!fileId) throw new Error('ファイルIDが未指定です');
  if (!isFileUnderPdfRoots_(fileId)) {
    throw new Error('PDFが見つからないか、アクセスできません');
  }

  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var bytes = blob.getBytes();
  if (bytes.length > PDF_PREVIEW_MAX_BYTES) {
    throw new Error(
      'ファイルが大きすぎてプレビューできません（'
      + Math.round(bytes.length / 1024 / 1024 * 10) / 10 + 'MB）。ダウンロードボタンをご利用ください。'
    );
  }

  return {
    fileId: fileId,
    fileName: file.getName(),
    mimeType: blob.getContentType() || 'application/pdf',
    size: bytes.length,
    base64: Utilities.base64Encode(bytes)
  };
}

function runPdfSearchFromMenu_() {
  var result = searchProductPdfs({
    scope: PDF_SCOPE_EXTERNAL,
    productCode: '',
    lotNo: '',
    mfgDate: ''
  });
  return '社外テスト検索: ' + result.count + '件'
    + (result.truncated ? '（上限到達）' : '');
}
