// ========================================
// 製品PDF検索 — 設定
// ========================================

var PDF_SCOPE_INTERNAL = 'internal';
var PDF_SCOPE_EXTERNAL = 'external';

/** 社内（原本 _original） / 社外（加工済）ルートフォルダID */
var PDF_ROOT_FOLDER_IDS = {
  internal: '1Pce4gcCvSGVxupCGgHWpxNUKHFNiSVWU',
  external: '1PElWA2VfabA6kw1J2Ta16yp6SFFCN3Tu'
};

var PDF_SCOPE_LABELS = {
  internal: '社内',
  external: '社外'
};

var PDF_SEARCH_MAX_RESULTS = 200;
var PDF_FOLDER_LIST_CACHE_TTL_SEC = 600;
/** プレビュー用の最大サイズ */
var PDF_PREVIEW_MAX_BYTES = 20 * 1024 * 1024;
/** 1回の google.script.run で返すチャンクサイズ */
var PDF_CHUNK_SIZE_BYTES = 512 * 1024;
/** 1リクエストでまとめて返すチャンク数上限 */
var PDF_CHUNK_BATCH_MAX = 4;

var PDF_CONFIG_SHEET = '設定';
var PDF_SETUP_SHEET = 'セットアップ手順';

/** WebアプリデプロイID（push/deploy後に更新） */
var PDF_WEBAPP_DEPLOY_ID = 'AKfycbxJoWYd4xk6Tycy6mLX0le-2sVwlJP6qID-HRBbMOJ5AI8fvr83HDpUJiKyOq38_g3D0Q';

function getPdfSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートに接続できません');
  return ss;
}

function getPdfRootFolderId_(scope) {
  scope = String(scope || PDF_SCOPE_EXTERNAL).trim();
  var id = PDF_ROOT_FOLDER_IDS[scope];
  if (!id) throw new Error('不明な検索対象: ' + scope);
  return id;
}

function getPdfAppBaseUrl_() {
  if (PDF_WEBAPP_DEPLOY_ID) {
    return 'https://script.google.com/macros/s/' + PDF_WEBAPP_DEPLOY_ID + '/exec';
  }
  try {
    var url = ScriptApp.getService().getUrl();
    if (url) return url;
  } catch (e) { Logger.log(e); }
  return '';
}

function buildPdfContentUrl_(fileId, mode) {
  var base = getPdfAppBaseUrl_();
  if (!base) return '';
  var page = mode === 'download' ? 'download' : 'view';
  return base + (base.indexOf('?') >= 0 ? '&' : '?') + 'page=' + page + '&fileId=' + encodeURIComponent(fileId);
}

function normalizeSearchText_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s\-\/\.]/g, '');
}

function matchesPartial_(target, query) {
  query = normalizeSearchText_(query);
  if (!query) return true;
  target = normalizeSearchText_(target);
  return target.indexOf(query) >= 0;
}

/** 製造日を yyyymmdd 形式に正規化 */
function normalizeMfgDateComparable_(value) {
  var digits = normalizeSearchText_(value).replace(/\D/g, '');
  if (digits.length >= 8) return digits.substring(0, 8);
  return digits;
}

/**
 * 製造日検索条件を解析
 * 例: 20260501～20260530 / 20260501-20260530 / 20260115（部分一致）
 */
function parseMfgDateFilter_(query) {
  query = String(query || '').trim();
  if (!query) return null;

  var compact = query.replace(/[\s\u3000]/g, '');
  var tildeParts = compact.split(/[～~〜]/);
  if (tildeParts.length === 2) {
    return buildMfgDateRange_(tildeParts[0], tildeParts[1]);
  }

  var rangeMatch = compact.match(/^(\d{8})[\-－](\d{8})$/);
  if (rangeMatch) {
    return buildMfgDateRange_(rangeMatch[1], rangeMatch[2]);
  }

  return { partial: normalizeSearchText_(query) };
}

function buildMfgDateRange_(fromStr, toStr) {
  var from = normalizeMfgDateComparable_(fromStr);
  var to = normalizeMfgDateComparable_(toStr);
  if (!from || !to || from.length < 8 || to.length < 8) {
    return { partial: normalizeSearchText_(fromStr + toStr) };
  }
  if (from > to) {
    var tmp = from;
    from = to;
    to = tmp;
  }
  return { from: from, to: to };
}

function matchesMfgDateFilter_(mfgDate, filter) {
  if (!filter) return true;
  var comparable = normalizeMfgDateComparable_(mfgDate);
  if (filter.from && filter.to) {
    if (!comparable || comparable.length < 8) return false;
    return comparable >= filter.from && comparable <= filter.to;
  }
  if (filter.partial) {
    return normalizeSearchText_(mfgDate).indexOf(filter.partial) >= 0;
  }
  return true;
}

/** 範囲検索時に走査する西暦フォルダを絞り込む */
function getMfgDateYearSpan_(filter) {
  if (!filter || !filter.from || !filter.to) return null;
  var yFrom = parseInt(filter.from.substring(0, 4), 10);
  var yTo = parseInt(filter.to.substring(0, 4), 10);
  if (isNaN(yFrom) || isNaN(yTo)) return null;
  var years = {};
  for (var y = yFrom; y <= yTo; y++) years[String(y)] = true;
  return years;
}
