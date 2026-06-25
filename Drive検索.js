// ========================================
// 製品PDF検索 — Drive走査・検索
// ========================================

/**
 * ファイル名解析
 * 例: ABC123-LOT01-20260115.pdf
 * 例: ABC123-LOT01-20260115_original.pdf
 */
function parsePdfFileName_(name) {
  name = String(name || '').trim();
  if (!name) return null;

  var isOriginal = /_original\.pdf$/i.test(name);
  var base = name.replace(/\.pdf$/i, '').replace(/_original$/i, '');
  var parts = base.split('-');
  if (parts.length < 3) return null;

  var productCode = parts[0];
  var lotNo = parts[1];
  var mfgDate = parts.slice(2).join('-');
  var year = /^\d{4}/.test(mfgDate) ? mfgDate.substring(0, 4) : '';

  return {
    fileName: name,
    productCode: productCode,
    lotNo: lotNo,
    mfgDate: mfgDate,
    year: year,
    isOriginal: isOriginal
  };
}

function isPdfScopeMatch_(parsed, scope) {
  if (!parsed) return false;
  if (scope === PDF_SCOPE_INTERNAL) return parsed.isOriginal === true;
  if (scope === PDF_SCOPE_EXTERNAL) return parsed.isOriginal !== true;
  return true;
}

function matchesSearchFilters_(parsed, filters) {
  if (!parsed) return false;
  if (!matchesPartial_(parsed.productCode, filters.productCode)) return false;
  if (!matchesPartial_(parsed.lotNo, filters.lotNo)) return false;
  if (!matchesMfgDateFilter_(parsed.mfgDate, filters.mfgDateFilter)) return false;
  return true;
}

function filterYearFolders_(yearFolders, yearSpan) {
  if (!yearSpan) return yearFolders;
  return yearFolders.filter(function(f) {
    var name = String(f.name || '').trim();
    return yearSpan[name] === true || yearSpan[normalizeSearchText_(name)] === true;
  });
}

function getCachedProductFolders_(rootFolderId) {
  var cacheKey = 'pdf_product_folders_' + rootFolderId;
  try {
    var raw = CacheService.getScriptCache().get(cacheKey);
    if (raw) return JSON.parse(raw);
  } catch (e) { /* ignore */ }

  var root = DriveApp.getFolderById(rootFolderId);
  var list = [];
  var it = root.getFolders();
  while (it.hasNext()) {
    var folder = it.next();
    list.push({ id: folder.getId(), name: folder.getName() });
  }
  list.sort(function(a, b) { return a.name.localeCompare(b.name, 'ja'); });

  try {
    CacheService.getScriptCache().put(
      cacheKey,
      JSON.stringify(list),
      PDF_FOLDER_LIST_CACHE_TTL_SEC
    );
  } catch (e) { Logger.log(e); }

  return list;
}

function filterProductFolders_(folders, productCodeQuery) {
  productCodeQuery = String(productCodeQuery || '').trim();
  if (!productCodeQuery) return folders;
  var q = normalizeSearchText_(productCodeQuery);
  return folders.filter(function(f) {
    return normalizeSearchText_(f.name).indexOf(q) >= 0;
  });
}

function listYearFolders_(productFolderId) {
  var folder = DriveApp.getFolderById(productFolderId);
  var list = [];
  var it = folder.getFolders();
  while (it.hasNext()) {
    var sub = it.next();
    list.push({ id: sub.getId(), name: sub.getName() });
  }
  return list;
}

function collectPdfsFromYearFolder_(yearFolderId, scope, filters, results, state) {
  if (state.truncated) return;

  var folder = DriveApp.getFolderById(yearFolderId);
  var files = folder.getFilesByType(MimeType.PDF);
  while (files.hasNext()) {
    if (state.truncated) return;
    var file = files.next();
    var parsed = parsePdfFileName_(file.getName());
    if (!parsed) continue;
    if (!isPdfScopeMatch_(parsed, scope)) continue;
    if (!matchesSearchFilters_(parsed, filters)) continue;

    results.push({
      fileId: file.getId(),
      fileName: parsed.fileName,
      productCode: parsed.productCode,
      lotNo: parsed.lotNo,
      mfgDate: parsed.mfgDate,
      year: parsed.year || folder.getName(),
      updatedAt: Utilities.formatDate(
        file.getLastUpdated(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm'
      ),
      viewUrl: buildPdfContentUrl_(file.getId(), 'view'),
      downloadUrl: buildPdfContentUrl_(file.getId(), 'download')
    });

    if (results.length >= PDF_SEARCH_MAX_RESULTS) {
      state.truncated = true;
      return;
    }
  }
}

function searchProductPdfs(payload) {
  payload = payload || {};
  var scope = String(payload.scope || PDF_SCOPE_EXTERNAL).trim();
  var rootFolderId = getPdfRootFolderId_(scope);
  var mfgDateFilter = parseMfgDateFilter_(payload.mfgDate || '');
  var yearSpan = getMfgDateYearSpan_(mfgDateFilter);
  var filters = {
    productCode: payload.productCode || '',
    lotNo: payload.lotNo || '',
    mfgDateFilter: mfgDateFilter
  };

  var productFolders = getCachedProductFolders_(rootFolderId);
  productFolders = filterProductFolders_(productFolders, filters.productCode);

  var results = [];
  var state = { truncated: false };

  for (var i = 0; i < productFolders.length; i++) {
    if (state.truncated) break;
    var yearFolders = listYearFolders_(productFolders[i].id);
    yearFolders = filterYearFolders_(yearFolders, yearSpan);
    for (var j = 0; j < yearFolders.length; j++) {
      collectPdfsFromYearFolder_(yearFolders[j].id, scope, filters, results, state);
      if (state.truncated) break;
    }
  }

  results.sort(function(a, b) {
    var da = a.mfgDate || '';
    var db = b.mfgDate || '';
    if (da !== db) return db.localeCompare(da);
    return (b.updatedAt || '').localeCompare(a.updatedAt || '');
  });

  return {
    scope: scope,
    scopeLabel: PDF_SCOPE_LABELS[scope] || scope,
    count: results.length,
    truncated: state.truncated,
    maxResults: PDF_SEARCH_MAX_RESULTS,
    items: results
  };
}

function isFileUnderPdfRoots_(fileId) {
  fileId = String(fileId || '').trim();
  if (!fileId) return false;

  try {
    var file = DriveApp.getFileById(fileId);
    var mime = String(file.getMimeType() || '');
    if (mime !== 'application/pdf' && mime.indexOf('pdf') < 0) return false;

    var folder = file.getParents().hasNext() ? file.getParents().next() : null;
    while (folder) {
      var fid = folder.getId();
      if (fid === PDF_ROOT_FOLDER_IDS.internal || fid === PDF_ROOT_FOLDER_IDS.external) {
        return true;
      }
      var parents = folder.getParents();
      folder = parents.hasNext() ? parents.next() : null;
    }
  } catch (e) {
    Logger.log('isFileUnderPdfRoots_: ' + e.message);
  }
  return false;
}

function testPdfFolderConnection_(scope) {
  scope = scope || PDF_SCOPE_EXTERNAL;
  var rootId = getPdfRootFolderId_(scope);
  var folder = DriveApp.getFolderById(rootId);
  var productFolders = getCachedProductFolders_(rootId);
  return {
    scope: scope,
    scopeLabel: PDF_SCOPE_LABELS[scope],
    rootId: rootId,
    rootName: folder.getName(),
    productFolderCount: productFolders.length,
    sampleProductCodes: productFolders.slice(0, 5).map(function(f) { return f.name; })
  };
}
